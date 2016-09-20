VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmManCampos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Campos - Huertos"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13095
   Icon            =   "frmManCampos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   13095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Cuenta Principal|N|N|0|1|cltebanc|ctaprpal|0||"
   Begin VB.Frame Frame2 
      Height          =   1125
      Index           =   0
      Left            =   240
      TabIndex        =   63
      Top             =   510
      Width           =   12735
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   23
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   1
         Tag             =   "Nro Orden|N|N|0|999999|rcampos|nrocampo|000000||"
         Top             =   660
         Width           =   735
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   22
         Left            =   5400
         MaxLength       =   40
         TabIndex        =   106
         Top             =   675
         Width           =   4995
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   22
         Left            =   4290
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Código Propietario|N|N|1|999999|rcampos|codpropiet|000000|N|"
         Top             =   675
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   4275
         MaxLength       =   6
         TabIndex        =   2
         Tag             =   "Código Socio|N|N|1|999999|rcampos|codsocio|000000|N|"
         Top             =   315
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   5400
         MaxLength       =   40
         TabIndex        =   85
         Top             =   315
         Width           =   4995
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   0
         Tag             =   "Código Campo|N|N|1|99999999|rcampos|codcampo|00000000|S|"
         Top             =   315
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Nº Orden"
         Height          =   255
         Left            =   270
         TabIndex        =   108
         Top             =   660
         Width           =   1185
      End
      Begin VB.Label Label26 
         Caption         =   "Propietario"
         Height          =   255
         Left            =   3015
         TabIndex        =   107
         Top             =   720
         Width           =   915
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   3960
         ToolTipText     =   "Buscar Propietario"
         Top             =   675
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   3960
         ToolTipText     =   "Buscar Socio"
         Top             =   315
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Socio"
         Height          =   255
         Left            =   3015
         TabIndex        =   65
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Campo"
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
         Height          =   255
         Index           =   0
         Left            =   270
         TabIndex        =   64
         Top             =   315
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   195
      TabIndex        =   61
      Top             =   6660
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
         TabIndex        =   62
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   11910
      TabIndex        =   60
      Top             =   6810
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   10800
      TabIndex        =   59
      Top             =   6810
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3960
      Top             =   6930
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
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   26
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
            Object.ToolTipText     =   "Verificación Errores"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Sigpac"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Goolzoom"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Chequeo Nro.Orden"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambio de Socio"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Informe Gastos/Campo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Asignación Globalgap"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Index           =   0
         Left            =   10440
         TabIndex        =   68
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   11910
      TabIndex        =   66
      Top             =   6810
      Visible         =   0   'False
      Width           =   1035
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4890
      Left            =   240
      TabIndex        =   69
      Top             =   1740
      Width           =   12765
      _ExtentX        =   22516
      _ExtentY        =   8625
      _Version        =   393216
      Style           =   1
      Tabs            =   10
      Tab             =   1
      TabsPerRow      =   11
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
      TabPicture(0)   =   "frmManCampos.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(26)"
      Tab(0).Control(1)=   "Label28"
      Tab(0).Control(2)=   "imgZoom(0)"
      Tab(0).Control(3)=   "Label29"
      Tab(0).Control(4)=   "Label6(0)"
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(6)=   "imgBuscar(2)"
      Tab(0).Control(7)=   "imgBuscar(3)"
      Tab(0).Control(8)=   "Label2"
      Tab(0).Control(9)=   "imgBuscar(13)"
      Tab(0).Control(10)=   "Label36"
      Tab(0).Control(11)=   "Frame4"
      Tab(0).Control(12)=   "FrameDatosDtoAdministracion"
      Tab(0).Control(13)=   "Text5(3)"
      Tab(0).Control(14)=   "Text4(3)"
      Tab(0).Control(15)=   "Text1(37)"
      Tab(0).Control(16)=   "Text1(21)"
      Tab(0).Control(17)=   "Text1(3)"
      Tab(0).Control(18)=   "Text1(2)"
      Tab(0).Control(19)=   "Text2(2)"
      Tab(0).Control(20)=   "Text2(3)"
      Tab(0).Control(21)=   "Combo1(0)"
      Tab(0).Control(22)=   "Text5(0)"
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "Otros Datos"
      TabPicture(1)   =   "frmManCampos.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame6"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame7"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame8"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame9"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame10"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Coopropietarios"
      TabPicture(2)   =   "frmManCampos.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameAux0"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Clasificación"
      TabPicture(3)   =   "frmManCampos.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrameAux1"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Parcelas"
      TabPicture(4)   =   "frmManCampos.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "FrameAux2"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Agroseguro"
      TabPicture(5)   =   "frmManCampos.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "FrameAux3"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Histórico"
      TabPicture(6)   =   "frmManCampos.frx":00B4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "FrameAux4"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Control Gastos"
      TabPicture(7)   =   "frmManCampos.frx":00D0
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "FrameAux5"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Ordenes Rec."
      TabPicture(8)   =   "frmManCampos.frx":00EC
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "FrameAux6"
      Tab(8).Control(1)=   "ListView4"
      Tab(8).ControlCount=   2
      TabCaption(9)   =   "Registro Visitas"
      TabPicture(9)   =   "frmManCampos.frx":0108
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "FrameAux7"
      Tab(9).ControlCount=   1
      Begin VB.Frame Frame10 
         Caption         =   "Puntos"
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
         Height          =   675
         Left            =   150
         TabIndex        =   257
         Top             =   3270
         Width           =   6915
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   47
            Left            =   6270
            MaxLength       =   2
            TabIndex        =   47
            Tag             =   "Puntos Estado Vegetativo|N|S|1|4|rcampos|ptosestadovege||#0|"
            Top             =   240
            Width           =   270
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   46
            Left            =   4320
            MaxLength       =   2
            TabIndex        =   46
            Tag             =   "Puntos Calibre|N|S|1|4|rcampos|ptoscalibre||#0|"
            Top             =   240
            Width           =   270
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   45
            Left            =   3090
            MaxLength       =   2
            TabIndex        =   45
            Tag             =   "Puntos Estado Fito|N|S|1|4|rcampos|ptosestadofito||#0|"
            Top             =   240
            Width           =   270
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   44
            Left            =   960
            MaxLength       =   2
            TabIndex        =   44
            Tag             =   "Puntos Calidad|N|S|1|4|rcampos|ptoscalidad||#0|"
            Top             =   240
            Width           =   270
         End
         Begin VB.Label Label49 
            Caption         =   "Calibre"
            Height          =   255
            Left            =   3660
            TabIndex        =   261
            Top             =   270
            Width           =   885
         End
         Begin VB.Label Label48 
            Caption         =   "Estado Vegetativo"
            Height          =   255
            Left            =   4830
            TabIndex        =   260
            Top             =   270
            Width           =   1515
         End
         Begin VB.Label Label47 
            Caption         =   "Calidad"
            Height          =   255
            Left            =   210
            TabIndex        =   259
            Top             =   270
            Width           =   1155
         End
         Begin VB.Label Label46 
            Caption         =   "Estado Fitosanitario"
            Height          =   255
            Left            =   1530
            TabIndex        =   258
            Top             =   270
            Width           =   1515
         End
      End
      Begin VB.Frame FrameAux7 
         BorderStyle     =   0  'None
         Height          =   3910
         Left            =   -74820
         TabIndex        =   248
         Top             =   450
         Width           =   12450
         Begin VB.TextBox txtAux9 
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   4
            Left            =   5700
            TabIndex        =   254
            Tag             =   "Observaciones|T|S|||rcampos_revision|observac|||"
            Text            =   "Observac"
            Top             =   2910
            Visible         =   0   'False
            Width           =   5385
         End
         Begin VB.TextBox txtAux9 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   2
            Left            =   1710
            MaxLength       =   10
            TabIndex        =   252
            Tag             =   "Fecha|F|N|||rcampos_revision|fecha|dd/mm/yyyy||"
            Text            =   "Fecha"
            Top             =   2910
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtAux9 
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   3
            Left            =   2850
            MaxLength       =   50
            TabIndex        =   253
            Tag             =   "Tecnico|T|S|||rcampos_revision|tecnico|||"
            Text            =   "tecnico"
            Top             =   2910
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.TextBox txtAux9 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   1
            Left            =   945
            MaxLength       =   6
            TabIndex        =   251
            Tag             =   "Linea|N|N|||rcampos_revision|numlinea|000|S|"
            Text            =   "linea"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtAux9 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   0
            Left            =   225
            MaxLength       =   8
            TabIndex        =   250
            Tag             =   "Campo|N|N|0|99999999|rcampos_revision|codcampo|00000000|S|"
            Text            =   "campo"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   315
            Index           =   12
            Left            =   2520
            TabIndex        =   249
            ToolTipText     =   "Buscar fecha"
            Top             =   2880
            Visible         =   0   'False
            Width           =   195
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   7
            Left            =   45
            TabIndex        =   255
            Top             =   0
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   688
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   4
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
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Impresión"
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc AdoAux 
            Height          =   375
            Index           =   7
            Left            =   5280
            Top             =   210
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
            Bindings        =   "frmManCampos.frx":0124
            Height          =   3225
            Index           =   7
            Left            =   45
            TabIndex        =   256
            Top             =   450
            Width           =   12300
            _ExtentX        =   21696
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
         Begin VB.Image imgZoom 
            Height          =   240
            Index           =   1
            Left            =   11580
            Tag             =   "-1"
            ToolTipText     =   "Zoom Observaciones"
            Top             =   150
            Width           =   225
         End
      End
      Begin VB.Frame Frame9 
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
         Height          =   675
         Left            =   150
         TabIndex        =   223
         Top             =   3960
         Width           =   6915
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   43
            Left            =   5310
            MaxLength       =   10
            TabIndex        =   49
            Tag             =   "Fecha Alta Programa Operativo|F|S|||rcampos|fecaltapropera|dd/mm/yyyy||"
            Top             =   240
            Width           =   1200
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   40
            Left            =   1530
            MaxLength       =   4
            TabIndex        =   48
            Tag             =   "%Comision sobre precio|N|S|||rcampos|dtoprecio|##0.00||"
            Top             =   240
            Width           =   675
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   3
            Left            =   5010
            Picture         =   "frmManCampos.frx":013C
            ToolTipText     =   "Buscar fecha"
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label45 
            Caption         =   "Fecha Alta Programa Operativo"
            Height          =   255
            Left            =   2640
            TabIndex        =   247
            Top             =   270
            Width           =   2235
         End
         Begin VB.Label Label40 
            Caption         =   "% Comisión"
            Height          =   255
            Left            =   210
            TabIndex        =   224
            Top             =   270
            Width           =   1155
         End
      End
      Begin VB.Frame FrameAux5 
         BorderStyle     =   0  'None
         Height          =   3910
         Left            =   -74820
         TabIndex        =   211
         Top             =   450
         Width           =   12210
         Begin VB.CheckBox chkAux 
            BackColor       =   &H80000005&
            Height          =   255
            Index           =   1
            Left            =   7830
            TabIndex        =   222
            Tag             =   "Contabilizado|N|N|0|1|rcampos_gastos|contabilizado|||"
            Top             =   2970
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   315
            Index           =   10
            Left            =   6750
            TabIndex        =   219
            ToolTipText     =   "Buscar fecha"
            Top             =   2940
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtAux7 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   0
            Left            =   225
            MaxLength       =   8
            TabIndex        =   217
            Tag             =   "Campo|N|N|0|99999999|rcampos_gastos|codcampo|00000000|S|"
            Text            =   "campo"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtAux7 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   1
            Left            =   945
            MaxLength       =   6
            TabIndex        =   215
            Tag             =   "Linea|N|N|||rcampos_gastos|numlinea|000|S|"
            Text            =   "linea"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtAux7 
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   2
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   214
            Tag             =   "Concepto|N|S|||rcampos_gastos|codgasto|00||"
            Text            =   "Co"
            Top             =   2940
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtAux7 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   3
            Left            =   6000
            MaxLength       =   10
            TabIndex        =   216
            Tag             =   "Fecha|F|N|||rcampos_gastos|fecha|dd/mm/yyyy||"
            Text            =   "Fecha"
            Top             =   2970
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtAux7 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   4
            Left            =   6960
            MaxLength       =   10
            TabIndex        =   218
            Tag             =   "Importe|N|N|||rcampos_gastos|importe|###,###,##0.00||"
            Text            =   "Importe"
            Top             =   2970
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   315
            Index           =   9
            Left            =   2460
            TabIndex        =   213
            ToolTipText     =   "Buscar concepto gasto"
            Top             =   2910
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtAux2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   5
            Left            =   2700
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   212
            Text            =   "Nombre concepto"
            Top             =   2940
            Visible         =   0   'False
            Width           =   3285
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   5
            Left            =   45
            TabIndex        =   220
            Top             =   0
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   688
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   4
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
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Integracion Contable"
                  Object.Tag             =   "2"
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc AdoAux 
            Height          =   375
            Index           =   5
            Left            =   5280
            Top             =   210
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
            Bindings        =   "frmManCampos.frx":01C7
            Height          =   3225
            Index           =   5
            Left            =   45
            TabIndex        =   221
            Top             =   450
            Width           =   9120
            _ExtentX        =   16087
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
      Begin VB.Frame Frame8 
         Caption         =   "Cliente Tienda"
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
         Height          =   675
         Left            =   7260
         TabIndex        =   207
         Top             =   3960
         Width           =   5265
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   39
            Left            =   1170
            MaxLength       =   4
            TabIndex        =   209
            Tag             =   "Codigo Cliente|N|S|||rcampos|codclien|||"
            Top             =   240
            Width           =   675
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   39
            Left            =   1890
            TabIndex        =   208
            Top             =   240
            Width           =   3180
         End
         Begin VB.Label Label39 
            Caption         =   "Código"
            Height          =   255
            Left            =   330
            TabIndex        =   210
            Top             =   270
            Width           =   525
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   15
            Left            =   870
            ToolTipText     =   "Buscar globalgap"
            Top             =   270
            Width           =   240
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Globalgap"
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
         Height          =   675
         Left            =   7260
         TabIndex        =   204
         Top             =   3270
         Width           =   5265
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   38
            Left            =   1890
            TabIndex        =   206
            Top             =   240
            Width           =   3180
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   38
            Left            =   1170
            MaxLength       =   4
            TabIndex        =   58
            Tag             =   "Codigo GlobalGap|T|S|||rcampos|codigoggap|||"
            Top             =   240
            Width           =   675
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   14
            Left            =   870
            ToolTipText     =   "Buscar globalgap"
            Top             =   270
            Width           =   240
         End
         Begin VB.Label Label38 
            Caption         =   "Código"
            Height          =   255
            Left            =   330
            TabIndex        =   205
            Top             =   270
            Width           =   555
         End
      End
      Begin VB.Frame FrameAux4 
         BorderStyle     =   0  'None
         Height          =   3930
         Left            =   -74820
         TabIndex        =   188
         Top             =   450
         Width           =   12210
         Begin VB.TextBox txtAux2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   4
            Left            =   2700
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   202
            Text            =   "Nombre socio"
            Top             =   2940
            Visible         =   0   'False
            Width           =   3285
         End
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   315
            Index           =   8
            Left            =   2460
            TabIndex        =   201
            ToolTipText     =   "Buscar socio"
            Top             =   2910
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtaux6 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   4
            Left            =   6960
            MaxLength       =   10
            TabIndex        =   197
            Tag             =   "Fecha Baja|F|S|||rcampos_hco|fechabaja|dd/mm/yyyy||"
            Text            =   "Fec.Baja"
            Top             =   2940
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtaux6 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   3
            Left            =   6000
            MaxLength       =   10
            TabIndex        =   196
            Tag             =   "Fecha Alta|F|N|||rcampos_hco|fechaalta|dd/mm/yyyy||"
            Text            =   "Fec.Alta"
            Top             =   2940
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtaux6 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   2
            Left            =   1680
            MaxLength       =   7
            TabIndex        =   195
            Tag             =   "Socio|N|S|||rcampos_hco|codsocio|000000||"
            Text            =   "Socio"
            Top             =   2940
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtaux6 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   5
            Left            =   7980
            MaxLength       =   4
            TabIndex        =   198
            Tag             =   "Incidencia|N|S|||rcampos_hco|codincid|0000||"
            Text            =   "In"
            Top             =   2940
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox txtaux6 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   1
            Left            =   945
            MaxLength       =   6
            TabIndex        =   194
            Tag             =   "Linea|N|N|||rcampos_hco|numlinea|000|S|"
            Text            =   "linea"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux6 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   0
            Left            =   225
            MaxLength       =   8
            TabIndex        =   193
            Tag             =   "Campo|N|N|0|99999999|rcampos_hco|codcampo|00000000|S|"
            Text            =   "campo"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtAux2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   3
            Left            =   8820
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   192
            Text            =   "Nombre incidencia"
            Top             =   2940
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   315
            Index           =   7
            Left            =   8580
            TabIndex        =   191
            ToolTipText     =   "Buscar incidencia"
            Top             =   2940
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   315
            Index           =   6
            Left            =   6750
            TabIndex        =   190
            ToolTipText     =   "Buscar fecha"
            Top             =   2940
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   315
            Index           =   5
            Left            =   7740
            TabIndex        =   189
            ToolTipText     =   "Buscar fecha"
            Top             =   2940
            Visible         =   0   'False
            Width           =   195
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   4
            Left            =   45
            TabIndex        =   199
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
            Index           =   4
            Left            =   5280
            Top             =   210
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
            Bindings        =   "frmManCampos.frx":01DF
            Height          =   3225
            Index           =   4
            Left            =   45
            TabIndex        =   200
            Top             =   450
            Width           =   10710
            _ExtentX        =   18891
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
      Begin VB.TextBox Text5 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -70950
         MaxLength       =   30
         TabIndex        =   186
         Top             =   1920
         Width           =   1530
      End
      Begin VB.Frame FrameAux3 
         BorderStyle     =   0  'None
         Height          =   3930
         Left            =   -74820
         TabIndex        =   170
         Top             =   450
         Width           =   12210
         Begin VB.TextBox txtaux5 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   7
            Left            =   7800
            MaxLength       =   7
            TabIndex        =   176
            Tag             =   "Kilos Aportacion|N|S|||rcampos_seguros|kilosaportacion|###,##0||"
            Text            =   "Kilos A"
            Top             =   2970
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   315
            Index           =   4
            Left            =   10170
            TabIndex        =   184
            ToolTipText     =   "Buscar fecha"
            Top             =   2970
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.CheckBox chkAux 
            BackColor       =   &H80000005&
            Height          =   255
            Index           =   0
            Left            =   10470
            TabIndex        =   179
            Tag             =   "Siniestro|N|N|0|1|rcampos_seguros|essiniestro|||"
            Top             =   3000
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   315
            Index           =   3
            Left            =   2430
            TabIndex        =   183
            ToolTipText     =   "Buscar fecha"
            Top             =   2940
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   315
            Index           =   2
            Left            =   3390
            TabIndex        =   182
            ToolTipText     =   "Buscar incidencia"
            Top             =   2910
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtAux2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   2
            Left            =   3600
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   185
            Text            =   "Nombre incidencia"
            Top             =   2940
            Visible         =   0   'False
            Width           =   3285
         End
         Begin VB.TextBox txtaux5 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   0
            Left            =   225
            MaxLength       =   8
            TabIndex        =   171
            Tag             =   "Campo|N|N|0|99999999|rcampos_seguros|codcampo|00000000|S|"
            Text            =   "campo"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux5 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   1
            Left            =   945
            MaxLength       =   6
            TabIndex        =   172
            Tag             =   "Linea|N|N|||rcampos_seguros|numlinea|000|S|"
            Text            =   "linea"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux5 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   3
            Left            =   2640
            MaxLength       =   4
            TabIndex        =   174
            Tag             =   "Incidencia|N|S|||rcampos_seguros|codincid|0000||"
            Text            =   "In"
            Top             =   2940
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtaux5 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   4
            Left            =   6990
            MaxLength       =   7
            TabIndex        =   175
            Tag             =   "Kilos Indemniz.|N|S|||rcampos_seguros|kilos|###,##0||"
            Text            =   "Kilos"
            Top             =   2970
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtaux5 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   2
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   173
            Tag             =   "Fecha|F|N|||rcampos_seguros|fecha|dd/mm/yyyy||"
            Text            =   "Fec"
            Top             =   2940
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtaux5 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   5
            Left            =   8580
            MaxLength       =   13
            TabIndex        =   177
            Tag             =   "Importe|N|S|||rcampos_seguros|importe|##,###,##0.00||"
            Text            =   "Importe"
            Top             =   2970
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtaux5 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   6
            Left            =   9390
            MaxLength       =   10
            TabIndex        =   178
            Tag             =   "Fecha Pago|F|S|||rcampos_seguros|fechapago|dd/mm/yyyy||"
            Text            =   "Fec.Pago"
            Top             =   2970
            Visible         =   0   'False
            Width           =   735
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   3
            Left            =   45
            TabIndex        =   180
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
            Left            =   5280
            Top             =   210
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
            Bindings        =   "frmManCampos.frx":01F7
            Height          =   3225
            Index           =   3
            Left            =   45
            TabIndex        =   181
            Top             =   450
            Width           =   12120
            _ExtentX        =   21378
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
      Begin VB.Frame Frame6 
         Caption         =   "Datos Seguros Campaña Anterior"
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
         Height          =   1305
         Left            =   7260
         TabIndex        =   165
         Top             =   1950
         Width           =   5235
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   36
            Left            =   1650
            MaxLength       =   6
            TabIndex        =   55
            Tag             =   "Kg.Seguro|N|S|0|999999|rcampos|kilosaseant|###,###||"
            Top             =   450
            Width           =   1305
         End
         Begin VB.CheckBox chkAbonos 
            Caption         =   "Seguro"
            Height          =   315
            Index           =   5
            Left            =   330
            TabIndex        =   54
            Tag             =   "Seguro|N|N|||rcampos|aseguradoant||N|"
            Top             =   240
            Width           =   915
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   35
            Left            =   3420
            MaxLength       =   13
            TabIndex        =   56
            Tag             =   "Importe.Seguro|N|S|||rcampos|costeseguroant|##,###,##0.00||"
            Top             =   450
            Width           =   1275
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   34
            Left            =   360
            MaxLength       =   2
            TabIndex        =   57
            Tag             =   "Seguro Opcion|T|S|||rcampos|codseguroant|||"
            Top             =   870
            Width           =   735
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   34
            Left            =   1140
            TabIndex        =   166
            Top             =   870
            Width           =   3540
         End
         Begin VB.Label Label1 
            Caption         =   "Kg.seguro"
            Height          =   255
            Index           =   4
            Left            =   1650
            TabIndex        =   169
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Coste seguro"
            Height          =   255
            Index           =   3
            Left            =   3420
            TabIndex        =   168
            Top             =   180
            Width           =   735
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   12
            Left            =   960
            ToolTipText     =   "Buscar Desarrollo Vegetativo"
            Top             =   570
            Width           =   240
         End
         Begin VB.Label Label35 
            Caption         =   "Opción"
            Height          =   255
            Left            =   360
            TabIndex        =   167
            Top             =   600
            Width           =   555
         End
      End
      Begin VB.Frame FrameAux2 
         BorderStyle     =   0  'None
         Height          =   4200
         Left            =   -74820
         TabIndex        =   149
         Top             =   450
         Width           =   12210
         Begin VB.TextBox txtAux2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   13
            Left            =   10080
            TabIndex        =   236
            Top             =   30
            Width           =   1600
         End
         Begin VB.TextBox txtAux2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   10080
            TabIndex        =   230
            Top             =   3720
            Width           =   1600
         End
         Begin VB.TextBox txtAux2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   8460
            TabIndex        =   228
            Top             =   3720
            Width           =   1600
         End
         Begin VB.TextBox txtAux2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   5220
            TabIndex        =   227
            Top             =   3720
            Width           =   1600
         End
         Begin VB.TextBox txtAux2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   6840
            TabIndex        =   226
            Top             =   3720
            Width           =   1600
         End
         Begin VB.TextBox txtaux4 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   10
            Left            =   9990
            MaxLength       =   10
            TabIndex        =   162
            Tag             =   "Sup.Cult Catas|N|N|0|9999.9999|rcampos_parcelas|supcultcatas|###0.0000||"
            Text            =   "Sup.Cult C"
            Top             =   2970
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txtaux4 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   9
            Left            =   8400
            MaxLength       =   10
            TabIndex        =   161
            Tag             =   "Sup.Catas|N|N|0|9999.9999|rcampos_parcelas|supcatas|###0.0000||"
            Text            =   "Sup.Catas"
            Top             =   2970
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txtaux4 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   8
            Left            =   7260
            MaxLength       =   10
            TabIndex        =   160
            Tag             =   "Sup.Cult Sigpac|N|N|0|9999.9999|rcampos_parcelas|supcultsigpa|###0.0000||"
            Text            =   "Sup.Sigpa"
            Top             =   2970
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtaux4 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   7
            Left            =   6060
            MaxLength       =   10
            TabIndex        =   159
            Tag             =   "Sup.Sigpac|N|N|0|9999.9999|rcampos_parcelas|supsigpa|###0.0000||"
            Text            =   "Sup.Sigpa"
            Top             =   2970
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtaux4 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   6
            Left            =   5220
            MaxLength       =   6
            TabIndex        =   158
            Tag             =   "Cod.SigpaC|N|S|0|999999|rcampos_parcelas|codsigpa|000000||"
            Text            =   "sig"
            Top             =   2970
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtaux4 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   5
            Left            =   4350
            MaxLength       =   3
            TabIndex        =   157
            Tag             =   "Recinto|N|N|0|999|rcampos_parcelas|recintos|000||"
            Text            =   "rec"
            Top             =   2970
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtaux4 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   2
            Left            =   1680
            MaxLength       =   3
            TabIndex        =   154
            Tag             =   "Poligono|N|N|0|999|rcampos_parcelas|poligono|000||"
            Text            =   "pol"
            Top             =   2940
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtaux4 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   3
            Left            =   2580
            MaxLength       =   6
            TabIndex        =   155
            Tag             =   "Parcela|N|N|0|999999|rcampos_parcelas|parcela|000000||"
            Text            =   "par"
            Top             =   2940
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtaux4 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   4
            Left            =   3420
            MaxLength       =   2
            TabIndex        =   156
            Tag             =   "Subparcela|T|S|||rcampos_parcelas|subparce|||"
            Text            =   "su"
            Top             =   2940
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtaux4 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   1
            Left            =   945
            MaxLength       =   6
            TabIndex        =   151
            Tag             =   "Linea|N|N|||rcampos_parcelas|numlinea|000|S|"
            Text            =   "linea"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux4 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   0
            Left            =   225
            MaxLength       =   8
            TabIndex        =   150
            Tag             =   "Campo|N|N|0|99999999|rcampos_parcelas|codcampo|00000000|S|"
            Text            =   "campo"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   2
            Left            =   45
            TabIndex        =   152
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
            Left            =   3960
            Top             =   630
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
            Bindings        =   "frmManCampos.frx":020F
            Height          =   3195
            Index           =   2
            Left            =   45
            TabIndex        =   153
            Top             =   450
            Width           =   11910
            _ExtentX        =   21008
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
         Begin VB.Label Label1 
            Caption         =   "Código Conselleria:  "
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   6
            Left            =   8520
            TabIndex        =   237
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "TOTALES:  "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   5
            Left            =   3930
            TabIndex        =   229
            Top             =   3780
            Width           =   945
         End
      End
      Begin VB.Frame FrameAux1 
         BorderStyle     =   0  'None
         Height          =   3930
         Left            =   -74820
         TabIndex        =   139
         Top             =   450
         Width           =   12210
         Begin VB.TextBox txtaux1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   3
            Left            =   5940
            MaxLength       =   6
            TabIndex        =   145
            Tag             =   "Muestra|N|N|0|100|rcampos_clasif|muestra|##0.00||"
            Text            =   "muestr"
            Top             =   2970
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   2
            Left            =   945
            MaxLength       =   6
            TabIndex        =   144
            Tag             =   "Variedad|N|N|||rcampos_clasif|codvarie|000000|N|"
            Text            =   "var"
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
            MaxLength       =   8
            TabIndex        =   143
            Tag             =   "Campo|N|N|1|99999999|rcampos_clasif|codcampo|00000000|S|"
            Text            =   "campo"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   1
            Left            =   1710
            MaxLength       =   3
            TabIndex        =   142
            Tag             =   "Calidad|N|N|||rcampos_clasif|codcalid|00|S|"
            Text            =   "cal"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   315
            Index           =   1
            Left            =   2385
            TabIndex        =   141
            ToolTipText     =   "Buscar calidad"
            Top             =   2925
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtAux2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   1
            Left            =   2565
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   140
            Text            =   "Nombre calidad"
            Top             =   2925
            Visible         =   0   'False
            Width           =   3285
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   1
            Left            =   45
            TabIndex        =   146
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
            Bindings        =   "frmManCampos.frx":0227
            Height          =   3195
            Index           =   1
            Left            =   45
            TabIndex        =   147
            Top             =   450
            Width           =   6450
            _ExtentX        =   11377
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
         Begin MSChart20Lib.MSChart MSChart1 
            Height          =   3300
            Left            =   6750
            OleObjectBlob   =   "frmManCampos.frx":023F
            TabIndex        =   148
            Top             =   450
            Width           =   5370
         End
      End
      Begin VB.Frame FrameAux0 
         BorderStyle     =   0  'None
         Height          =   4230
         Left            =   -74820
         TabIndex        =   130
         Top             =   450
         Width           =   12210
         Begin VB.TextBox txtAux2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   0
            Left            =   2565
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   136
            Text            =   "Nombre socio"
            Top             =   2925
            Visible         =   0   'False
            Width           =   3285
         End
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   315
            Index           =   0
            Left            =   2385
            TabIndex        =   135
            ToolTipText     =   "Buscar socio"
            Top             =   2925
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtaux3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   2
            Left            =   1710
            MaxLength       =   6
            TabIndex        =   133
            Tag             =   "Socio|N|N|||rcampos_cooprop|codsocio|000000|S|"
            Text            =   "socio"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   0
            Left            =   225
            MaxLength       =   8
            TabIndex        =   132
            Tag             =   "Campo|N|N|1|99999999|rcampos_clasif|codcampo|00000000|S|"
            Text            =   "campo"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   1
            Left            =   945
            MaxLength       =   6
            TabIndex        =   131
            Tag             =   "Linea|N|N|||rcampos_cooprop|numlinea|000|S|"
            Text            =   "linea"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   3
            Left            =   5940
            MaxLength       =   6
            TabIndex        =   134
            Tag             =   "Porcentaje|N|N|0|100|rcampos_cooprop|porcentaje|##0.00||"
            Text            =   "porc"
            Top             =   2940
            Visible         =   0   'False
            Width           =   645
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   0
            Left            =   45
            TabIndex        =   137
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
            Left            =   5280
            Top             =   210
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
            Bindings        =   "frmManCampos.frx":5968
            Height          =   3195
            Index           =   0
            Left            =   45
            TabIndex        =   138
            Top             =   450
            Width           =   7080
            _ExtentX        =   12488
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
      End
      Begin VB.Frame Frame5 
         Caption         =   "Datos Técnicos"
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
         Height          =   2775
         Left            =   150
         TabIndex        =   117
         Top             =   510
         Width           =   6915
         Begin VB.CheckBox chkAbonos 
            Caption         =   "Acabado Recol."
            Height          =   315
            Index           =   7
            Left            =   5100
            TabIndex        =   38
            Tag             =   "Acabado Recol.|N|N|||rcampos|acabadorecol||N|"
            Top             =   1350
            Width           =   1515
         End
         Begin VB.CheckBox chkAbonos 
            Caption         =   "Naturane"
            Height          =   315
            Index           =   6
            Left            =   3390
            TabIndex        =   37
            Tag             =   "Naturane|N|N|||rcampos|esnaturane||N|"
            Top             =   1350
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   32
            Left            =   1530
            MaxLength       =   10
            TabIndex        =   41
            Tag             =   "Nro LLave|N|S|||rcampos|nrollave|#########0||"
            Top             =   2400
            Width           =   1395
         End
         Begin VB.CheckBox chkAbonos 
            Caption         =   "Residuos"
            Height          =   315
            Index           =   4
            Left            =   5100
            TabIndex        =   43
            Tag             =   "Residuos|N|N|||rcampos|conresiduos||N|"
            Top             =   2400
            Width           =   1545
         End
         Begin VB.CheckBox chkAbonos 
            Caption         =   "Poda"
            Height          =   315
            Index           =   3
            Left            =   3390
            TabIndex        =   42
            Tag             =   "Con Poda|N|N|||rcampos|conpoda||N|"
            Top             =   2400
            Width           =   1665
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   31
            Left            =   2340
            MaxLength       =   30
            TabIndex        =   127
            Top             =   2070
            Width           =   4185
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   31
            Left            =   1530
            MaxLength       =   2
            TabIndex        =   40
            Tag             =   "Patrón a Pie|N|S|0|99|rcampos|codpatron|00||"
            Top             =   2070
            Width           =   765
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   30
            Left            =   2340
            MaxLength       =   30
            TabIndex        =   125
            Top             =   1740
            Width           =   4185
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   30
            Left            =   1530
            MaxLength       =   2
            TabIndex        =   39
            Tag             =   "Procedencia Riego|N|S|0|99|rcampos|codproce|00||"
            Top             =   1740
            Width           =   735
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   25
            Left            =   2340
            TabIndex        =   120
            Top             =   270
            Width           =   4200
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   25
            Left            =   1530
            MaxLength       =   2
            TabIndex        =   33
            Tag             =   "Marco Plantación|N|S|0|99|rcampos|codplanta|00||"
            Top             =   270
            Width           =   735
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   26
            Left            =   2340
            TabIndex        =   119
            Top             =   630
            Width           =   4200
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   26
            Left            =   1530
            MaxLength       =   2
            TabIndex        =   34
            Tag             =   "Código Desarrollo|N|S|0|99|rcampos|coddesa|00||"
            Top             =   630
            Width           =   735
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   2
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Tag             =   "Sistema Riego|N|N|||rcampos|codriego||N|"
            Top             =   1350
            Width           =   1440
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   27
            Left            =   2340
            MaxLength       =   30
            TabIndex        =   118
            Top             =   1050
            Width           =   4185
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   27
            Left            =   1530
            MaxLength       =   2
            TabIndex        =   35
            Tag             =   "Tipo Tierra|N|S|0|99|rcampos|codtierra|00||"
            Top             =   990
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Nro LLave"
            Height          =   255
            Index           =   4
            Left            =   210
            TabIndex        =   129
            Top             =   2430
            Width           =   945
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   11
            Left            =   1230
            ToolTipText     =   "Buscar Patrón Pie"
            Top             =   2070
            Width           =   240
         End
         Begin VB.Label Label6 
            Caption         =   "Patrón Pie"
            Height          =   255
            Index           =   3
            Left            =   210
            TabIndex        =   128
            Top             =   2100
            Width           =   945
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   10
            Left            =   1230
            ToolTipText     =   "Buscar Procedencia"
            Top             =   1740
            Width           =   240
         End
         Begin VB.Label Label6 
            Caption         =   "Procedencia"
            Height          =   255
            Index           =   2
            Left            =   210
            TabIndex        =   126
            Top             =   1770
            Width           =   945
         End
         Begin VB.Label Label27 
            Caption         =   "Marco Plant."
            Height          =   255
            Left            =   210
            TabIndex        =   124
            Top             =   300
            Width           =   975
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   6
            Left            =   1230
            ToolTipText     =   "Buscar Marco Plantación"
            Top             =   270
            Width           =   240
         End
         Begin VB.Label Label30 
            Caption         =   "Desarrollo"
            Height          =   255
            Left            =   210
            TabIndex        =   123
            Top             =   630
            Width           =   885
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   7
            Left            =   1230
            ToolTipText     =   "Buscar Tipo Tierra"
            Top             =   990
            Width           =   240
         End
         Begin VB.Label Label12 
            Caption         =   "Sistema Riego"
            Height          =   255
            Left            =   210
            TabIndex        =   122
            Top             =   1380
            Width           =   1035
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   8
            Left            =   1230
            ToolTipText     =   "Buscar Desarrollo Vegetativo"
            Top             =   630
            Width           =   240
         End
         Begin VB.Label Label6 
            Caption         =   "Tipo Tierra"
            Height          =   255
            Index           =   1
            Left            =   210
            TabIndex        =   121
            Top             =   990
            Width           =   945
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Datos Seguros"
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
         Left            =   7260
         TabIndex        =   112
         Top             =   510
         Width           =   5235
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   29
            Left            =   1140
            TabIndex        =   115
            Top             =   900
            Width           =   3540
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   29
            Left            =   360
            MaxLength       =   2
            TabIndex        =   53
            Tag             =   "Seguro Opcion|T|S|||rcampos|codseguro|||"
            Top             =   900
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   28
            Left            =   3420
            MaxLength       =   13
            TabIndex        =   52
            Tag             =   "Importe.Seguro|N|S|||rcampos|costeseguro|##,###,##0.00||"
            Top             =   480
            Width           =   1275
         End
         Begin VB.CheckBox chkAbonos 
            Caption         =   "Seguro"
            Height          =   315
            Index           =   0
            Left            =   330
            TabIndex        =   50
            Tag             =   "Seguro|N|N|||rcampos|asegurado||N|"
            Top             =   270
            Width           =   915
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   1650
            MaxLength       =   6
            TabIndex        =   51
            Tag             =   "Kg.Seguro|N|S|0|999999|rcampos|kilosase|###,###||"
            Top             =   480
            Width           =   1305
         End
         Begin VB.Label Label32 
            Caption         =   "Opción"
            Height          =   255
            Left            =   360
            TabIndex        =   116
            Top             =   630
            Width           =   555
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   9
            Left            =   960
            ToolTipText     =   "Buscar Desarrollo Vegetativo"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Coste seguro"
            Height          =   255
            Index           =   2
            Left            =   3420
            TabIndex        =   114
            Top             =   210
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Kg.seguro"
            Height          =   255
            Index           =   1
            Left            =   1650
            TabIndex        =   113
            Top             =   210
            Width           =   735
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   -73620
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Tag             =   "Tipo Parcela|N|N|0|1|rcampos|tipoparc||N|"
         Top             =   1890
         Width           =   1440
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -72735
         MaxLength       =   30
         TabIndex        =   87
         Top             =   855
         Width           =   3315
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -72735
         MaxLength       =   30
         TabIndex        =   86
         Top             =   520
         Width           =   3315
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   -73620
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "Variedad|N|N|1|9999|rcampos|codvarie|0000||"
         Top             =   520
         Width           =   840
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   -73620
         MaxLength       =   4
         TabIndex        =   5
         Tag             =   "Partida|N|N|1|9999|rcampos|codparti|0000||"
         Top             =   855
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   435
         Index           =   21
         Left            =   -69060
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Tag             =   "Observaciones|T|S|||rcampos|observac|||"
         Top             =   4320
         Width           =   6435
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   37
         Left            =   -73620
         MaxLength       =   6
         TabIndex        =   6
         Tag             =   "Zona|N|N|1|9999|rcampos|codzonas|0000||"
         Top             =   1185
         Width           =   855
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -72735
         MaxLength       =   30
         TabIndex        =   79
         Top             =   1185
         Width           =   3315
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -73620
         MaxLength       =   30
         TabIndex        =   78
         Top             =   1530
         Width           =   4200
      End
      Begin VB.Frame FrameDatosDtoAdministracion 
         Caption         =   "Datos Administrativos y Geográficos"
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
         Height          =   3615
         Left            =   -69180
         TabIndex        =   75
         Top             =   450
         Width           =   6720
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   42
            Left            =   5220
            MaxLength       =   10
            TabIndex        =   24
            Tag             =   "Fecha Revisión|F|S|||rcampos|fecrevision|dd/mm/yyyy||"
            Top             =   2640
            Width           =   1245
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   41
            Left            =   1470
            MaxLength       =   20
            TabIndex        =   23
            Tag             =   "Referencia Catastral|T|S|||rcampos|refercatas|||"
            Text            =   "12345678901234567890"
            Top             =   2610
            Width           =   2055
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   4
            Left            =   4710
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Tag             =   "Entrega Ficha Cultivo|N|N|||rcampos|entregafichaculti||N|"
            Top             =   1860
            Width           =   1770
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   3
            Left            =   4710
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Tag             =   "Tipo Campo|N|N|0|1|rcampos|tipocampo||N|"
            Top             =   1470
            Width           =   1770
         End
         Begin VB.CheckBox chkAbonos 
            Caption         =   "Presentado Pago Único"
            Height          =   315
            Index           =   2
            Left            =   3720
            TabIndex        =   22
            Tag             =   "Presenta Pago Unico|N|N|0|1|rcampos|pagounico||N|"
            Top             =   2220
            Width           =   2445
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   1
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Tag             =   "Recolectado|N|N|0|1|rcampos|recolect||N|"
            Top             =   1500
            Width           =   1440
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   24
            Left            =   1470
            MaxLength       =   4
            TabIndex        =   16
            Tag             =   "Código Responsable|N|S|0|9999|rcampos|codcapat|0000||"
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   24
            Left            =   2250
            TabIndex        =   109
            Top             =   1080
            Width           =   4200
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   20
            Left            =   5370
            MaxLength       =   9
            TabIndex        =   30
            Tag             =   "Longitud|N|S|0|99.999999|rcampos|longitud|#0.000000||"
            Top             =   3210
            Width           =   1110
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   19
            Left            =   4140
            MaxLength       =   9
            TabIndex        =   29
            Tag             =   "Latitud|N|S|0|99.999999|rcampos|latitud|#0.000000||"
            Top             =   3210
            Width           =   1200
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   18
            Left            =   2910
            MaxLength       =   10
            TabIndex        =   28
            Tag             =   "OID|N|S|0|9999999999|rcampos|numeroid|########||"
            Top             =   3210
            Width           =   1200
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   17
            Left            =   1470
            MaxLength       =   4
            TabIndex        =   19
            Tag             =   "Año Plantacion|N|S|0|2100|rcampos|anoplant|0000||"
            Top             =   1890
            Width           =   1200
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   16
            Left            =   1470
            MaxLength       =   3
            TabIndex        =   21
            Tag             =   "Recinto|N|N|0|999|rcampos|recintos|000||"
            Top             =   2250
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   15
            Left            =   1830
            MaxLength       =   2
            TabIndex        =   27
            Tag             =   "Subparcela|T|S|||rcampos|subparce|||"
            Top             =   3210
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   14
            Left            =   990
            MaxLength       =   6
            TabIndex        =   26
            Tag             =   "Parcela|N|N|0|999999|rcampos|parcela|000000||"
            Top             =   3210
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   13
            Left            =   150
            MaxLength       =   3
            TabIndex        =   25
            Tag             =   "Poligono|N|N|0|999|rcampos|poligono|000||"
            Top             =   3210
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   10
            Left            =   1470
            MaxLength       =   10
            TabIndex        =   13
            Tag             =   "Fecha Alta|F|N|||rcampos|fecaltas|dd/mm/yyyy||"
            Top             =   360
            Width           =   1200
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   11
            Left            =   5190
            MaxLength       =   10
            TabIndex        =   14
            Tag             =   "Fecha Baja|F|S|||rcampos|fecbajas|dd/mm/yyyy||"
            Top             =   360
            Width           =   1245
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   12
            Left            =   2250
            TabIndex        =   76
            Top             =   720
            Width           =   4200
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   12
            Left            =   1470
            MaxLength       =   2
            TabIndex        =   15
            Tag             =   "Código Situacion|N|N|0|99|rcampos|codsitua|00||"
            Top             =   720
            Width           =   765
         End
         Begin VB.CheckBox chkAbonos 
            Caption         =   "Entrega Ficha Cultivo"
            Height          =   315
            Index           =   1
            Left            =   3690
            TabIndex        =   32
            Top             =   750
            Width           =   2445
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   2
            Left            =   4920
            Picture         =   "frmManCampos.frx":5980
            ToolTipText     =   "Buscar fecha"
            Top             =   2640
            Width           =   240
         End
         Begin VB.Label Label44 
            Caption         =   "Fecha Revisión"
            Height          =   255
            Left            =   3720
            TabIndex        =   246
            Top             =   2640
            Width           =   1155
         End
         Begin VB.Label Label42 
            Caption         =   "Refer.Catastral"
            Height          =   255
            Left            =   150
            TabIndex        =   225
            Top             =   2640
            Width           =   1245
         End
         Begin VB.Label Label37 
            Caption         =   "Ficha Cultivo"
            Height          =   255
            Left            =   3720
            TabIndex        =   203
            Top             =   1920
            Width           =   1785
         End
         Begin VB.Label Label34 
            Caption         =   "Tipo Campo"
            Height          =   255
            Left            =   3720
            TabIndex        =   164
            Top             =   1530
            Width           =   1245
         End
         Begin VB.Label Label31 
            Caption         =   "Recolectado"
            Height          =   255
            Left            =   150
            TabIndex        =   111
            Top             =   1530
            Width           =   1035
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   1170
            ToolTipText     =   "Buscar Responsable"
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label13 
            Caption         =   "Responsable"
            Height          =   255
            Left            =   150
            TabIndex        =   110
            Top             =   1080
            Width           =   945
         End
         Begin VB.Label Label24 
            Caption         =   "Longitud"
            Height          =   255
            Left            =   5400
            TabIndex        =   105
            Top             =   2970
            Width           =   1185
         End
         Begin VB.Label Label21 
            Caption         =   "Latitud"
            Height          =   255
            Left            =   4140
            TabIndex        =   104
            Top             =   2970
            Width           =   945
         End
         Begin VB.Label Label19 
            Caption         =   "Nº OID"
            Height          =   255
            Left            =   2910
            TabIndex        =   103
            Top             =   2970
            Width           =   705
         End
         Begin VB.Label Label17 
            Caption         =   "Año plantación"
            Height          =   255
            Left            =   150
            TabIndex        =   102
            Top             =   1920
            Width           =   1185
         End
         Begin VB.Label Label14 
            Caption         =   "Nº Recinto"
            Height          =   255
            Left            =   150
            TabIndex        =   99
            Top             =   2280
            Width           =   915
         End
         Begin VB.Label Label10 
            Caption         =   "Subparcela"
            Height          =   255
            Left            =   1830
            TabIndex        =   98
            Top             =   2970
            Width           =   945
         End
         Begin VB.Label Label8 
            Caption         =   "Parcela"
            Height          =   255
            Left            =   990
            TabIndex        =   97
            Top             =   2940
            Width           =   705
         End
         Begin VB.Label Label7 
            Caption         =   "Poligono"
            Height          =   255
            Left            =   150
            TabIndex        =   96
            Top             =   2940
            Width           =   705
         End
         Begin VB.Label Label18 
            Caption         =   "Fecha Alta"
            Height          =   255
            Left            =   150
            TabIndex        =   90
            Top             =   360
            Width           =   795
         End
         Begin VB.Label Label22 
            Caption         =   "Fecha Baja"
            Height          =   255
            Left            =   3720
            TabIndex        =   89
            Top             =   360
            Width           =   975
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   0
            Left            =   1170
            Picture         =   "frmManCampos.frx":5A0B
            ToolTipText     =   "Buscar fecha"
            Top             =   360
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   1
            Left            =   4860
            Picture         =   "frmManCampos.frx":5A96
            ToolTipText     =   "Buscar fecha"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label23 
            Caption         =   "Situación"
            Height          =   255
            Left            =   150
            TabIndex        =   77
            Top             =   720
            Width           =   945
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   0
            Left            =   1170
            ToolTipText     =   "Buscar Situación"
            Top             =   720
            Width           =   240
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Datos Producción y Superficies"
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
         Height          =   2445
         Left            =   -74850
         TabIndex        =   70
         Top             =   2310
         Width           =   5415
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   6
            Left            =   1230
            MaxLength       =   10
            TabIndex        =   235
            Tag             =   "Sup.Sigpac|N|N|0|9999.9999|rcampos|supsigpa|###0.0000||"
            Text            =   "1234567890"
            Top             =   915
            Width           =   1395
         End
         Begin VB.TextBox txtAux2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   10
            Left            =   2670
            TabIndex        =   234
            Top             =   915
            Width           =   1185
         End
         Begin VB.TextBox txtAux2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   11
            Left            =   2670
            TabIndex        =   232
            Top             =   1275
            Width           =   1185
         End
         Begin VB.TextBox txtAux2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   12
            Left            =   2670
            TabIndex        =   231
            Top             =   1620
            Width           =   1185
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   33
            Left            =   1230
            MaxLength       =   10
            TabIndex        =   10
            Tag             =   "Sup.Cultivable|N|N|0|9999.9999|rcampos|supculti|###0.0000||"
            Top             =   1620
            Width           =   1395
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
            Height          =   285
            Index           =   33
            Left            =   3900
            MaxLength       =   10
            TabIndex        =   94
            Top             =   1620
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   9
            Left            =   3900
            MaxLength       =   5
            TabIndex        =   12
            Tag             =   "Arboles|N|S|0|9999|rcampos|nroarbol|#,##0||"
            Top             =   1995
            Width           =   1335
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
            Height          =   285
            Index           =   5
            Left            =   3900
            MaxLength       =   10
            TabIndex        =   91
            Text            =   "1234567890"
            Top             =   555
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   5
            Left            =   1230
            MaxLength       =   10
            TabIndex        =   8
            Tag             =   "Sup.Coop.|N|N|0|9999.9999|rcampos|supcoope|###0.0000||"
            Text            =   "1234567890"
            Top             =   555
            Width           =   1395
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
            Height          =   285
            Index           =   7
            Left            =   3900
            MaxLength       =   10
            TabIndex        =   93
            Top             =   1275
            Width           =   1335
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
            Height          =   285
            Index           =   6
            Left            =   3900
            MaxLength       =   10
            TabIndex        =   92
            Text            =   "1234567890"
            Top             =   915
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   7
            Left            =   1230
            MaxLength       =   10
            TabIndex        =   9
            Tag             =   "Sup.Catastro|N|N|0|9999.9999|rcampos|supcatas|###0.0000||"
            Top             =   1275
            Width           =   1395
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   8
            Left            =   1230
            MaxLength       =   7
            TabIndex        =   11
            Tag             =   "Aforo|N|S|0|999999|rcampos|canaforo|###,###||"
            Top             =   1995
            Width           =   1395
         End
         Begin VB.Image imgDoc 
            Height          =   405
            Index           =   1
            Left            =   3450
            ToolTipText     =   "Actualizar Hectáreas"
            Top             =   420
            Width           =   435
         End
         Begin VB.Label Label43 
            Caption         =   "Total Has.  Parcelas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   465
            Left            =   2670
            TabIndex        =   233
            Top             =   210
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label Label33 
            Caption         =   "Cultivable"
            Height          =   255
            Left            =   135
            TabIndex        =   163
            Top             =   1665
            Width           =   990
         End
         Begin VB.Label Label16 
            Caption         =   "Hanegadas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   255
            Left            =   3900
            TabIndex        =   101
            Top             =   270
            Width           =   1155
         End
         Begin VB.Label Label15 
            Caption         =   "Hectáreas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   255
            Left            =   1230
            TabIndex        =   100
            Top             =   270
            Width           =   1395
         End
         Begin VB.Label Label3 
            Caption         =   "Arboles"
            Height          =   255
            Left            =   2970
            TabIndex        =   95
            Top             =   2010
            Width           =   720
         End
         Begin VB.Label Label41 
            Caption         =   "Cooperativa"
            Height          =   255
            Left            =   135
            TabIndex        =   74
            Top             =   600
            Width           =   915
         End
         Begin VB.Label Label25 
            Caption         =   "Aforo"
            Height          =   255
            Left            =   135
            TabIndex        =   73
            Top             =   2040
            Width           =   960
         End
         Begin VB.Label Label20 
            Caption         =   "Sigpac"
            Height          =   255
            Left            =   135
            TabIndex        =   72
            Top             =   960
            Width           =   870
         End
         Begin VB.Label Label11 
            Caption         =   "Catastro"
            Height          =   255
            Left            =   135
            TabIndex        =   71
            Top             =   1320
            Width           =   990
         End
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   4155
         Left            =   -74910
         TabIndex        =   245
         Top             =   600
         Width           =   12120
         _ExtentX        =   21378
         _ExtentY        =   7329
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Campo"
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Partida"
            Object.Width           =   3000
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "C.Pobla"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Poblacion"
            Object.Width           =   2735
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Polígono"
            Object.Width           =   1464
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Parcela"
            Object.Width           =   1482
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Sp."
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Nro."
            Object.Width           =   1059
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Hdas"
            Object.Width           =   1305
         EndProperty
      End
      Begin VB.Frame FrameAux6 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   3910
         Left            =   -74760
         TabIndex        =   238
         Top             =   660
         Width           =   12210
         Begin VB.TextBox txtAux8 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   285
            Index           =   2
            Left            =   1710
            MaxLength       =   10
            TabIndex        =   242
            Tag             =   "Fecha|F|N|||rcampos_ordrec|fecimpre|dd/mm/yyyy||"
            Text            =   "Fecha"
            Top             =   2940
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtAux8 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   1
            Left            =   945
            MaxLength       =   6
            TabIndex        =   241
            Tag             =   "Orden|N|N|||rcampos_ordrec|nroorden|0000000|S|"
            Text            =   "linea"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtAux8 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   0
            Left            =   210
            MaxLength       =   8
            TabIndex        =   240
            Tag             =   "Campo|N|N|0|99999999|rcampos_ordrec|codcampo|00000000|S|"
            Text            =   "campo"
            Top             =   2910
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.CommandButton cmdAux 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   315
            Index           =   11
            Left            =   2550
            TabIndex        =   239
            ToolTipText     =   "Buscar fecha"
            Top             =   2910
            Visible         =   0   'False
            Width           =   195
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   6
            Left            =   45
            TabIndex        =   243
            Top             =   0
            Width           =   1710
            _ExtentX        =   3016
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
            Index           =   6
            Left            =   5280
            Top             =   210
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
            Bindings        =   "frmManCampos.frx":5B21
            Height          =   3225
            Index           =   6
            Left            =   45
            TabIndex        =   244
            Top             =   450
            Width           =   3900
            _ExtentX        =   6879
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
      Begin VB.Label Label36 
         Caption         =   "Nº Hidrante"
         Height          =   255
         Left            =   -71910
         TabIndex        =   187
         Top             =   1950
         Width           =   885
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   -73920
         ToolTipText     =   "Buscar Zona"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Parcela"
         Height          =   255
         Left            =   -74850
         TabIndex        =   88
         Top             =   1935
         Width           =   1035
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   -73920
         ToolTipText     =   "Buscar Partida"
         Top             =   870
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   -73920
         ToolTipText     =   "Buscar Variedad"
         Top             =   570
         Width           =   240
      End
      Begin VB.Label Label5 
         Caption         =   "Variedad"
         Height          =   255
         Left            =   -74820
         TabIndex        =   84
         Top             =   525
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Partida"
         Height          =   255
         Index           =   0
         Left            =   -74820
         TabIndex        =   83
         Top             =   855
         Width           =   735
      End
      Begin VB.Label Label29 
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   -69060
         TabIndex        =   82
         Top             =   4080
         Width           =   1140
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   -67830
         ToolTipText     =   "Zoom descripción"
         Top             =   4080
         Width           =   240
      End
      Begin VB.Label Label28 
         Caption         =   "Poblacion"
         Height          =   255
         Left            =   -74820
         TabIndex        =   81
         Top             =   1530
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Zona"
         Height          =   255
         Index           =   26
         Left            =   -74820
         TabIndex        =   80
         Top             =   1200
         Width           =   735
      End
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
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnVerificacionErr 
         Caption         =   "Verificacion Errores"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnSigpac 
         Caption         =   "Sigpac"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnGoolzoom 
         Caption         =   "Goolzoom"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnChequeoNroOrden 
         Caption         =   "Chequeo Nro.Orden"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnCambioSocio 
         Caption         =   "Cambio de Socio"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnGastosCampos 
         Caption         =   "Informe Gastos/Campos "
         Shortcut        =   ^D
      End
      Begin VB.Menu mnGlobalGap 
         Caption         =   "Asignación GlobalGap"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnBarra4 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmManCampos"
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

Public NroCampo As String

' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmC1 As frmCal 'calendario fecha
Attribute frmC1.VB_VarHelpID = -1
Private WithEvents frmC2 As frmCal 'calendario fecha
Attribute frmC2.VB_VarHelpID = -1
Private WithEvents frmC3 As frmCal 'calendario fecha
Attribute frmC3.VB_VarHelpID = -1
Private WithEvents frmC4 As frmCal 'calendario fecha
Attribute frmC4.VB_VarHelpID = -1
Private WithEvents frmC5 As frmCal 'calendario fecha
Attribute frmC5.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmZ2 As frmZoom  'Zoom para campos Text (observaciones de la revision)
Attribute frmZ2.VB_VarHelpID = -1

Private WithEvents frmPar As frmManPartidas 'partidas
Attribute frmPar.VB_VarHelpID = -1
Private WithEvents frmZon As frmManZonas 'zonas
Attribute frmZon.VB_VarHelpID = -1
Private WithEvents frmVar As frmComVar 'variedades
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmSit As frmManSituCamp 'situaciones de campos
Attribute frmSit.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmSoc1 As frmManSocios 'socios
Attribute frmSoc1.VB_VarHelpID = -1
Private WithEvents frmSoc2 As frmManSocios 'socios
Attribute frmSoc2.VB_VarHelpID = -1
Private WithEvents frmCalid As frmManCalidades 'calidades
Attribute frmCalid.VB_VarHelpID = -1
Private WithEvents frmInc As frmManInciden 'incidencias
Attribute frmInc.VB_VarHelpID = -1
Private WithEvents frmGlo As frmManGlobalGap 'ayuda de globalgap
Attribute frmGlo.VB_VarHelpID = -1
Private WithEvents frmCli As frmBasico 'ayuda de cliente de ariges(suministros)
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmGto As frmManConcepGasto ' ayuda de concepto de gasto
Attribute frmGto.VB_VarHelpID = -1

Private WithEvents frmRes As frmManCapataz 'responsable
Attribute frmRes.VB_VarHelpID = -1
Private WithEvents frmPlan As frmManPlantacion 'marco de plantacion
Attribute frmPlan.VB_VarHelpID = -1
Private WithEvents frmDesa As frmManDesarrollo 'desarrollo vegetativo
Attribute frmDesa.VB_VarHelpID = -1
Private WithEvents frmTie As frmManTierra 'tipo de tierra
Attribute frmTie.VB_VarHelpID = -1
Private WithEvents frmProc As frmManProceRiego 'procedencia
Attribute frmProc.VB_VarHelpID = -1
Private WithEvents frmPat As frmManPatronaPie 'patron a pie
Attribute frmPat.VB_VarHelpID = -1
Private WithEvents frmSegOp As frmManSeguroOpc 'seguro opcion
Attribute frmSegOp.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'mensajes traemos los campos que tienen mal el nro.orden
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmMens2 As frmMensajes 'mensajes traemos los contadores de pozos que tienen ese codigo de campo para cambiarle el socio
Attribute frmMens2.VB_VarHelpID = -1

' *****************************************************


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

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim VarieAnt As String
Dim SocioAnt As String
Dim FecBajaAnt As String
Dim cadCampos As String
Dim cadHidrantes As String

Dim indCodigo As Integer


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

Private Sub chkAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
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
                    CargarUnCampo CLng(Text1(0)), "I"
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
                Modificar
'                If ModificaDesdeFormulario2(Me, 1) Then
'                    TerminaBloquear
'                    PosicionarData
'                    CargaGrid 1, True
'                End If
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
                    SumaTotalPorcentajes NumTabMto
            End Select
        ' **************************
'            If NumTabMto = 1 Then
'                If Not vSeccion Is Nothing Then
'                    vSeccion.CerrarConta
'                    Set vSeccion = Nothing
'                End If
'            End If
    
    End Select
    Screen.MousePointer = vbDefault
    
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub cmdAux_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 1 'Calidades de la variedad de cabecera
            Set frmCalid = New frmManCalidades
            frmCalid.DatosADevolverBusqueda = "0|1|2|3|"
            frmCalid.CodigoActual = txtAux1(1).Text
            frmCalid.ParamVariedad = txtAux1(2).Text
            frmCalid.Show vbModal
            Set frmCalid = Nothing
            PonerFoco txtAux1(1)

        Case 0 ' Socios coopropietarios
            Set frmSoc1 = New frmManSocios
            frmSoc1.DatosADevolverBusqueda = "0|1|"
            frmSoc1.Show vbModal
            Set frmSoc1 = Nothing
            PonerFoco txtAux3(2)
            
        Case 2 ' Incidencias
            Set frmInc = New frmManInciden
            frmInc.DatosADevolverBusqueda = "0|1|"
            frmInc.Show vbModal
            Set frmInc = Nothing
            PonerFoco txtAux5(3)
        
        Case 3, 4 ' fecha de incidencia de agroseguro
           Screen.MousePointer = vbHourglass
           
           Dim esq As Long
           Dim dalt As Long
           Dim menu As Long
           Dim obj As Object
        
           Set frmC2 = New frmCal
            
           esq = cmdAux(Index).Left
           dalt = cmdAux(Index).Top
            
           Set obj = cmdAux(Index).Container
        
           While cmdAux(Index).Parent.Name <> obj.Name
                esq = esq + obj.Left
                dalt = dalt + obj.Top
                Set obj = obj.Container
           Wend
            
           menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
        
           frmC2.Left = esq + cmdAux(Index).Parent.Left + 30
           frmC2.Top = dalt + cmdAux(Index).Parent.Top + cmdAux(Index).Height + menu - 40
           
           frmC2.NovaData = Now
           Select Case Index
                Case 3
                    indice = 2
                Case 4
                    indice = 6
           End Select
           
           Me.cmdAux(0).Tag = indice
           
           PonerFormatoFecha txtAux5(indice)
           If txtAux5(indice).Text <> "" Then frmC2.NovaData = CDate(txtAux5(indice).Text)
        
           Screen.MousePointer = vbDefault
           frmC2.Show vbModal
           Set frmC2 = Nothing
           PonerFoco txtAux5(indice)
        
        Case 9 ' concepto de gasto
            Set frmGto = New frmManConcepGasto
            frmGto.DatosADevolverBusqueda = "0|1|"
            frmGto.Show vbModal
            Set frmGto = Nothing
            PonerFoco txtAux7(2)
        
        Case 10 ' fecha de concepto de gasto
           Screen.MousePointer = vbHourglass
           
           Set frmC3 = New frmCal
            
           esq = cmdAux(Index).Left
           dalt = cmdAux(Index).Top
            
           Set obj = cmdAux(Index).Container
        
           While cmdAux(Index).Parent.Name <> obj.Name
                esq = esq + obj.Left
                dalt = dalt + obj.Top
                Set obj = obj.Container
           Wend
            
           menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
        
           frmC3.Left = esq + cmdAux(Index).Parent.Left + 30
           frmC3.Top = dalt + cmdAux(Index).Parent.Top + cmdAux(Index).Height + menu - 40
           
           frmC3.NovaData = Now
           
           indice = 3
           
           Me.cmdAux(0).Tag = indice
           
           PonerFormatoFecha txtAux7(indice)
           If txtAux7(indice).Text <> "" Then frmC3.NovaData = CDate(txtAux7(indice).Text)
        
           Screen.MousePointer = vbDefault
           frmC3.Show vbModal
           Set frmC3 = Nothing
           PonerFoco txtAux7(indice)
        
        
        Case 11 ' fecha de impresion de orden de confeccion
           Screen.MousePointer = vbHourglass
           
           Set frmC4 = New frmCal
            
           esq = cmdAux(Index).Left
           dalt = cmdAux(Index).Top
            
           Set obj = cmdAux(Index).Container
        
           While cmdAux(Index).Parent.Name <> obj.Name
                esq = esq + obj.Left
                dalt = dalt + obj.Top
                Set obj = obj.Container
           Wend
            
           menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
        
           frmC4.Left = esq + cmdAux(Index).Parent.Left + 30
           frmC4.Top = dalt + cmdAux(Index).Parent.Top + cmdAux(Index).Height + menu - 40
           
           frmC4.NovaData = Now
           
           indice = 2
           
           Me.cmdAux(0).Tag = indice
           
           PonerFormatoFecha txtAux8(indice)
           If txtAux8(indice).Text <> "" Then frmC4.NovaData = CDate(txtAux8(indice).Text)
        
           Screen.MousePointer = vbDefault
           frmC4.Show vbModal
           Set frmC4 = Nothing
           PonerFoco txtAux8(indice)
        
        Case 12 ' fecha de revision del campo
           Screen.MousePointer = vbHourglass
           
           Set frmC5 = New frmCal
            
           esq = cmdAux(Index).Left
           dalt = cmdAux(Index).Top
            
           Set obj = cmdAux(Index).Container
        
           While cmdAux(Index).Parent.Name <> obj.Name
                esq = esq + obj.Left
                dalt = dalt + obj.Top
                Set obj = obj.Container
           Wend
            
           menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
        
           frmC5.Left = esq + cmdAux(Index).Parent.Left + 30
           frmC5.Top = dalt + cmdAux(Index).Parent.Top + cmdAux(Index).Height + menu - 40
           
           frmC5.NovaData = Now
           
           indice = 2
           
           Me.cmdAux(0).Tag = indice
           
           PonerFormatoFecha txtAux8(indice)
           If txtAux9(indice).Text <> "" Then frmC5.NovaData = CDate(txtAux9(indice).Text)
        
           Screen.MousePointer = vbDefault
           frmC5.Show vbModal
           Set frmC5 = Nothing
           PonerFoco txtAux9(indice)
        
        
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
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
    If PrimeraVez Then PrimeraVez = False

    'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
    If NroCampo <> "" And Not Data1.Recordset.EOF Then PonerCadenaBusqueda

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
    Screen.MousePointer = vbDefault
    '[Monica]03/10/2011: añadido el modo = 3 para solucionar problema de Picassent
    If Modo = 3 Or Modo = 4 Or Modo = 5 Then TerminaBloquear
    
    Set dbAriagro = Nothing

    '[Monica]28/11/2011: cliente de ariges
    If vParamAplic.BDAriges <> "" Then CerrarConexionAriges
End Sub

Private Sub Form_Load()
Dim i As Integer

    PrimeraVez = True
    
    ' ICONETS DE LA BARRA
    btnPrimero = 23 'index del botó "primero"
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
        .Buttons(11).Image = 28  ' verificacion de errores
        .Buttons(12).Image = 29
        .Buttons(13).Image = 30
        .Buttons(14).Image = 21  ' chequeo del nro de orden
        .Buttons(15).Image = 32  ' cambio de socio
        .Buttons(16).Image = 26  ' informe de gastos campos
        .Buttons(17).Image = 31  ' asignacion de codigos globalgap
        
        .Buttons(19).Image = 10  'Imprimir
        .Buttons(20).Image = 11  'Eixir
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
        
        If i = 5 Then ' boton de contabilizar un gasto de campo
            With Me.ToolAux(i)
                .HotImageList = frmPpal.imgListComun_OM16
                .DisabledImageList = frmPpal.imgListComun_BN16
                .ImageList = frmPpal.imgListComun16
                .Buttons(4).Image = 13   'Contabilizar
            End With
        End If
        
        If i = 7 Then
            With Me.ToolAux(i)
                .HotImageList = frmPpal.imgListComun_OM16
                .DisabledImageList = frmPpal.imgListComun_BN16
                .ImageList = frmPpal.imgListComun16
                .Buttons(4).Image = 10   'Impresion de revisiones de campos
            End With
        End If
    Next i
    ' ***********************************
    '[Monica]03/02/2015: solo para el caso de eescalona ponemos Arrendador
    If vParamAplic.Cooperativa = 10 Then
        Label4.Caption = "Arrendador"
        Text1(1).Text = "Código Arrendador|N|N|1|999999|rcampos|codsocio|000000|N|"
    End If
    
    
    Me.imgDoc(1).Picture = frmPpal.imgListPpal.ListImages(24).Picture
    Me.imgDoc(1).visible = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16)
    Me.imgDoc(1).Enabled = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16)
    
    'cargar IMAGES de busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
'    'carga IMAGES de mail
'    For i = 0 To Me.imgMail.Count - 1
'        Me.imgMail(i).Picture = frmPpal.imgListImages16.ListImages(2).Picture
'    Next i
    'IMAGES para zoom
    For i = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(i).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next i
    
    '[Monica]28/11/2011: cliente de ariges
    If vParamAplic.BDAriges <> "" Then
        If Not AbrirConexionAriges Then
            Unload Me
        End If
    End If
    
    ' *** si n'hi han tabs, per a que per defecte sempre es pose al 1r***
    Me.SSTab1.Tab = 0
    ' *******************************************************************
    
    LimpiarCampos   'Neteja els camps TextBox
'    ' ******* si n'hi han llínies *******
'    DataGridAux(0).ClearFields
    DataGridAux(1).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "rcampos"
    Ordenacion = " ORDER BY codcampo"
    '************************************************
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    CadenaConsulta = "Select * from " & NombreTabla
    
    If NroCampo <> "" Then
        CadenaConsulta = CadenaConsulta & " where codcampo = " & DBSet(NroCampo, "N")
    Else
        CadenaConsulta = CadenaConsulta & " where codcampo = -1 "
    End If
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    '***** canviar el nom de la PK de la capçalera; repasar codEmpre *************
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
       
    
    ModoLineas = 0
       
         
    '[Monica]14/02/2013: Totales de parcelas solo para Picassent
    Label43.visible = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16)
    For i = 10 To 12
        txtAux2(i).visible = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16)
    Next i
         
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
    Set dbAriagro = New BaseDatos
    dbAriagro.abrir_MYSQL vConfig.SERVER, vUsu.CadenaConexion, vConfig.User, vConfig.password
    
    '[Monica]07/06/2013: cambio del nommbre de zona
    Label1(26).Caption = vParamAplic.NomZonaPOZ
    imgBuscar(13).ToolTipText = "Buscar " & vParamAplic.NomZonaPOZ
    
    '[Monica]23/09/2014: en el caso de alzira el campo poda lo usaran para indicar si el campo está sin Placa Identificativa
    If vParamAplic.Cooperativa = 4 Then
        chkAbonos(3).Tag = "Sin Placa Identif.|N|N|||rcampos|conpoda||N|"
        chkAbonos(3).Caption = "Sin Placa Identif."
    End If
    
    '[Monica]02/10/2014: Las revisiones de campos unicamente las ve catadau
    SSTab1.TabVisible(9) = (vParamAplic.Cooperativa = 0)
    
End Sub


Private Sub LimpiarCampos()
Dim i As Integer

    On Error Resume Next
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
    ' *** si n'hi han combos a la capçalera ***
    ' *****************************************
    For i = 0 To Combo1.Count - 1
        Combo1(i).ListIndex = -1
    Next i
    For i = 0 To chkAbonos.Count - 1
        Me.chkAbonos(i).Value = 0
    Next i
    Me.chkAux(0).Value = 0

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
    If DatosADevolverBusqueda <> "" Or NroCampo <> "" Then
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
    For i = 0 To 7
        BloquearChk Me.chkAbonos(i), (Modo = 0 Or Modo = 2 Or Modo = 5)
    Next i
    
    
    ' ***** bloquejar tots els controls visibles de la clau primaria de la capçalera ***
'    If Modo = 4 Then _
'        BloquearTxt Text1(0), True 'si estic en  modificar, bloqueja la clau primaria
    ' **********************************************************************************
    
    ' **** si n'hi han imagens de buscar en la capçalera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    For i = 0 To imgFec.Count - 1
        BloquearImgFec Me, i, Modo
    Next i
'    BloquearImgFec Me, 25, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
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
    
    ' el codigo de socio solo se puede modificar de un campo si se hace un cambio de socio
    BloquearTxt Text1(1), Not (Modo = 1 Or Modo = 3)
    imgBuscar(1).Enabled = (Modo = 1 Or Modo = 3)
    imgBuscar(1).visible = (Modo = 1 Or Modo = 3)
    
    
    PonerLongCampos

    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 0, False
        CargaGrid 1, False
        CargaGrid 2, False
        CargaGrid 3, False
        CargaGrid 4, False
        CargaGrid 5, False
        CargaGrid 7, False
        '[Monica]30/09/2013
        'CargaGrid 6, False
        CargarListaOrdenesRecogida "-1"
    End If
    
    b = (Modo = 4) Or (Modo = 2)
    

    DataGridAux(0).Enabled = b
    DataGridAux(1).Enabled = b
    DataGridAux(2).Enabled = b
    DataGridAux(3).Enabled = b
    DataGridAux(4).Enabled = b
    DataGridAux(5).Enabled = b
    DataGridAux(7).Enabled = b
    '[Monica]30/09/2013
    'DataGridAux(6).Enabled = b
    
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
        
'    'telefonos
'    b = (Modo = 5) And (NumTabMto = 0) 'And (ModoLineas <> 3)
'    For i = 1 To 4
'        BloquearTxt txtAux(i), Not b
'    Next i
'    For i = 5 To txtAux.Count - 1
'        BloquearTxt txtAux(i), Not b
'    Next i
'    Me.chkAbonos(1).Enabled = b
'    b = (Modo = 5) And (NumTabMto = 0) And ModoLineas = 2
'    BloquearTxt txtAux(1), b
'
    'clasificacion
    b = (Modo = 5) And (NumTabMto = 1) 'And (ModoLineas <> 3)
    For i = 1 To txtAux1.Count - 1
        BloquearTxt txtAux1(i), Not b
    Next i
    b = (Modo = 5) And (NumTabMto = 1) And ModoLineas = 2
    BloquearTxt txtAux1(1), b
    BloquearBtn cmdAux(1), b
     '-----------------------------
     
    Text1(37).Enabled = (Modo = 1)
    imgBuscar(13).Enabled = (Modo = 1)
    imgBuscar(13).visible = (Modo = 1)
    
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        Text1(37).Enabled = (Modo = 1) Or (Modo = 3) Or (Modo = 4)
        imgBuscar(13).Enabled = (Modo = 1) Or (Modo = 3) Or (Modo = 4)
        imgBuscar(13).visible = (Modo = 1) Or (Modo = 3) Or (Modo = 4)
    End If
     
    PonerModoOpcionesMenu (Modo) 'Activar opcions menú según modo
    PonerOpcionesMenu   'Activar opcions de menú según nivell
                        'de permisos de l'usuari

    '[Monica]06/05/2015: solo para el caso de alzira no dejamos modificar la superficie cooperativa si no tiene nivel 0
    If vParamAplic.Cooperativa = 4 And Modo = 4 Then
        Text1(5).Enabled = (vUsu.Nivel = 0)
        Text2(5).Enabled = (vUsu.Nivel = 0)
    End If


    ' bloqueo de todos los datos excepto de datos tecnicos cuando no es administrador y estamos modificando
    b = (Modo = 4) And vUsu.Nivel > 1
    
    BloquearTodoExceptoDatosTecnicos b
    
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
    b = (Modo = 2 Or Modo = 0) And NroCampo = ""
    'Buscar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnBuscar.Enabled = b
    'Vore Tots
    Toolbar1.Buttons(4).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnNuevo.Enabled = b
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0 And NroCampo = "") 'And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(9).Enabled = b
    Me.mnEliminar.Enabled = b
    'Verificacion de errores
    Toolbar1.Buttons(11).Enabled = b
    Me.mnVerificacionErr.Enabled = b
    'Sigpac
    Toolbar1.Buttons(12).Enabled = b
    Me.mnSigpac.Enabled = b
    'Goolzoom
    Toolbar1.Buttons(13).Enabled = b
    Me.mnGoolzoom.Enabled = b
    
    'Chequeo del Nro de Orden
    Toolbar1.Buttons(14).Enabled = (Modo = 0 Or Modo = 2) And vParamAplic.Cooperativa = 4
    Me.mnChequeoNroOrden.Enabled = (Modo = 0 Or Modo = 2) And vParamAplic.Cooperativa = 4
    
    'Cambio de socio de un campo
    Toolbar1.Buttons(15).Enabled = b
    Me.mnCambioSocio.Enabled = b
    
    'Gastos Pendientes de Integrar
    Toolbar1.Buttons(16).Enabled = b
    Me.mnGastosCampos.Enabled = b
    
    '[Monica]10/11/2015. nuevo punto de menu de recalculo de globalgap
    Toolbar1.Buttons(17).Enabled = (Modo = 0 Or Modo = 2)
    Me.mnGlobalGap.Enabled = (Modo = 0 Or Modo = 2)
    
    'Imprimir
    Toolbar1.Buttons(19).Enabled = (Modo = 2 And Data1.Recordset.RecordCount > 0)
    Me.mnImprimir.Enabled = (Modo = 2 And Data1.Recordset.RecordCount > 0)
    
    '[Monica]14/02/2013: Actualizacion de las superficies solo para Picassent
    imgDoc(1).visible = (Modo = 2 And Data1.Recordset.RecordCount > 0 And (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16))
    imgDoc(1).Enabled = (Modo = 2 And Data1.Recordset.RecordCount > 0 And (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16))
    
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
    b = (Modo = 3 Or Modo = 4 Or Modo = 2 And NroCampo = "")
    For i = 0 To ToolAux.Count - 1 '[Monica]30/09/2013: antes - 1
        If i <> 6 Then
            ToolAux(i).Buttons(1).Enabled = b
            If b Then bAux = (b And Me.AdoAux(i).Recordset.RecordCount > 0)
            ToolAux(i).Buttons(2).Enabled = bAux
            ToolAux(i).Buttons(3).Enabled = bAux
        End If
    Next i
    
    ToolAux(4).Buttons(1).Enabled = b And vUsu.Login = "root"
    If b Then bAux = (b And Me.AdoAux(4).Recordset.RecordCount > 0)
    ToolAux(4).Buttons(2).Enabled = bAux And vUsu.Login = "root"
    ToolAux(4).Buttons(3).Enabled = bAux And vUsu.Login = "root"
    
    ' boton de integracion contable
    bAux = b And Me.AdoAux(5).Recordset.RecordCount > 0
    If Me.AdoAux(5).Recordset.RecordCount > 0 Then
        bAux = bAux And CInt(AdoAux(5).Recordset.Fields(6).Value) = 0
    End If
        
    ToolAux(5).Buttons(4).Enabled = bAux
    
    ' boton de impresion de revisiones de campos
    ToolAux(7).Buttons(4).Enabled = True
    
    
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
        Case 0
            Sql = "select rcampos_cooprop.codcampo, rcampos_cooprop.numlinea, rcampos_cooprop.codsocio, rsocios.nomsocio, "
            Sql = Sql & " rcampos_cooprop.porcentaje "
            Sql = Sql & " FROM rcampos_cooprop INNER JOIN rsocios ON rcampos_cooprop.codsocio = rsocios.codsocio "
            Sql = Sql & " and rcampos_cooprop.codsocio = rsocios.codsocio "
            If enlaza Then
                '[Monica]08/07/2011
                Sql = Sql & Replace(ObtenerWhereCab2(True), "rcampos_parcelas", "rcampos_cooprop")
                'Sql = Sql & " WHERE rcampos_cooprop.codcampo = " & Text1(0).Text
            Else
                Sql = Sql & " WHERE rcampos_cooprop.codcampo = -1"
            End If
            Sql = Sql & " ORDER BY rcampos_cooprop.codsocio "
       
       Case 1 ' clasificacion
            Tabla = "rcampos_clasif"
            Sql = "SELECT rcampos_clasif.codcampo, rcampos_clasif.codvarie, rcampos_clasif.codcalid, rcalidad.nomcalid, rcampos_clasif.muestra "
            Sql = Sql & " FROM " & Tabla & " INNER JOIN rcalidad ON rcampos_clasif.codvarie = rcalidad.codvarie "
            Sql = Sql & " and rcampos_clasif.codcalid = rcalidad.codcalid "
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE rcampos_clasif.codcampo = -1"
            End If
            Sql = Sql
            Sql = Sql & " ORDER BY " & Tabla & ".codcalid "
            
       Case 2 ' parcelas
            Tabla = "rcampos_parcelas"
            Sql = "SELECT rcampos_parcelas.codcampo, rcampos_parcelas.numlinea, rcampos_parcelas.poligono,rcampos_parcelas.parcela,rcampos_parcelas.subparce, "
            Sql = Sql & "rcampos_parcelas.recintos,rcampos_parcelas.codsigpa,rcampos_parcelas.supsigpa,"
            Sql = Sql & "rcampos_parcelas.supcultsigpa,rcampos_parcelas.supcatas,rcampos_parcelas.supcultcatas"
            Sql = Sql & " FROM " & Tabla
            If enlaza Then
                Sql = Sql & ObtenerWhereCab2(True)
            Else
                Sql = Sql & " WHERE rcampos_parcelas.codcampo = -1"
            End If
            Sql = Sql & " ORDER BY " & Tabla & ".numlinea "
            
       Case 3 ' agroseguro
            Tabla = "rcampos_seguros"
            Sql = "SELECT rcampos_seguros.codcampo, rcampos_seguros.numlinea, rcampos_seguros.fecha, rcampos_seguros.codincid, rincidencia.nomincid, "
            Sql = Sql & "rcampos_seguros.kilos,rcampos_seguros.kilosaportacion, rcampos_seguros.importe,rcampos_seguros.fechapago,"
            Sql = Sql & "rcampos_seguros.essiniestro , IF(essiniestro=1,'*','') as dsiniestro "
            Sql = Sql & " FROM " & Tabla & " INNER JOIN rincidencia ON rcampos_seguros.codincid = rincidencia.codincid "
            If enlaza Then
                Sql = Sql & ObtenerWhereCab3(True)
            Else
                Sql = Sql & " WHERE rcampos_seguros.codcampo = -1"
            End If
            Sql = Sql & " ORDER BY " & Tabla & ".numlinea "
            
        Case 4 ' hco del campo
            Tabla = "rcampos_hco"
            Sql = "SELECT rcampos_hco.codcampo, rcampos_hco.numlinea, rcampos_hco.codsocio, rsocios.nomsocio, rcampos_hco.fechaalta, "
            Sql = Sql & "rcampos_hco.fechabaja, rcampos_hco.codincid, rincidencia.nomincid"
            Sql = Sql & " FROM (" & Tabla & " INNER JOIN rincidencia ON rcampos_hco.codincid = rincidencia.codincid) "
            Sql = Sql & " INNER JOIN rsocios ON rcampos_hco.codsocio = rsocios.codsocio "
            If enlaza Then
                Sql = Sql & ObtenerWhereCab4(True)
            Else
                Sql = Sql & " WHERE rcampos_hco.codcampo = -1"
            End If
            Sql = Sql & " ORDER BY " & Tabla & ".numlinea "
        
        Case 5 ' gastos del campo
            Tabla = "rcampos_gastos"
            Sql = "SELECT rcampos_gastos.codcampo, rcampos_gastos.numlinea, rcampos_gastos.codgasto, rconcepgasto.nomgasto, rcampos_gastos.fecha, "
            Sql = Sql & "rcampos_gastos.importe, rcampos_gastos.contabilizado, IF(contabilizado=1,'*','') as dcontabilizado "
            Sql = Sql & " FROM " & Tabla & " INNER JOIN rconcepgasto ON rcampos_gastos.codgasto = rconcepgasto.codgasto "
            
            If enlaza Then
                Sql = Sql & ObtenerWhereCab5(True)
            Else
                Sql = Sql & " WHERE rcampos_gastos.codcampo = -1"
            End If
            Sql = Sql & " ORDER BY " & Tabla & ".numlinea "
    
        Case 6 ' impresion de ordenes de recoleccion del campo
            Tabla = "rcampos_ordrec"
            Sql = "SELECT rcampos_ordrec.codcampo, rcampos_ordrec.nroorden, rcampos_ordrec.fecimpre "
            Sql = Sql & " FROM " & Tabla
            
            If enlaza Then
                Sql = Sql & ObtenerWhereCab6(True)
            Else
                Sql = Sql & " WHERE rcampos_ordrec.codcampo = -1"
            End If
            Sql = Sql & " ORDER BY " & Tabla & ".nroorden "
    
        Case 7 ' revisiones
            Tabla = "rcampos_revision"
            Sql = "SELECT rcampos_revision.codcampo, rcampos_revision.numlinea, rcampos_revision.fecha, rcampos_revision.tecnico, rcampos_revision.observac "
            Sql = Sql & " FROM " & Tabla
            
            If enlaza Then
                Sql = Sql & ObtenerWhereCab7(True)
            Else
                Sql = Sql & " WHERE rcampos_revision.codcampo = -1"
            End If
            Sql = Sql & " ORDER BY " & Tabla & ".numlinea "
    
    
    
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
    txtAux1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmC1_Selec(vFecha As Date)
Dim indice As Byte
    indice = CByte(Me.imgFec(0).Tag)
    Text1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmC2_Selec(vFecha As Date)
Dim indice As Byte
    indice = CByte(Me.cmdAux(0).Tag)
    txtAux5(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmC3_Selec(vFecha As Date)
Dim indice As Byte
    indice = CByte(Me.cmdAux(0).Tag)
    txtAux7(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmC4_Selec(vFecha As Date)
Dim indice As Byte
    indice = CByte(Me.cmdAux(0).Tag)
    txtAux8(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmC5_Selec(vFecha As Date)
Dim indice As Byte
    indice = CByte(Me.cmdAux(0).Tag)
    txtAux9(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCalid_DatoSeleccionado(CadenaSeleccion As String)
    txtAux1(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo variedad
    txtAux1(1).Text = RecuperaValor(CadenaSeleccion, 3) 'codigo calidad
    txtAux2(1).Text = RecuperaValor(CadenaSeleccion, 4) 'nombre calidad
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'cliente de suministros
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo de cliente de suministros
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre
End Sub

Private Sub frmDesa_DatoSeleccionado(CadenaSeleccion As String)
    Text1(26).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo desarrollo vegetativo
    FormateaCampo Text1(26)
    Text2(26).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre desarrollo vegetativo
End Sub

Private Sub frmGlo_DatoSeleccionado(CadenaSeleccion As String)
'globalgap
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo de globalgap
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmGto_DatoSeleccionado(CadenaSeleccion As String)
    txtAux7(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo concepto de gasto
    FormateaCampo txtAux7(2)
    txtAux2(5).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre concepto de gasto
End Sub

Private Sub frmInc_DatoSeleccionado(CadenaSeleccion As String)
    txtAux5(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo incidencia
    FormateaCampo txtAux5(3)
    txtAux2(2).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre incidencia
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String

    If CadenaSeleccion <> "" Then
        Sql = " rcampos.codcampo in (" & CadenaSeleccion & ")"
    Else
        Sql = ""
    End If
    cadCampos = Sql
    
End Sub

Private Sub frmMens2_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String

    If CadenaSeleccion <> "" Then
        Sql = " rpozos.hidrante in (" & CadenaSeleccion & ")"
    Else
        Sql = ""
    End If
    cadHidrantes = Sql
    
End Sub




Private Sub frmPar_DatoSeleccionado(CadenaSeleccion As String)
'partidas
Dim Zona As String
Dim Poblacion As String
Dim OtroCampo As String
Dim CodPobla As String

    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo de partida
    FormateaCampo Text1(3)
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre de partida
    
    
    '[Monica]23/05/2013: sustituyo esto por el ponerdatospartida
'    Text1(37).Text = RecuperaValor(CadenaSeleccion, 5) 'codzona
'    Text4(3).Text = RecuperaValor(CadenaSeleccion, 6) 'nomzona
'    Text5(3).Text = RecuperaValor(CadenaSeleccion, 4)
    PonerDatosPartida
    
End Sub

Private Sub frmPat_DatoSeleccionado(CadenaSeleccion As String)
    Text1(31).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo patron a pie
    FormateaCampo Text1(31)
    Text2(31).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre patron a pie
End Sub

Private Sub frmPlan_DatoSeleccionado(CadenaSeleccion As String)
    Text1(25).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo marco plantacion
    FormateaCampo Text1(25)
    Text2(25).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre marco plantacion
End Sub

Private Sub frmProc_DatoSeleccionado(CadenaSeleccion As String)
    Text1(30).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo procedencia de riego
    FormateaCampo Text1(30)
    Text2(30).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre procedencia de riego
End Sub

Private Sub frmRes_DatoSeleccionado(CadenaSeleccion As String)
    Text1(24).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo capataz responsable
    FormateaCampo Text1(24)
    Text2(24).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre capataz
End Sub

Private Sub frmSegOp_DatoSeleccionado(CadenaSeleccion As String)
    Text1(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo seguro opcion
    FormateaCampo Text1(indCodigo)
    Text2(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre seguro opcion
End Sub

Private Sub frmSit_DatoSeleccionado(CadenaSeleccion As String)
    Text1(12).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo situacion
    FormateaCampo Text1(12)
    Text2(12).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre situacion
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    Text1(1).Text = RecuperaValor(CadenaSeleccion, 1) 'codsocio
    FormateaCampo Text1(indice)
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2) 'nomsocio
End Sub

Private Sub frmSoc1_DatoSeleccionado(CadenaSeleccion As String)
    txtAux3(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codsocio
    FormateaCampo txtAux3(2)
    txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'nomsocio
End Sub


Private Sub frmTie_DatoSeleccionado(CadenaSeleccion As String)
    Text1(27).Text = RecuperaValor(CadenaSeleccion, 1) 'tipo de tierra
    FormateaCampo Text1(27)
    Text2(27).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre de tierra
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codvariedad
    FormateaCampo Text1(2)
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre variedad
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
    If indice = 21 Then
        Text1(indice).Text = vCampo
    Else
        txtAux9(indice).Text = vCampo
    End If
    
End Sub

Private Sub frmZon_DatoSeleccionado(CadenaSeleccion As String)
    Text1(37).Text = RecuperaValor(CadenaSeleccion, 1) 'codzona
    FormateaCampo Text1(indice)
    Text4(3).Text = RecuperaValor(CadenaSeleccion, 2) 'nomzona
End Sub

Private Sub imgDoc_Click(Index As Integer)
Dim Mens As String

    Mens = "Se va a proceder a actualizar las superficies de la ficha de campo "
    Mens = Mens & vbCrLf
    Mens = Mens & "con los datos de las parcelas."
    Mens = Mens & vbCrLf & vbCrLf
    Mens = Mens & "¿ Desea continuar ? " & vbCrLf & vbCrLf

    If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
    
        Text1(6).Text = txtAux2(10).Text
        Text1(7).Text = txtAux2(11).Text
        Text1(33).Text = txtAux2(12).Text
        
        mnModificar_Click
        cmdAceptar_Click
        PonerModo 2
        PonerCampos

    End If


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
                indice = Index + 10
            Case 2, 3
                indice = Index + 40
       End Select
       
       Me.imgFec(0).Tag = indice
       
       PonerFormatoFecha Text1(indice)
       If Text1(indice).Text <> "" Then frmC1.NovaData = CDate(Text1(indice).Text)
    
       Screen.MousePointer = vbDefault
       frmC1.Show vbModal
       Set frmC1 = Nothing
       PonerFoco Text1(indice)
    
End Sub

Private Sub imgZoom_Click(Index As Integer)
    Set frmZ = New frmZoom

    Select Case Index
        Case 0
            indice = 21
            frmZ.pTitulo = "Observaciones del Campo"
            frmZ.pValor = Text1(indice).Text
            frmZ.pModo = Modo
            frmZ.Show vbModal
            Set frmZ = Nothing
            PonerFoco Text1(indice)
            
        Case 1
            indice = 4
            frmZ.pTitulo = "Observaciones de la Revisión"
            If Modo = 5 Then
                frmZ.pModo = 3
                frmZ.pValor = txtAux9(indice).Text
            Else
                frmZ.pModo = Modo
                frmZ.pValor = DBLet(Me.AdoAux(7).Recordset!Observac, "T")
            End If
            frmZ.Show vbModal
            Set frmZ = Nothing
            PonerFoco Text1(indice)
            
            
    End Select
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnCambioSocio_Click()
    BotonCambioSocio
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnGastosCampos_Click()
    ' Impresion de los gastos por campo
    AbrirListado 36

End Sub

Private Sub mnGlobalGap_Click()
    AbrirListado 49
    
End Sub

Private Sub mnGoolzoom_Click()
Dim Direccion As String

'    Direccion = "www.goolzoom.com/mapa.html?lat=" & Trim(TransformaComasPuntos(Text1(19).Text)) & "&lng=" & Trim(TransformaComasPuntos(Text1(20).Text)) & "&zoom=18"

    If vParamAplic.GoolZoom <> "" Then
        Direccion = Replace(Replace(vParamAplic.GoolZoom, "LATITUD", TransformaComasPuntos(Text1(19).Text)), "LONGITUD", TransformaComasPuntos(Text1(20).Text))
    Else
        MsgBox "No tiene configurada en parámetros la dirección de Goolzoom. Llame a Soporte.", vbExclamation
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass

    If LanzaHomeGnral(Direccion) Then espera 2
    Screen.MousePointer = vbDefault


End Sub

Private Sub mnImprimir_Click()
    printNou
End Sub

Private Sub mnModificar_Click()
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************
    
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
End Sub

Private Sub mnChequeoNroOrden_Click()
    ChequeoNroOrden
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnSigpac_Click()
Dim Direccion As String
Dim Pobla As String
Dim Municipio As String

    'http://sigpac.mapa.es/fega/visor/LayerInfo.aspx?layer=PARCELA&id=OID&image=ORTOFOTOS
'    Direccion = "http://sigpac.mapa.es/fega/visor/LayerInfo.aspx?layer=PARCELA&id=" & Trim(Text1(18).Text) & "&image=ORTOFOTOS"
    
    If vParamAplic.SigPac <> "" Then
        If InStr(1, vParamAplic.SigPac, "NUMOID") <> 0 Then
            Direccion = Replace(vParamAplic.SigPac, "NUMOID", Text1(18).Text)
        Else
            Pobla = DevuelveDesdeBDNew(cAgro, "rpartida", "codpobla", "codparti", Text1(3).Text, "N")
            If Pobla = "" Then
                MsgBox "No existe el código de poblacion de la partida", vbExclamation
            Else
                Municipio = DevuelveDesdeBDNew(cAgro, "rpueblos", "codsigpa", "codpobla", Pobla, "T")
                Direccion = Replace(vParamAplic.SigPac, "[PR]", Mid(Pobla, 1, 2))
                Direccion = Replace(Direccion, "[MN]", CInt(Municipio))
                Direccion = Replace(Direccion, "[PL]", CInt(Text1(13).Text))
                Direccion = Replace(Direccion, "[PC]", CInt(Text1(14).Text))
                Direccion = Replace(Direccion, "[RC]", CInt(Text1(16).Text))
            End If
        End If
    Else
        MsgBox "No tiene configurada en parámetros la dirección de Sigpac. Llame a Soporte.", vbExclamation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass

'    If LanzaHomeGnral(Direccion) Then espera 2
    LanzaVisorMimeDocumento Me.hWnd, Direccion
    Screen.MousePointer = vbDefault

End Sub

Private Sub mnVerificacionErr_Click()
Dim Sql As String
Dim Rs As ADODB.Recordset
    
    InicializarVbles
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    Sql = "select rpueblos.codpobla, rcampos.poligono, rcampos.parcela, rcampos.recintos, count(*) "
    Sql = Sql & " from (rcampos inner join rpartida on rcampos.codparti = rpartida.codparti) "
    Sql = Sql & " inner join rpueblos on rpartida.codpobla = rpueblos.codpobla"
    Sql = Sql & " where rcampos.fecbajas is null "
    Sql = Sql & " group by 1,2,3,4 "
    Sql = Sql & " having count(*) > 1"
    
    If TotalRegistros(Sql) <> 0 Then
        cadNombreRPT = "rErroresCampos.rpt"
        cadTitulo = "Campos con duplicidades"
        frmImprimir.Opcion = 0
        LlamarImprimir
    End If

End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub







Private Sub Text2_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text2(Index), Modo
End Sub

Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text2_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

    If Not PerderFocoGnral(Text2(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 5, 6, 7, 33 'superficies en hectareas
            If Modo = 1 Then Exit Sub
            If Text2(Index).Text <> "" Then
                PonerFormatoDecimal Text2(Index), 3
                Text1(Index).Text = Round2(ImporteSinFormato(Text2(Index).Text) * vParamAplic.Faneca, 4)
                PonerFormatoDecimal Text1(Index), 7
            Else
                If Index = 5 Then
                    EstablecerOrden True
                End If
            End If
            
            If Index = 33 Then PonerFoco Text1(8)
                
            
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
        Case 11 ' Verificacion de Errores
            mnVerificacionErr_Click
        Case 12
            mnSigpac_Click
        Case 13
            mnGoolzoom_Click
        Case 14
            mnChequeoNroOrden_Click
        Case 15
            mnCambioSocio_Click
        Case 16
            mnGastosCampos_Click
        Case 17
            mnGlobalGap_Click
        Case 19 'Imprimir
'            AbrirListado (10)
            mnImprimir_Click
        Case 20    'Eixir
            mnSalir_Click
            
        Case btnPrimero To btnPrimero + 3 'Fleches Desplaçament
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonBuscar()
Dim i As Integer
' ***** Si la clau primaria de la capçalera no es Text1(0), canviar-ho en <=== *****
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonerFoco Text1(0) ' <===
        Text1(0).BackColor = vbYellow ' <===
        ' *** si n'hi han combos a la capçalera ***
        
        EstablecerOrden True
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
    Dim NombreTabla1 As String
        
    'Cridem al form
    ' **************** arreglar-ho per a vore lo que es desije ****************
    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
    Cad = ""
    Cad = Cad & "Código|rcampos.codcampo|N|000000|9·"
    Cad = Cad & "Socio|rcampos.codsocio|N|000000|9·"
    Cad = Cad & "Nombre|rsocios.nomsocio|T||34·"
    Cad = Cad & "Variedad|variedades.nomvarie|T||17·"
    Cad = Cad & "Partida|rpartida.nomparti|T||15·"
    Cad = Cad & "Pol.|rcampos.poligono|T||5·"
    Cad = Cad & "Parc.|rcampos.parcela|T||7·"
    Cad = Cad & "Sp.|rcampos.subparce|T||4·"
    
    NombreTabla1 = "((rcampos inner join rsocios on rcampos.codsocio = rsocios.codsocio) inner join variedades on rcampos.codvarie = variedades.codvarie) " & _
                   " inner join rpartida on rcampos.codparti = rpartida.codparti "
    
    
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = NombreTabla1
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        frmB.vDevuelve = "0|" '*** els camps que volen que torne ***
        frmB.vTitulo = "Campos" ' ***** repasa açò: títol de BuscaGrid *****
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
Dim j As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    Cad = ""
    i = 0
    Do
        j = i + 1
        i = InStr(j, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, j, i - j)
            j = Val(Aux)
            Cad = Cad & Text1(j).Text & "|"
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
    Text1(0).Text = SugerirCodigoSiguienteStr("rcampos", "codcampo")
    FormateaCampo Text1(0)
       
    Text1(16).Text = "1"
    Text1(5).Text = "0,0000"
    Text1(6).Text = "0,0000"
    Text1(7).Text = "0,0000"
    Text1(33).Text = "0,0000"
    
    Text1(10).Text = Format(Now, "dd/mm/yyyy")
       
    '[Monica]29/09/2014: comprobamos si vamos a dar de baja que no tenga fecha de alta en programa operativo
    FecBajaAnt = Text1(11).Text
       
    PonerFoco Text1(0) '*** 1r camp visible que siga PK ***
    ' ***********************************************************
    
    ' *** si n'hi han camps de descripció a la capçalera ***
    PosarDescripcions
    ' ******************************************************

    ' *** si n'hi han tabs, em posicione al 1r ***
    Me.SSTab1.Tab = 0
    ' ********************************************
    
    EstablecerOrden True
End Sub


Private Sub BotonModificar()

    PonerModo 4

    ' *** bloquejar els camps visibles de la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    ' *************************************************************************
    
    VarieAnt = Text1(2).Text
    SocioAnt = Text1(1).Text
    '[Monica]29/09/2014: comprobamos si vamos a dar de baja que no tenga fecha de alta en programa operativo
    FecBajaAnt = Text1(11).Text
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(22)
    ' *********************************************************
    
    EstablecerOrden True
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
    Cad = "¿Seguro que desea eliminar el Campo?"
    Cad = Cad & vbCrLf & "Código: " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
    Cad = Cad & vbCrLf & "Socio : " & Data1.Recordset.Fields(1)
    Cad = Cad & vbCrLf & "Nombre: " & Text2(1).Text
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
    For i = 0 To DataGridAux.Count - 1 '[Monica]30/09/2013: antes - 1
        If i <> 6 Then
            CargaGrid i, True
            If Not AdoAux(i).Recordset.EOF Then _
                PonerCamposForma2 Me, AdoAux(i), 2, "FrameAux" & i
        End If
    Next i
    '[Monica]30/09/2013
    CargarListaOrdenesRecogida Text1(23).Text
    
    ' *******************************************

    ' *** si n'hi han llínies sense datagrid ***
    ' ***************************************
    
    ' ************* configurar els camps de les descripcions de la capçalera *************
'    txtAux2(22).Text = PonerNombreDeCod(txtAux(22), "poblacio", "despobla", "codpobla", "N")

    PosarDescripcions

    '[Monica]23/10/2013: Solo si es Escalona o Utxera (o de momento montifrut) damos mensaje de que el socio tiene pagos pendientes
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Or vParamAplic.Cooperativa = 12 Then
        '[Monica]15/05/2013: Visualizamos los cobros pendientes del socio
        ComprobarCobrosSocio CStr(Data1.Recordset!Codsocio), ""
    End If

    PonerClasificacionGrafica

    VisualizaClasificacion

    SumaTotalPorcentajes 0
    ' ********************************************************************************
    
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu   'Activar opcions de menú según nivell
                        'de permisos de l'usuari
    
End Sub


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
                    If NumTabMto = 0 Or NumTabMto = 1 Or NumTabMto = 2 Or NumTabMto = 3 Or NumTabMto = 4 Or NumTabMto = 5 Or NumTabMto = 6 Or NumTabMto = 7 Then
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
'                        Select Case NumTabMto
'                            Case 0 'coopropietarios
'                                'BotonModificar
'                                For I = 0 To txtAux3.Count - 1
'                                    txtAux3(I).Text = ""
'                                    BloquearTxt txtAux3(I), True
'                                Next I
'                                txtAux2(0).Text = ""
'                                BloquearTxt txtAux2(0), True
'                            Case 1 'secciones
'                                For I = 0 To txtaux1.Count - 1
'                                    txtaux1(I).Text = ""
'                                    BloquearTxt txtaux1(I), True
'                                Next I
'                                txtAux2(1).Text = ""
'                                BloquearTxt txtAux2(1), True
'                        End Select
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
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
            End Select
            
            SumaTotalPorcentajes NumTabMto

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
Dim Cad As String
Dim Rs As ADODB.Recordset
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
    
    'miramos si hay otros campos con la misma ubicacion
    If b And (Modo = 3 Or Modo = 4) Then
        ' select que utilizamos para ello:
        'select count(*)
        'From rcampos
        'Where rcampos.fecbajas Is Null
        'and rcampos.poligono = 1
        'and rcampos.parcela = 1
        'and rcampos.recintos = 1
        'and rcampos.codcampo <> 5
        'and rcampos.codparti in
        '(select codparti from rpartida where codpobla in (select codpobla from rpartida where codparti = 8));
        Sql = "select count(*) "
        Sql = Sql & " from rcampos "
        Sql = Sql & " where rcampos.fecbajas is null "
        Sql = Sql & " and rcampos.poligono = " & DBSet(Text1(13), "N")
        Sql = Sql & " and rcampos.parcela = " & DBSet(Text1(14), "N")
        Sql = Sql & " and rcampos.recintos = " & DBSet(Text1(16), "N")
        Sql = Sql & " and rcampos.codcampo <> " & DBSet(Text1(0), "N")
        
        '[Monica]25/09/2012: para escalona tenemos que mirar la subparcela tambien
        If vParamAplic.Cooperativa = 10 Then
            Sql = Sql & " and rcampos.subparce = " & DBSet(Text1(15).Text, "T")
        End If
        
        Sql = Sql & " and rcampos.codparti in (select codparti from rpartida where "
        Sql = Sql & " codpobla in (select codpobla from rpartida where codparti = " & DBSet(Text1(3), "N") & "))"
    
        If TotalRegistros(Sql) <> 0 Then
            Sql = "Existe otra parcela dada de alta en la misma ubicación. " & vbCrLf & vbCrLf
            Sql = Sql & "                     ¿ Desea continuar ?"
            
            If MsgBox(Sql, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                 b = False
            End If
        End If
        
        ' comprobamos que el socio no esté dado de baja
        If b Then
            Sql = "select fechabaja from rsocios where codsocio = " & DBSet(Text1(1).Text, "N")
            
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If DBLet(Rs.Fields(0).Value, "F") <> "" Then
                Cad = "Este socio tiene fecha de baja. ¿ Desea continuar ?"
                If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                    b = False
                End If
            End If
        End If
        
        '[Monica]19/09/2012: introducimos la referencia catastral en el campo
        ' y damos aviso si existe una parcela con la misma referencia catastral
        If b Then
            ' si han metido algun valor en el campo
            If Text1(41).Text <> "" Then
                Sql = "select count(*) "
                Sql = Sql & " from rcampos "
                Sql = Sql & " where rcampos.fecbajas is null and refercatas = " & DBSet(Trim(Text1(41).Text), "T")
                Sql = Sql & " and rcampos.codcampo <> " & DBSet(Text1(0), "N")
            
                If TotalRegistros(Sql) <> 0 Then
                    Sql = "Existe otra parcela con la misma Referencia Catastral. " & vbCrLf & vbCrLf
                    Sql = Sql & "                     ¿ Desea continuar ?"
                    
                    If MsgBox(Sql, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                         b = False
                    End If
                End If
            End If
        End If
        
        '[Monica]31/10/2014: si la fecha de alta es superior a la fecha de alta del socio de la seccion de horto damos un aviso
        If b Then
            Sql = "select fecalta from rsocios_seccion where codsocio = " & DBSet(Text1(1).Text, "N") & " and codsecci = " & DBSet(vParamAplic.Seccionhorto, "N")
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not Rs.EOF Then
                If DBLet(Rs.Fields(0).Value, "F") > CDate(Text1(10).Text) Then
                    Sql = "La fecha de alta del socio en la Seccion de Horto es superior a la fecha de alta del campo." & vbCrLf & vbCrLf
                    Sql = Sql & "                     ¿ Desea continuar ?"
                    
                    If MsgBox(Sql, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                        b = False
                    End If
                End If
            Else
                Sql = "El socio no se encuentra en la Sección de Horto." & vbCrLf & vbCrLf
                Sql = Sql & "                     ¿ Desea continuar ?"
                If MsgBox(Sql, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                    b = False
                End If
            End If
            Set Rs = Nothing
        End If
        
    End If
    
    
    If b And Modo = 4 Then
        If CInt(Text1(2).Text) <> CInt(VarieAnt) Then
            Sql = "select count(*) from rcampos_clasif where codcampo = " & DBSet(Text1(0).Text, "N")
            If TotalRegistros(Sql) <> 0 Then
                Cad = "Si se modifica la variedad, perderá la clasificación." & vbCrLf
                Cad = Cad & "               ¿ Desea continuar ?"
                
                If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                    b = False
                End If
            End If
        End If
        
        '[Monica]08/04/2015: si es alzira y hay entradas de esa variedad para el campo no dejamos cambiarlo
        If b And (vParamAplic.Cooperativa = 4) And CInt(Text1(2).Text) <> CInt(VarieAnt) Then
            Sql = "select count(*) from rentradas where codcampo = " & DBSet(Text1(0).Text, "N") & " and codvarie = " & DBSet(VarieAnt, "N")
            If TotalRegistros(Sql) <> 0 Then
                MsgBox "Existen entradas en báscula para este campo." & vbCrLf & "Debería crear otro campo.", vbExclamation
                b = False
            Else
                Sql = "select count(*) from rclasifica where codcampo = " & DBSet(Text1(0).Text, "N") & " and codvarie = " & DBSet(VarieAnt, "N")
                If TotalRegistros(Sql) <> 0 Then
                    MsgBox "Existen entradas en clasificacion para este campo." & vbCrLf & "Debería crear otro campo.", vbExclamation
                    b = False
                Else
                    Sql = "select count(*) from rhisfruta where codcampo = " & DBSet(Text1(0).Text, "N") & " and codvarie = " & DBSet(VarieAnt, "N")
                    If TotalRegistros(Sql) <> 0 Then
                        MsgBox "Existen entradas en el histórico para este campo." & vbCrLf & "Debería crear otro campo.", vbExclamation
                        b = False
                    End If
                End If
            End If
        End If
        
        If b And (CInt(Text1(2).Text) <> CInt(VarieAnt) Or CLng(Text1(1).Text) <> CLng(SocioAnt)) Then
            If HayEntradasCampoSocioVariedad(Text1(0).Text, SocioAnt, VarieAnt) Then
                Cad = "Exiten entradas del campo, para el socio, variedad anterior que se modificarán." & vbCrLf & vbCrLf
                Cad = Cad & "                ¿ Desea continuar ?  "
                If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then
                    b = False
                End If
            End If
            
            If b And HayAnticiposPdtesCampoSocioVariedad(Text1(0).Text, SocioAnt, VarieAnt) Then
                Cad = "Exiten anticipos pendientes de descontar del campo, para el socio, variedad anterior. "
                Cad = Cad & " Si posteriormente liquida no se descontarán. " & vbCrLf & vbCrLf
                Cad = Cad & "       ¿ Seguro que desea continuar ?  "
            
                If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                    b = False
                End If
            End If
        End If
        
    End If
    
    '[Monica]29/09/2014: no podemos poner de baja un campo que tenga fecha de alta en el programa operativo.
    '                    Lo utiliza CATADAU
    If b And (Modo = 3 Or Modo = 4) Then
        If Text1(11).Text <> "" And Text1(43).Text <> "" Then
            MsgBox "No se puede dar de baja un campo que tenga fecha de alta en el Programa Operativo, ni dar de alta en dicho Programa si el campo está dado de baja. Revise.", vbExclamation
            b = False
        End If
    End If
    
    
    ' ************************************************************************************
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    Cad = "(codcampo=" & Text1(0).Text & ")"
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
    vWhere = " WHERE codcampo=" & Data1.Recordset!codcampo
        ' ***********************************************************************
        
    ' ***** elimina les llínies ****
    conn.Execute "DELETE FROM rcampos_clasif " & vWhere

    conn.Execute "DELETE FROM rcampos_parcelas " & vWhere

    conn.Execute "DELETE FROM rcampos_seguros " & vWhere

    conn.Execute "DELETE FROM rcampos_hco " & vWhere
    
    conn.Execute "DELETE FROM rcampos_gastos " & vWhere

    conn.Execute "DELETE FROM rcampos_revision " & vWhere

'[Monica]30/10/2013: no dejaba borrar una parcela
'    conn.Execute "DELETE FROM rcampos_ordrec " & vWhere


'    Conn.Execute "DELETE FROM rsocios_telefonos " & vWhere
'
'    ' *******************************
'

    CargarUnCampo CLng(Data1.Recordset!codcampo), "D"


'    'Eliminar la CAPÇALERA
'    vWhere = " WHERE codsocio=" & Data1.Recordset!codsocio
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
    Select Case Index
        Case 0 'cod campo
            PonerFormatoEntero Text1(0)

        Case 1 'SOCIO
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rsocios", "nomsocio")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Socio: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmSoc = New frmManSocios
                        frmSoc.DatosADevolverBusqueda = "0|1|"
'                        frmVar.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmSoc.Show vbModal
                        Set frmSoc = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
            ' si estamos insertando ponemos el propietario iguasl que el socio
            If Modo = 3 Then
                Text1(22).Text = Text1(Index).Text
                Text2(22).Text = Text2(Index).Text
            End If
        
        Case 22 'PROPIETARIO
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rsocios", "nomsocio", "codsocio", "N")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Socio: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmSoc = New frmManSocios
                        frmSoc.DatosADevolverBusqueda = "0|1|"
'                        frmVar.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmSoc.Show vbModal
                        Set frmSoc = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        Case 2 'VARIEDAD
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "variedades", "nomvarie")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Variedad: " & Text1(Index).Text & vbCrLf
                    MsgBox cadMen, vbExclamation
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
                
        Case 3 'PARTIDA
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rpartida", "nomparti", "codparti", "N")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Partida: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmPar = New frmManPartidas
                        frmPar.DatosADevolverBusqueda = "0|1|"
                        frmPar.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmPar.Show vbModal
                        Set frmPar = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                Else
                    PonerDatosPartida
                End If
            Else
                Text2(Index).Text = ""
            End If
                
        Case 5, 6, 7, 33 'superficies en hectareas
            If Modo = 1 Then Exit Sub
            If Text1(Index).Text <> "" Then
                PonerFormatoDecimal Text1(Index), 7
                Text2(Index).Text = Round2(ImporteSinFormato(Text1(Index).Text) / vParamAplic.Faneca, 2)
                PonerFormatoDecimal Text2(Index), 3
                
                If Index = 5 Then
                    If ComprobarCero(Text1(5).Text) = 0 Then
                        EstablecerOrden False 'por hanegadas
                        
                        PonerFoco Text2(Index)
                    Else
                        EstablecerOrden True 'por hectareas
                        PonerFoco Text1(6)
                    End If
                End If
            Else
                If Index = 5 Then
                    EstablecerOrden False 'por hanegadas
                    
                    PonerFoco Text2(Index)
                End If
            End If
            
                
        Case 8, 9 'aforo, arboles
            PonerFormatoEntero Text1(Index)
                
        Case 12 'SITUACION Campo
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rsituacioncampo", "nomsitua")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Situación Campo: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmSit = New frmManSituCamp
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
            
        Case 13, 14 'poligono y parcela
            PonerFormatoEntero Text1(Index)
        
        Case 16 'recinto
            PonerFormatoEntero Text1(Index)
            
        Case 17, 18 'año de plantacion y OID
            PonerFormatoEntero Text1(Index)
            
        Case 19, 20 ' longitud y latitud
            'PonerFormatoDecimal Text1(Index), 9
            If Text1(Index).Text <> "" Then
                Text1(Index).Text = Format(TransformaPuntosComas(Text1(Index).Text), "#0.000000")
            End If
                
                
        '[Monica]29/09/2014: campo 43, fecha de alta en programa operativo
        Case 10, 11, 42, 43 'Fecha no comprobaremos que esté dentro de campaña
                    'Fecha de alta y fecha baja
            If Modo = 1 Then Exit Sub
            PonerFormatoFecha Text1(Index), False
            
            
        Case 23 'nro de campo scampo en Multibase
            PonerFormatoEntero Text1(Index)
            
            
        Case 24 'Responsable
            If Modo = 1 Then Exit Sub
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rcapataz", "nomcapat")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Responsable: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmRes = New frmManCapataz
                        frmRes.DatosADevolverBusqueda = "0|1|"
                        frmRes.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmRes.Show vbModal
                        Set frmRes = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 25 'Marco de Plantacion
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rplantacion", "nomplanta")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el marco de Plantación: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmPlan = New frmManPlantacion
                        frmPlan.DatosADevolverBusqueda = "0|1|"
                        frmPlan.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmPlan.Show vbModal
                        Set frmPlan = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
            
        Case 26 'Desarrollo vegetativo
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rdesarrollo", "nomdesa")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Desarrollo Vegetativo: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmDesa = New frmManDesarrollo
                        frmDesa.DatosADevolverBusqueda = "0|1|"
                        frmDesa.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmDesa.Show vbModal
                        Set frmDesa = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 27 ' Tipo de tierra
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rtierra", "nomtierra")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Tipo de Tierra: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmTie = New frmManTierra
                        frmTie.DatosADevolverBusqueda = "0|1|"
                        frmTie.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmTie.Show vbModal
                        Set frmTie = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 28, 35 ' importe coste seguro
            PonerFormatoDecimal Text1(Index), 1
            
        Case 29 'opcion seguro
            If Modo = 1 Then Exit Sub
'            If PonerFormatoEntero(Text1(Index)) Then
            If Text1(Index).Text <> "" Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rseguroopcion", "nomseguro", "codseguro", "T")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Opción Seguro: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmSegOp = New frmManSeguroOpc
                        frmSegOp.DatosADevolverBusqueda = "0|1|"
                        frmSegOp.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmSegOp.Show vbModal
                        Set frmSegOp = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 34 'opcion seguro
            If Modo = 1 Then Exit Sub
'            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rseguroopcion", "nomseguro", "codseguro", "T")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Opción Seguro: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmSegOp = New frmManSeguroOpc
                        frmSegOp.DatosADevolverBusqueda = "0|1|"
                        frmSegOp.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmSegOp.Show vbModal
                        Set frmSegOp = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
'            Else
'                Text2(Index).Text = ""
'            End If
            
            
            
        Case 30 'procedencia de riego
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rproceriego", "nomproce")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Procedencia de Riego: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmProc = New frmManProceRiego
                        frmProc.DatosADevolverBusqueda = "0|1|"
                        frmProc.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmProc.Show vbModal
                        Set frmProc = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 31 'patron a pie
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rpatronpie", "nompatron")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Patrón Pie: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmPat = New frmManPatronaPie
                        frmPat.DatosADevolverBusqueda = "0|1|"
                        frmPat.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmPat.Show vbModal
                        Set frmPat = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 32 ' nro de llave
            PonerFormatoEntero Text1(Index)
            
        '[Monica]06/05/2013: Faltaba sacar la descripcion de la zona
        Case 37 ' codigo de zona
            If PonerFormatoEntero(Text1(Index)) Then
                Text4(3).Text = DevuelveDesdeBDNew(cAgro, "rzonas", "nomzonas", "codzonas", Text1(Index).Text, "N")
            End If
            
        Case 38 ' globalgap
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
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 39 ' codigo de cliente
            If vParamAplic.BDAriges <> "" Then
                If Text1(39).Text <> "" Then
                    Text2(39).Text = DevuelveDesdeBDNew(cAriges, "sclien", "nomclien", "codclien", Text1(39).Text, "N")
                End If
            End If
            
        '[Monica]03/02/2012: cambio en los anticipos/liqudiaciones para Picassent
        Case 40 ' % Comision sobre precio de anticipo/liquidacion
            PonerFormatoDecimal Text1(Index), 4
            
        '[Monica]26/09/2016: para coopic
        Case 44, 45, 46, 47 ' puntos
            PonerFormatoEntero Text1(Index)
            
    End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 5: KEYBusqueda KeyAscii, 0 'poblacion
                Case 7: KEYBusqueda KeyAscii, 1 'actividad
                Case 8: KEYBusqueda KeyAscii, 2 'grupo
                Case 24: KEYBusqueda KeyAscii, 5 'responsable
                Case 25: KEYBusqueda KeyAscii, 6 'marco de plantacion
                Case 26: KEYBusqueda KeyAscii, 8 'desrrollo vegetativo
                Case 27: KEYBusqueda KeyAscii, 7 'tipo de tierra
                Case 30: KEYBusqueda KeyAscii, 10 'procedencia de riego
                Case 31: KEYBusqueda KeyAscii, 11 'patron pie
                Case 29: KEYBusqueda KeyAscii, 9 'seguro opcion
                Case 38: KEYBusqueda KeyAscii, 14 'codigo de globalgap
            End Select
        End If
    Else
        If Index <> 21 Then KEYpress KeyAscii
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me

End Sub


Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

' **** si n'hi han camps de descripció a la capçalera ****
Private Sub PosarDescripcions()
Dim NomEmple As String

    On Error GoTo EPosarDescripcions

    Text2(1).Text = PonerNombreDeCod(Text1(1), "rsocios", "nomsocio", "codsocio", "N")
    Text2(22).Text = PonerNombreDeCod(Text1(22), "rsocios", "nomsocio", "codsocio", "N")
    Text2(12).Text = PonerNombreDeCod(Text1(12), "rsituacioncampo", "nomsitua", "codsitua", "N")
    Text2(2).Text = PonerNombreDeCod(Text1(2), "variedades", "nomvarie", "codvarie", "N")
    Text2(3).Text = PonerNombreDeCod(Text1(3), "rpartida", "nomparti", "codparti", "N")
    Text2(38).Text = PonerNombreDeCod(Text1(38), "rglobalgap", "descripcion", "codigo", "T")
    
    '[Monica]14/02/2013: sacamos el codigo de conselleria de las lineas
    txtAux2(13).Text = DevuelveValor("select codconse from variedades where codvarie = " & DBSet(Text1(2).Text, "N"))
    
    
    If vParamAplic.BDAriges <> "" Then
        Text2(39).Text = DevuelveDesdeBDNew(cAriges, "sclien", "nomclien", "codclien", Text1(39).Text, "N")
    End If
    
    PonerDatosPartida
    
    If Text1(5).Text <> "" Then
        Text2(5).Text = Round2(ImporteSinFormato(Text1(5).Text) / vParamAplic.Faneca, 2)
        PonerFormatoDecimal Text2(5), 3
    End If
    
    If Text1(6).Text <> "" Then
        Text2(6).Text = Round2(ImporteSinFormato(Text1(6).Text) / vParamAplic.Faneca, 2)
        PonerFormatoDecimal Text2(6), 3
    End If
    
    If Text1(7).Text <> "" Then
        Text2(7).Text = Round2(ImporteSinFormato(Text1(7).Text) / vParamAplic.Faneca, 2)
        PonerFormatoDecimal Text2(7), 3
    End If
    
    If Text1(33).Text <> "" Then
        Text2(33).Text = Round2(ImporteSinFormato(Text1(33).Text) / vParamAplic.Faneca, 2)
        PonerFormatoDecimal Text2(33), 3
    End If
    
    Text2(24).Text = PonerNombreDeCod(Text1(24), "rcapataz", "nomcapat", "codcapat", "N")
    Text2(25).Text = PonerNombreDeCod(Text1(25), "rplantacion", "nomplanta", "codplanta", "N")
    Text2(26).Text = PonerNombreDeCod(Text1(26), "rdesarrollo", "nomdesa", "coddesa", "N")
    Text2(27).Text = PonerNombreDeCod(Text1(27), "rtierra", "nomtierra", "codtierra", "N")
    Text2(29).Text = PonerNombreDeCod(Text1(29), "rseguroopcion", "nomseguro", "codseguro", "T")
    Text2(30).Text = PonerNombreDeCod(Text1(30), "rproceriego", "nomproce", "codproce", "N")
    Text2(31).Text = PonerNombreDeCod(Text1(31), "rpatronpie", "nompatron", "codpatron", "N")
    
    '[Monica]19/09/2011: si el campo está regado por un hidrante que aparezca el hidrante
    Text5(0).Text = DevuelveDesdeBDNew(cAgro, "rpozos_campos", "hidrante", "codcampo", Text1(0).Text, "N")
    
    
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
        Case 4
            If Index = 7 Then
                BotonImpresionRevisiones
            Else
                BotonContabilizarGasto
            End If
        Case Else
    End Select
    'End If
End Sub

Private Sub BotonImpresionRevisiones()

    Screen.MousePointer = vbHourglass
    
    frmListado.OpcionListado = 45
    frmListado.NumCod = "rcampos_revision.codcampo = " & Me.Data1.Recordset!codcampo
    frmListado.Show vbModal
    
    Screen.MousePointer = vbDefault

End Sub




Private Sub BotonContabilizarGasto()

    Screen.MousePointer = vbHourglass
    
    frmListado.OpcionListado = 37
    frmListado.NumCod = "rcampos_gastos.codcampo = " & AdoAux(5).Recordset!codcampo & " and rcampos_gastos.numlinea = " & AdoAux(5).Recordset!numlinea
    frmListado.Show vbModal
    CargaGrid NumTabMto, True
    
    Screen.MousePointer = vbDefault

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
        Case 0 'coopropietarios
            Sql = "¿Seguro que desea eliminar el coopropietario?"
            Sql = Sql & vbCrLf & "Coopropietario: " & AdoAux(Index).Recordset!Codsocio & " - " & AdoAux(Index).Recordset!nomsocio
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                Sql = "DELETE FROM rcampos_cooprop"
                Sql = Sql & " WHERE rcampos_cooprop.codcampo = " & DBLet(AdoAux(Index).Recordset!codcampo, "N")
                Sql = Sql & " and codsocio = " & DBLet(AdoAux(Index).Recordset!Codsocio, "N")
            End If
        
        Case 1 'clasificacion
            Sql = "¿Seguro que desea eliminar la clasificación?"
            Sql = Sql & vbCrLf & "Clasificación: " & AdoAux(Index).Recordset!codcalid & " - " & AdoAux(Index).Recordset!nomcalid
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                Sql = "DELETE FROM rcampos_clasif"
                Sql = Sql & vWhere & " AND codvarie= " & DBLet(AdoAux(Index).Recordset!codvarie, "N")
                Sql = Sql & " and codcalid = " & DBLet(AdoAux(Index).Recordset!codcalid, "N")
            End If
    
        Case 2 'parcelas
            vWhere = ObtenerWhereCab2(True)
            
            Sql = "¿Seguro que desea eliminar la parcela?"
            Sql = Sql & vbCrLf & "Póligono: " & AdoAux(Index).Recordset!Poligono & " - Parcela : " & AdoAux(Index).Recordset!Parcela
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                Sql = "DELETE FROM rcampos_parcelas"
                Sql = Sql & vWhere & " AND numlinea= " & DBLet(AdoAux(Index).Recordset!numlinea, "N")
            End If
    
        Case 3 'agroseguro
            vWhere = ObtenerWhereCab3(True)
            
            Sql = "¿Seguro que desea eliminar la Línea?"
            Sql = Sql & vbCrLf & "Fecha: " & AdoAux(Index).Recordset!Fecha & " - Incidencia : " & AdoAux(Index).Recordset!nomincid
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                Sql = "DELETE FROM rcampos_seguros"
                Sql = Sql & vWhere & " AND numlinea= " & DBLet(AdoAux(Index).Recordset!numlinea, "N")
            End If
    
        Case 4 'hco de campos
            vWhere = ObtenerWhereCab4(True)
            
            Sql = "¿Seguro que desea eliminar la Línea?" & vbCrLf
            Sql = Sql & "Socio: " & Format(AdoAux(Index).Recordset!Codsocio, "000000") & " - " & AdoAux(Index).Recordset!nomsocio
            Sql = Sql & vbCrLf & "Fecha Alta: " & AdoAux(Index).Recordset!FechaAlta
            Sql = Sql & vbCrLf & "Fecha Baja: " & AdoAux(Index).Recordset!FechaBaja
            Sql = Sql & vbCrLf & "Incidencia : " & AdoAux(Index).Recordset!nomincid
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                Sql = "DELETE FROM rcampos_hco"
                Sql = Sql & vWhere & " AND numlinea= " & DBLet(AdoAux(Index).Recordset!numlinea, "N")
            End If
    
        Case 5 'gastos de campos
            vWhere = ObtenerWhereCab5(True)
            
            If AdoAux(Index).Recordset!contabilizado Then
                Sql = "Este Gasto está contabilizado. Si continua deberá modificar la contabilidad." & vbCrLf
                Sql = Sql & " ¿ Desea continuar ? "
                If MsgBox(Sql, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            End If
            
            
            Sql = "¿Seguro que desea eliminar la Línea?" & vbCrLf
            Sql = Sql & "Concepto: " & Format(AdoAux(Index).Recordset!Codgasto, "00") & " - " & AdoAux(Index).Recordset!NomGasto
            Sql = Sql & vbCrLf & "Fecha: " & AdoAux(Index).Recordset!Fecha
            Sql = Sql & vbCrLf & "Importe: " & AdoAux(Index).Recordset!Importe
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                Sql = "DELETE FROM rcampos_gastos"
                Sql = Sql & vWhere & " AND numlinea= " & DBLet(AdoAux(Index).Recordset!numlinea, "N")
            End If
        
        Case 6 'ordenes de recoleccion
            vWhere = ObtenerWhereCab6(True)
            
            
            Sql = "¿Seguro que desea eliminar la Línea?" & vbCrLf
            Sql = Sql & "Orden: " & Format(AdoAux(Index).Recordset!nroorden, "0000000")
            Sql = Sql & vbCrLf & "Fecha: " & AdoAux(Index).Recordset!fecimpre
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                Sql = "DELETE FROM rcampos_ordrec"
                Sql = Sql & vWhere & " AND nroorden= " & DBLet(AdoAux(Index).Recordset!nroorden, "N")
            End If
    
        Case 7 'revisiones de campos
            vWhere = ObtenerWhereCab7(True)
            
            
            Sql = "¿Seguro que desea eliminar la Línea?" & vbCrLf
            Sql = Sql & "Fecha: " & AdoAux(Index).Recordset!Fecha
            Sql = Sql & vbCrLf & "Técnico: " & AdoAux(Index).Recordset!tecnico
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                Sql = "DELETE FROM rcampos_revision"
                Sql = Sql & vWhere & " AND numlinea= " & DBLet(AdoAux(Index).Recordset!numlinea, "N")
            End If
    
    
    End Select

    If eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        conn.Execute Sql
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
'        If Index <> 3 Then
            CargaGrid Index, True
        ' ***************************************************
        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
            
        End If
        SumaTotalPorcentajes NumTabMto
        ' *** si n'hi han tabs sense datagrid ***
'        If Index = 3 Then CargaFrame 3, True
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
        Case 0: vTabla = "rcampos_cooprop"
        Case 1: vTabla = "rcampos_clasif"
        Case 2: vTabla = "rcampos_parcelas"
        Case 3: vTabla = "rcampos_seguros"
        Case 4: vTabla = "rcampos_hco"
        Case 5: vTabla = "rcampos_gastos"
        Case 6: vTabla = "rcampos_ordrec"
        Case 7: vTabla = "rcampos_revision"
    End Select
    ' ********************************************************

    vWhere = ObtenerWhereCab(False)

    Select Case Index
         Case 0, 1, 2, 3, 4, 5, 6, 7 'clasificacion
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            Select Case Index
                Case 0
                    NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", "rcampos_cooprop.codcampo = " & Val(Text1(0).Text))
                Case 1
                    NumF = ""
                Case 2
                    NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", "rcampos_parcelas.codcampo = " & Val(Text1(0).Text))
                Case 3
                    NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", "rcampos_seguros.codcampo = " & Val(Text1(0).Text))
                Case 4
                    NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", "rcampos_hco.codcampo = " & Val(Text1(0).Text))
                Case 5
                    NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", "rcampos_gastos.codcampo = " & Val(Text1(0).Text))
                Case 7
                    NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", "rcampos_revision.codcampo = " & Val(Text1(0).Text))
            End Select
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
                    For i = 0 To txtAux1.Count - 1
                        txtAux1(i).Text = ""
                    Next i
                    txtAux1(0).Text = Text1(0).Text 'codcampo
                    txtAux1(2).Text = Text1(2).Text 'codvariedad
                    txtAux2(1).Text = ""
                    PonerFoco txtAux1(1)

                Case 0 'copropietarios
                    For i = 0 To txtAux3.Count - 1
                        txtAux3(i).Text = ""
                    Next i
                    txtAux2(0).Text = ""
                    txtAux3(0).Text = Text1(0).Text 'codcampo
                    txtAux3(1).Text = NumF 'numlinea
                    txtAux3(2).Text = ""
                    PonerFoco txtAux3(2)
                
                Case 2 ' parcelas
                    For i = 0 To txtAux4.Count - 1
                        txtAux4(i).Text = ""
                    Next i
                    txtAux4(0).Text = Text1(0).Text 'codcampo
                    txtAux4(1).Text = NumF 'numlinea
                    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then txtAux4(6).Text = "0"
                    PonerFoco txtAux4(2)
                
                Case 3 ' seguros
                    For i = 0 To txtAux5.Count - 1
                        txtAux5(i).Text = ""
                    Next i
                    txtAux2(2).Text = ""
                    
                    txtAux5(0).Text = Text1(0).Text 'codcampo
                    txtAux5(1).Text = NumF 'numlinea
                    PonerFoco txtAux5(2)
                
                    Me.chkAux(0).Value = 0
                
                Case 4 ' hco de campos
                    For i = 0 To txtaux6.Count - 1
                        txtaux6(i).Text = ""
                    Next i
                    txtaux6(0).Text = Text1(0).Text 'codcampo
                    txtaux6(1).Text = NumF 'numlinea
                    PonerFoco txtaux6(2)
                
                Case 5 ' gastos de  campos
                    For i = 0 To txtAux7.Count - 1
                        txtAux7(i).Text = ""
                    Next i
                    txtAux2(5).Text = ""
                    txtAux7(0).Text = Text1(0).Text 'codcampo
                    txtAux7(1).Text = NumF 'numlinea
                    PonerFoco txtAux7(2)
                    
                    Me.chkAux(1).Value = 0
                
                Case 6 ' ordenes de recoleccion de campos
                    For i = 0 To txtAux8.Count - 1
                        txtAux8(i).Text = ""
                    Next i
                    txtAux8(0).Text = Text1(0).Text 'codcampo
                    PonerFoco txtAux8(1)
                    
                Case 7 ' revisiones de campo
                    For i = 0 To txtAux9.Count - 1
                        txtAux9(i).Text = ""
                    Next i
                    txtAux9(0).Text = Text1(0).Text 'codcampo
                    txtAux9(1).Text = NumF 'numlinea
                    txtAux9(2).Text = Now ' fecha
                    PonerFoco txtAux9(2)
                
            End Select


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
    Dim j As Integer
    Dim Sql As String
    

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub

    If Index = 5 Then
        If CInt(AdoAux(Index).Recordset!contabilizado) = 1 Then
            Sql = "Este Gasto está contabilizado, deberá modificar la contabilidad." & vbCrLf
            Sql = Sql & " ¿ Desea continuar ? "
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
    End If
    
    
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
        Case 0, 1, 2, 3, 4, 5, 6, 7 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
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
        Case 0 'coopropietarios
            txtAux3(0).Text = DataGridAux(Index).Columns(0).Text
            txtAux3(1).Text = DataGridAux(Index).Columns(1).Text
            txtAux3(2).Text = DataGridAux(Index).Columns(2).Text
            
            txtAux2(0).Text = DataGridAux(Index).Columns(3).Text
            txtAux3(3).Text = DataGridAux(Index).Columns(4).Text
        
        Case 1 'clasificacion
            txtAux1(0).Text = DataGridAux(Index).Columns(0).Text
            txtAux1(1).Text = DataGridAux(Index).Columns(2).Text
            txtAux1(2).Text = DataGridAux(Index).Columns(1).Text
            
            txtAux2(1).Text = DataGridAux(Index).Columns(3).Text
            txtAux1(3).Text = DataGridAux(Index).Columns(4).Text
    
        Case 2 'parcelas
            For i = 0 To 10
                txtAux4(i).Text = DataGridAux(Index).Columns(i).Text
            Next i
    
        Case 3 'seguros
            For i = 0 To 3
                txtAux5(i).Text = DataGridAux(Index).Columns(i).Text
            Next i
            txtAux2(2).Text = DataGridAux(Index).Columns(4).Text
            '[Monica]26/01/2016: añadida nueva columna de kilos de aportacion
'            For I = 4 To 6
'                txtAux5(I).Text = DataGridAux(Index).Columns(I + 1).Text
'            Next I
            txtAux5(4).Text = DataGridAux(Index).Columns(5).Text
            txtAux5(7).Text = DataGridAux(Index).Columns(6).Text
            txtAux5(5).Text = DataGridAux(Index).Columns(7).Text
            txtAux5(6).Text = DataGridAux(Index).Columns(8).Text
        
            Me.chkAux(0).Value = Me.AdoAux(3).Recordset!essiniestro
        
        
        Case 4 'hco de campos
            For i = 0 To 2
                txtaux6(i).Text = DataGridAux(Index).Columns(i).Text
            Next i
            txtAux2(4).Text = DataGridAux(Index).Columns(3).Text
            For i = 3 To 5
                txtaux6(i).Text = DataGridAux(Index).Columns(i + 1).Text
            Next i
            txtAux2(3).Text = DataGridAux(Index).Columns(7).Text
        
        Case 5 'gastos de campos
            For i = 0 To 2
                txtAux7(i).Text = DataGridAux(Index).Columns(i).Text
            Next i
            txtAux2(5).Text = DataGridAux(Index).Columns(3).Text
            For i = 3 To 4
                txtAux7(i).Text = DataGridAux(Index).Columns(i + 1).Text
            Next i
            chkAux(1).Value = DataGridAux(Index).Columns(6).Text
        
        Case 6 'ordenes de recoleccion de campos
            For i = 0 To 2
                txtAux8(i).Text = DataGridAux(Index).Columns(i).Text
            Next i
            
        Case 7 'revisiones de campo
            For i = 0 To txtAux9.Count - 1
                txtAux9(i).Text = DataGridAux(Index).Columns(i).Text
            Next i
            
    End Select

    LLamaLineas Index, ModoLineas, anc

    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'coopropietarios
            PonerFoco txtAux3(2)
        Case 1 'clasificacion
            PonerFoco txtAux1(3)
        Case 2 'parcelas
            PonerFoco txtAux4(2)
        Case 3 'agroseguro
            PonerFoco txtAux5(2)
        Case 4 'hco
            PonerFoco txtaux6(2)
        Case 5 'gastos de campos
            PonerFoco txtAux7(2)
        Case 6 'ordenes de recoleccion de campos
            PonerFoco txtAux8(1)
        Case 7 'revisiones de campo
            PonerFoco txtAux9(2)
    End Select
    ' ***************************************************************************************
End Sub


Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    If Index <> 3 Then DeseleccionaGrid DataGridAux(Index)
    ' ***************************************************

    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 1 'clasificacion
            For jj = 1 To txtAux1.Count - 1
                If jj <> 2 Then
                    txtAux1(jj).visible = b
                    txtAux1(jj).Top = alto
                End If
            Next jj
            
            txtAux2(1).visible = b
            txtAux2(1).Top = alto

            For jj = 1 To 1
                cmdAux(jj).visible = b
                cmdAux(jj).Top = txtAux1(3).Top
                cmdAux(jj).Height = txtAux1(3).Height
            Next jj
        Case 0 ' coopropietarios
            For jj = 2 To txtAux3.Count - 1
                txtAux3(jj).visible = b
                txtAux3(jj).Top = alto
            Next jj
            txtAux2(0).visible = b
            txtAux2(0).Top = alto
            cmdAux(0).visible = b
            cmdAux(0).Top = txtAux3(2).Top
            cmdAux(0).Height = txtAux3(2).Height
    
        Case 2 'parcelas
            For jj = 2 To txtAux4.Count - 1
                If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
                    If jj <> 6 Then
                        txtAux4(jj).visible = b
                        txtAux4(jj).Top = alto
                    End If
                Else
                    txtAux4(jj).visible = b
                    txtAux4(jj).Top = alto
                End If
            Next jj
        
        Case 3 'seguros
            For jj = 2 To txtAux5.Count - 1
                txtAux5(jj).visible = b
                txtAux5(jj).Top = alto
            Next jj
            txtAux2(2).visible = b
            txtAux2(2).Top = alto
            
            For jj = 2 To 4
                cmdAux(jj).visible = b
                cmdAux(jj).Top = txtAux5(3).Top
                cmdAux(jj).Height = txtAux5(3).Height
            Next jj
            
            chkAux(0).visible = b
            chkAux(0).Top = txtAux5(3).Top
            chkAux(0).Height = txtAux5(3).Height
            
        Case 4 'hco de campos
            For jj = 2 To txtaux6.Count - 1
                txtaux6(jj).visible = b
                txtaux6(jj).Top = alto
            Next jj
            txtAux2(3).visible = b
            txtAux2(3).Top = alto
            txtAux2(4).visible = b
            txtAux2(4).Top = alto
            
            For jj = 5 To 8
                cmdAux(jj).visible = b
                cmdAux(jj).Top = txtaux6(2).Top
                cmdAux(jj).Height = txtaux6(2).Height
            Next jj
            
        Case 5 'gastos de campos
            For jj = 2 To txtAux7.Count - 1
                txtAux7(jj).visible = b
                txtAux7(jj).Top = alto
            Next jj
            txtAux2(5).visible = b
            txtAux2(5).Top = alto
            
            For jj = 9 To 10
                cmdAux(jj).visible = b
                cmdAux(jj).Top = txtAux7(2).Top
                cmdAux(jj).Height = txtAux7(2).Height
            Next jj
            
        Case 6 'ordenes de recoleccion
            For jj = 1 To txtAux8.Count - 1
                txtAux8(jj).visible = b
                txtAux8(jj).Top = alto
            Next jj
            
            For jj = 11 To 11
                cmdAux(jj).visible = b
                cmdAux(jj).Top = txtAux8(2).Top
                cmdAux(jj).Height = txtAux8(2).Height
            Next jj
            
        Case 7 'revisiones de campos
            For jj = 2 To txtAux9.Count - 1
                txtAux9(jj).visible = b
                txtAux9(jj).Top = alto
            Next jj
            
            For jj = 12 To 12
                cmdAux(jj).visible = b
                cmdAux(jj).Top = txtAux9(2).Top
                cmdAux(jj).Height = txtAux9(2).Height
            Next jj
            
            
    End Select
End Sub


Private Sub TxtAux3_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

    If Not PerderFocoGnral(txtAux3(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 2 'NIF
            If PonerFormatoEntero(txtAux3(Index)) Then
                txtAux2(0).Text = PonerNombreDeCod(txtAux3(Index), "rsocios", "nomsocio")
                If txtAux2(0).Text = "" Then
                    cadMen = "No existe el Socio: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmSoc1 = New frmManSocios
                        frmSoc1.DatosADevolverBusqueda = "0|1|"
'                        frmVar.NuevoCodigo = Text1(Index).Text
                        txtAux3(Index).Text = ""
                        TerminaBloquear
                        frmSoc1.Show vbModal
                        Set frmSoc1 = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtAux3(Index).Text = ""
                    End If
                    PonerFoco txtAux3(Index)
                Else
                    ' comprobamos que el socio no esté dado de baja
                    If Not EstaSocioDeAlta(txtAux3(Index).Text) Then
                        If MsgBox("Este socio tiene fecha de baja. ¿ Continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                            txtAux3(Index).Text = ""
                            txtAux2(0).Text = ""
                            PonerFoco txtAux3(Index)
                        End If
                    End If
                End If
            Else
                txtAux2(0).Text = ""
            End If
            
        Case 3 'porcentaje de
            PonerFormatoDecimal txtAux3(Index), 4
            If txtAux3(2).Text <> "" Then cmdAceptar.SetFocus
    
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
    
    If b And (Modo = 5 And ModoLineas = 1) And nomframe = "FrameAux1" Then  'insertar
        'comprobar si existe ya el cod. de la calidad para ese campo
        Sql = ""
        Sql = DevuelveDesdeBDNew(cAgro, "rcampos_clasif", "codcalid", "codcampo", txtAux1(0).Text, "N", , "codvarie", txtAux1(2).Text, "N", "codcalid", txtAux1(1).Text, "N")
        If Sql <> "" Then
            MsgBox "Ya existe la calidad para el campo.", vbExclamation
            PonerFoco txtAux1(1)
            b = False
        End If
    End If
    
    If b And (Modo = 5 And ModoLineas = 1) And nomframe = "FrameAux0" Then  'insertar
        'comprobar que el porcentaje sea distinto de cero
        If txtAux3(3).Text = "" Then
            MsgBox "El porcentaje de coopropiedad debe ser superior a 0.", vbExclamation
            PonerFoco txtAux3(3)
            b = False
        Else
            If CInt(txtAux3(3).Text) = 0 Then
                MsgBox "El porcentaje de coopropiedad debe ser superior a 0.", vbExclamation
                PonerFoco txtAux3(3)
                b = False
            End If
        End If
    End If
    
    
    
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


Private Function ActualisaCtaprpal(ByRef numlinea As Integer)
Dim Sql As String
'yo
'    On Error Resume Next
'    'tot lo que no siga un SELECT no fa falta un Record Set
'    SQL = "UPDATE cltebanc SET ctaprpal = 0"
'    SQL = SQL & " WHERE codclien=" & text1(0).Text & " AND codempre= " & vSesion.Empresa & " AND numlinea<> " & numlinea
'    Conn.Execute SQL
'
'    If Err.Number <> 0 Then Err.Clear
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
            Set frmSit = New frmManSituCamp
            frmSit.DatosADevolverBusqueda = "0|1|"
            frmSit.CodigoActual = Text1(12).Text
            frmSit.Show vbModal
            Set frmSit = Nothing
            PonerFoco Text1(12)
        
       Case 1 'Socios
            indice = 1
            Set frmSoc = New frmManSocios
'            frmSoc.DeConsulta = True
            frmSoc.DatosADevolverBusqueda = "0|1|"
'            frmSoc.CodigoActual = Text1(1).Text
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco Text1(indice)
    
       Case 4 ' Propietario
            indice = 22
            Set frmSoc = New frmManSocios
'            frmSoc.DeConsulta = True
            frmSoc.DatosADevolverBusqueda = "0|1|"
'            frmSoc.CodigoActual = Text1(1).Text
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco Text1(indice)
    
       Case 2 'Variedades
            Set frmVar = New frmComVar
'            frmVar.DeConsulta = True
            frmVar.DatosADevolverBusqueda = "0|1|"
'            frmVar.CodigoActual = Text1(2).Text
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco Text1(2)
    
       Case 3 'Partidas
            Set frmPar = New frmManPartidas
            frmPar.DeConsulta = True
            frmPar.DatosADevolverBusqueda = "0|1|2|3|4|5|"
            frmPar.CodigoActual = Text1(3).Text
            frmPar.Show vbModal
            Set frmPar = Nothing
            PonerFoco Text1(3)
    
       Case 5 'Responsable
            Set frmRes = New frmManCapataz
            frmRes.DeConsulta = True
            frmRes.DatosADevolverBusqueda = "0|1|"
            frmRes.CodigoActual = Text1(24).Text
            frmRes.Show vbModal
            Set frmRes = Nothing
            PonerFoco Text1(24)
    
       Case 6 'marco de plantacion
            Set frmPlan = New frmManPlantacion
            frmPlan.DeConsulta = True
            frmPlan.DatosADevolverBusqueda = "0|1|"
            frmPlan.CodigoActual = Text1(25).Text
            frmPlan.Show vbModal
            Set frmPlan = Nothing
            PonerFoco Text1(25)
    
       Case 8 'desarrollo vegetativo
            Set frmDesa = New frmManDesarrollo
            frmDesa.DeConsulta = True
            frmDesa.DatosADevolverBusqueda = "0|1|"
            frmDesa.CodigoActual = Text1(26).Text
            frmDesa.Show vbModal
            Set frmDesa = Nothing
            PonerFoco Text1(26)
    
       Case 7 'tipo de tierra
            Set frmTie = New frmManTierra
            frmTie.DeConsulta = True
            frmTie.DatosADevolverBusqueda = "0|1|"
            frmTie.CodigoActual = Text1(27).Text
            frmTie.Show vbModal
            Set frmTie = Nothing
            PonerFoco Text1(27)
    
        Case 9, 12 ' Opcion seguro y campaña anterior
            Select Case Index
                Case 9
                    indCodigo = 29
                Case 12
                    indCodigo = 34
            End Select
            
            Set frmSegOp = New frmManSeguroOpc
            frmSegOp.DeConsulta = True
            frmSegOp.DatosADevolverBusqueda = "0|1|"
            frmSegOp.CodigoActual = Text1(indCodigo).Text
            frmSegOp.Show vbModal
            Set frmSegOp = Nothing
            PonerFoco Text1(indCodigo)
        
        Case 10 ' procedencia de riego
            Set frmProc = New frmManProceRiego
            frmProc.DeConsulta = True
            frmProc.DatosADevolverBusqueda = "0|1|"
            frmProc.CodigoActual = Text1(30).Text
            frmProc.Show vbModal
            Set frmProc = Nothing
            PonerFoco Text1(30)
    
        Case 11 ' patrón pie
            Set frmPat = New frmManPatronaPie
            frmPat.DeConsulta = True
            frmPat.DatosADevolverBusqueda = "0|1|"
            frmPat.CodigoActual = Text1(31).Text
            frmPat.Show vbModal
            Set frmPat = Nothing
            PonerFoco Text1(31)
    
        Case 13 ' codigo de zona
            Set frmZon = New frmManZonas
            frmZon.DeConsulta = True
            frmZon.DatosADevolverBusqueda = "0|1|"
            frmZon.CodigoActual = Text1(37).Text
            frmZon.Show vbModal
            Set frmZon = Nothing
            PonerFoco Text1(37)
        
        Case 14 ' globalgap
            indice = 38
            
            '[Monica]25/04/2012
            'Set frmGlo = New frmBasico
            'AyudaGlobalGap frmGlo, Text1(indice)
            Set frmGlo = New frmManGlobalGap
            
            frmGlo.DeConsulta = True
            frmGlo.DatosADevolverBusqueda = "0|1|"
            frmGlo.CodigoActual = Text1(38).Text
            frmGlo.Show vbModal
            
            Set frmGlo = Nothing
            PonerFoco Text1(indice)
    
        Case 15 ' codigo de cliente de ariges (suministros)
            indice = 39
            Set frmCli = New frmBasico
            AyudaClienteAriges frmCli, Text1(indice)
            Set frmCli = Nothing
            PonerFoco Text1(indice)
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub


' *********************************************************************************
Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
    If Index = 5 Then
        PonerModoOpcionesMenu Modo
    End If
'Dim I As Byte
'
'    If ModoLineas <> 1 Then
'        Select Case Index
'            Case 0 'telefonos
'                If DataGridAux(Index).Columns.Count > 2 Then
'                    For I = 5 To txtAux.Count - 1
'                        txtAux(I).Text = DataGridAux(Index).Columns(I).Text
'                    Next I
'                    Me.chkAbonos(1).Value = DataGridAux(Index).Columns(17).Text
'
'                End If
'            Case 1 'secciones
'                If DataGridAux(Index).Columns.Count > 2 Then
'                    txtAux2(4).Text = ""
'                    txtAux2(5).Text = ""
'                    txtAux2(0).Text = ""
'                    Set vSeccion = New CSeccion
'                    If vSeccion.LeerDatos(AdoAux(1).Recordset!codsecci) Then
'                        If vSeccion.AbrirConta Then
'                            If DBLet(AdoAux(1).Recordset!codmaccli, "T") <> "" Then
'                                txtAux2(4).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", AdoAux(1).Recordset!codmaccli, "T")
'                            End If
'                            If DBLet(AdoAux(1).Recordset!codmacpro, "T") <> "" Then
'                                txtAux2(5).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", AdoAux(1).Recordset!codmacpro, "T")
'                            End If
'                            txtAux2(0).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", AdoAux(1).Recordset!CodIVA, "N")
'                            vSeccion.CerrarConta
'                        End If
'                    End If
'                    Set vSeccion = Nothing
'                End If
'        End Select
'    End If
End Sub

' ***** si n'hi han varios nivells de tabs *****
Private Sub SituarTab(numTab As Integer)
    On Error Resume Next
    
    If numTab = 0 Then
        SSTab1.Tab = 2
    ElseIf numTab = 1 Then
        SSTab1.Tab = 3
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub
' **********************************************


Private Sub CargaFrame(Index As Integer, enlaza As Boolean)
'Dim tip As Integer
'Dim I As Byte
'
'    AdoAux(Index).ConnectionString = Conn
'    AdoAux(Index).RecordSource = MontaSQLCarga(Index, enlaza)
'    AdoAux(Index).CursorType = adOpenDynamic
'    AdoAux(Index).LockType = adLockPessimistic
'    AdoAux(Index).Refresh
'
'    If Not AdoAux(Index).Recordset.EOF Then
'        PonerCamposForma2 Me, AdoAux(Index), 2, "FrameAux" & Index
'        ' *** si n'hi han tabs sense datagrids, li pose els valors als camps ***
'        If (Index = 3) Then 'datos facturacion
'            tip = AdoAux(Index).Recordset!tipclien
'            If (tip = 1) Then 'persona
'                txtAux2(27).Text = AdoAux(Index).Recordset!ape_raso & "," & AdoAux(Index).Recordset!Nom_Come
'            ElseIf (tip = 2) Then 'empresa
'                txtAux2(27).Text = AdoAux(Index).Recordset!Nom_Come
'            End If
'            txtAux2(28).Text = DBLet(AdoAux(Index).Recordset!desforpa, "T")
'            txtAux2(29).Text = DBLet(AdoAux(Index).Recordset!desrutas, "T")
'            'txtAux2(31).Text = DBLet(AdoAux(Index).Recordset!comision, "T") & " %"
'            txtAux2(32).Text = DBLet(AdoAux(Index).Recordset!nomrapel, "T")
'            'Descripcion cuentas contables de la Contabilidad
'            For I = 35 To 38
'                txtAux2(I).Text = PonerNombreDeCod(txtAux(I), "cuentas", "nommacta", "codmacta", , cConta)
'            Next I
'        End If
'        ' ************************************************************************
'    Else
'        ' *** si n'hi han tabs sense datagrids, li pose els valors als camps ***
'        NetejaFrameAux "FrameAux3" 'neteja només lo que te TAG
'        txtAux2(0).Text = ""
'        txtAux2(1).Text = ""
'
''        txtaux2(27).Text = ""
''        txtaux2(28).Text = ""
''        txtaux2(29).Text = ""
'        'txtAux2(31).Text = ""
''        txtaux2(32).Text = ""
''        For i = 35 To 38
''            txtaux2(i).Text = ""
''        Next i
'        ' **********************************************************************
'    End If
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
Dim Sql2 As String


    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)
    Sql2 = tots
    
    b = DataGridAux(Index).Enabled
    DataGridAux(Index).Enabled = False
    
    AdoAux(Index).ConnectionString = conn
    AdoAux(Index).RecordSource = Sql2
    AdoAux(Index).CursorType = adOpenDynamic
    AdoAux(Index).LockType = adLockPessimistic
    DataGridAux(Index).ScrollBars = dbgNone
    AdoAux(Index).Refresh
    Set DataGridAux(Index).DataSource = AdoAux(Index)
    
    DataGridAux(Index).AllowRowSizing = False
    DataGridAux(Index).RowHeight = 290
    If PrimeraVez Then
        DataGridAux(Index).ClearFields
        DataGridAux(Index).ReBind
        DataGridAux(Index).Refresh
    End If

    For i = 0 To DataGridAux(Index).Columns.Count - 1
        DataGridAux(Index).Columns(i).AllowSizing = False
    Next i
    
    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    
    'DataGridAux(Index).Enabled = b
'    PrimeraVez = False
    
    Select Case Index
        Case 1 'clasificacion segun la calidad
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;S|txtaux1(1)|T|Cód.|800|;S|cmdAux(1)|B|||;" 'codsocio,codsecci
            tots = tots & "S|txtAux2(1)|T|Nombre|3870|;"
            tots = tots & "S|txtaux1(3)|T|Muestra|1200|;"
            arregla tots, DataGridAux(Index), Me
        
            DataGridAux(Index).Columns(4).Alignment = dbgRight
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))

            If (enlaza = True) And (Not AdoAux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
                If VisualizaClasificacion Then
                    PonerClasificacionGrafica
                End If
            Else
                For i = 0 To 3
                    txtAux1(i).Text = ""
                Next i
                txtAux2(1).Text = ""
                Me.MSChart1.visible = False
            End If
        
        Case 0 ' coopropietarios
            tots = "N||||0|;N||||0|;S|txtaux3(2)|T|Cód.|1000|;S|cmdAux(0)|B|||;" 'codsocio,numlinea
            tots = tots & "S|txtAux2(0)|T|Nombre|3870|;"
            tots = tots & "S|txtaux3(3)|T|Porcentaje|1200|;"
            arregla tots, DataGridAux(Index), Me
        
            DataGridAux(Index).Columns(4).Alignment = dbgRight
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))

            If (enlaza = True) And (Not AdoAux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
'                SumaTotalPorcentajes
            Else
                For i = 0 To 3
                    txtAux3(i).Text = ""
                Next i
                txtAux2(0).Text = ""
            End If
         
         
        Case 2 'parcelas del campo
            'si es visible|control|tipo campo|nombre campo|ancho control|
            If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
                tots = "N||||0|;N||||0|;S|txtaux4(2)|T|Poligono|1225|;" 'codsocio,codsecci
                tots = tots & "S|txtaux4(3)|T|Parcela|1225|;"
                tots = tots & "S|txtaux4(4)|T|Subrecinto|1225|;"
                tots = tots & "S|txtaux4(5)|T|Recinto|1225|;"
                tots = tots & "N|txtaux4(6)|T|CodSigpac|1000|;"
                tots = tots & "S|txtaux4(7)|T|Has.Parc.Sig|1600|;"
                tots = tots & "S|txtaux4(8)|T|Has.Recinto|1600|;"
                tots = tots & "S|txtaux4(9)|T|Has.Catastro|1600|;"
                tots = tots & "S|txtaux4(10)|T|Has.Cult.Recinto|1600|;"
            Else
                tots = "N||||0|;N||||0|;S|txtaux4(2)|T|Poligono|1225|;" 'codsocio,codsecci
                tots = tots & "S|txtaux4(3)|T|Parcela|1225|;"
                tots = tots & "S|txtaux4(4)|T|Subparcela|1225|;"
                tots = tots & "S|txtaux4(5)|T|Recinto|1225|;"
                tots = tots & "S|txtaux4(6)|T|CodSigpac|1000|;"
                tots = tots & "S|txtaux4(7)|T|Has.Sigpac|1350|;"
                tots = tots & "S|txtaux4(8)|T|Has.Cult.Sigpac|1350|;"
                tots = tots & "S|txtaux4(9)|T|Has.Catastro|1350|;"
                tots = tots & "S|txtaux4(10)|T|Has.Cult.Catastro|1350|;"
            End If
            arregla tots, DataGridAux(Index), Me
        
            DataGridAux(Index).Columns(4).Alignment = dbgRight
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))

            If (enlaza = True) And (Not AdoAux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
            
            Else
                For i = 0 To 3
                    txtAux4(i).Text = ""
                Next i
            End If
                 
            CalcularTotalSuperficie Sql2
                 
        Case 3 'incidencias de seguro
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;S|txtaux5(2)|T|Fecha|1100|;S|cmdAux(3)|B|||;" 'codcampo,numlinea
            tots = tots & "S|txtaux5(3)|T|Incidencia|1100|;S|cmdAux(2)|B|||;"
            tots = tots & "S|txtAux2(2)|T|Descripcion|3500|;"
            '[Monica]26/01/2016: nueva columna de kilos aportacion, cambio etiqueta de los kilos a indemnizables
            tots = tots & "S|txtaux5(4)|T|Kilos Indemniza.|1500|;"
            tots = tots & "S|txtaux5(7)|T|Kilos Aportacion|1500|;"
            
            tots = tots & "S|txtaux5(5)|T|Importe|1500|;"
            tots = tots & "S|txtaux5(6)|T|Fecha Pago|1100|;S|cmdAux(4)|B|||;"
            tots = tots & "N||||0|;S|chkAux(0)|CB|Sin|360|;"
            
            arregla tots, DataGridAux(Index), Me
        
            DataGridAux(Index).Columns(4).Alignment = dbgLeft
            DataGridAux(Index).Columns(5).Alignment = dbgRight
            DataGridAux(Index).Columns(6).Alignment = dbgRight
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))

            If (enlaza = True) And (Not AdoAux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
            
            Else
                For i = 0 To 3
                    txtAux5(i).Text = ""
                Next i
            End If
                 
                 
        Case 4 'hco del campo
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;" 'codcampo,numlinea
            tots = tots & "S|txtaux6(2)|T|Socio|1100|;S|cmdAux(8)|B|||;"
            tots = tots & "S|txtAux2(4)|T|Nombre|2500|;"
            tots = tots & "S|txtaux6(3)|T|Fecha Alta|1100|;S|cmdAux(6)|B|||;"
            tots = tots & "S|txtaux6(4)|T|Fecha Baja|1100|;S|cmdAux(5)|B|||;"
            tots = tots & "S|txtaux6(5)|T|Incidencia|1100|;S|cmdAux(7)|B|||;"
            tots = tots & "S|txtAux2(3)|T|Descripcion|3500|;"
            
            arregla tots, DataGridAux(Index), Me
        
            DataGridAux(Index).Columns(2).Alignment = dbgLeft
            DataGridAux(Index).Columns(3).Alignment = dbgLeft
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))

            If (enlaza = True) And (Not AdoAux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
            
            Else
                For i = 0 To 5
                    txtaux6(i).Text = ""
                Next i
            End If
                 
        Case 5 'gastos del campo
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;" 'codcampo,numlinea
            tots = tots & "S|txtAux7(2)|T|Código|1100|;S|cmdAux(9)|B|||;"
            tots = tots & "S|txtAux2(5)|T|Concepto|4500|;"
            tots = tots & "S|txtAux7(3)|T|Fecha|1100|;S|cmdAux(10)|B|||;"
            tots = tots & "S|txtAux7(4)|T|Importe|1500|;"
            tots = tots & "N||||0|;S|chkAux(1)|CB|Id|360|;"
            
            arregla tots, DataGridAux(Index), Me
        
            DataGridAux(Index).Columns(2).Alignment = dbgLeft
            DataGridAux(Index).Columns(3).Alignment = dbgLeft
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))

            If (enlaza = True) And (Not AdoAux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
            
            Else
                For i = 0 To 4
                    txtAux7(i).Text = ""
                Next i
            End If
                 
        Case 6  'ordenes de recoleccion
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;" 'codcampo
            tots = tots & "S|txtAux8(1)|T|Nro.Orden|1100|;"
            tots = tots & "S|txtAux8(2)|T|Fecha|1100|;S|cmdAux(11)|B|||;"
            
            arregla tots, DataGridAux(Index), Me
        
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))

            If (enlaza = True) And (Not AdoAux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
            
            Else
                For i = 0 To 0
                    txtAux8(i).Text = ""
                Next i
            End If
                 
        Case 7 ' revisiones
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;" 'codcampo,numlinea
            tots = tots & "S|txtAux9(2)|T|Fecha|1100|;S|cmdAux(12)|B|||;"
            tots = tots & "S|txtAux9(3)|T|Técnico|4500|;"
            tots = tots & "S|txtAux9(4)|T|Observaciones|6100|;"
            
            arregla tots, DataGridAux(Index), Me
        
            DataGridAux(Index).Columns(2).Alignment = dbgLeft
            DataGridAux(Index).Columns(3).Alignment = dbgLeft
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))

            If (enlaza = True) And (Not AdoAux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
            
            Else
                For i = 0 To txtAux9.Count - 1
                    txtAux9(i).Text = ""
                Next i
            End If
        
                 
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
        Case 0: nomframe = "FrameAux0" 'coopropietarios
        Case 1: nomframe = "FrameAux1" 'clasificacion
        Case 2: nomframe = "FrameAux2" 'parcelas
        Case 3: nomframe = "FrameAux3" 'agroseguro
        Case 4: nomframe = "FrameAux4" 'hco
        Case 5: nomframe = "FrameAux5" 'concepto de gastos
        Case 6: nomframe = "FrameAux6" 'ordenes de recoleccion
        Case 7: nomframe = "FrameAux7" 'revision de campos
    End Select
    ' ***************************************************************
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomframe) Then
            ' *** si n'hi ha que fer alguna cosa abas d'insertar
            ' *************************************************
            b = BLOQUEADesdeFormulario2(Me, Data1, 1)
            
            '++monica: en caso de estar insertando seccion y que no existan las
            'cuentas contables hacemos esto para que las inserte en contabilidad.
'            If NumTabMto = 1 Then
'               txtAux2(4).Text = PonerNombreCuenta(txtaux1(4), 3, Text1(0))
'               txtAux2(5).Text = PonerNombreCuenta(txtaux1(5), 3, Text1(0))
'            End If
            
            Select Case NumTabMto
                Case 0, 1, 2, 3, 4, 6, 7 ' *** els index de les llinies en grid (en o sense tab) ***
                    CargaGrid NumTabMto, True
                    If b Then BotonAnyadirLinea NumTabMto
                    
                Case 5 ' Caso de gastos de campo, tenemos que insertar un asiento en el diario
                    Screen.MousePointer = vbHourglass
                    
                    frmListado.OpcionListado = 37
                    frmListado.NumCod = "rcampos_gastos.codcampo = " & DBSet(txtAux7(0).Text, "N") & " and rcampos_gastos.numlinea = " & DBSet(txtAux7(1).Text, "N")
                    frmListado.Show vbModal
                    CargaGrid NumTabMto, True
                    
                    Screen.MousePointer = vbDefault
                    
                    CargaGrid NumTabMto, True
                    If b Then BotonAnyadirLinea NumTabMto
                
            End Select
           
            'SituarTab (NumTabMto)
        End If
    End If
End Sub


Private Sub ModificarLinea()
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Integer
Dim Cad As String
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'coopropietarios
        Case 1: nomframe = "FrameAux1" 'secciones
        Case 2: nomframe = "FrameAux2" 'parcelas
        Case 3: nomframe = "FrameAux3" 'seguros
        Case 4: nomframe = "FrameAux4" 'hco
        Case 5: nomframe = "FrameAux5" 'conceptos de gastos
        Case 7: nomframe = "FrameAux7" 'revisiones de campos
        
    End Select
    ' **************************************************************

    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
            ' *** si cal que fer alguna cosa abas d'insertar ***
            If NumTabMto = 0 Then
            End If
            ' ******************************************************
            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
            ModoLineas = 0

'            If NumTabMto <> 3 Then
                V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
                CargaGrid NumTabMto, True
'            End If

            ' *** si n'hi han tabs ***
            'SituarTab (NumTabMto)

            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
'            If NumTabMto <> 3 Then
                DataGridAux(NumTabMto).SetFocus
                AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
'            End If
            ' ***********************************************************

            LLamaLineas NumTabMto, 0
            
        End If
    End If
        
End Sub


Private Sub Modificar()
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Integer
Dim Cad As String
Dim Sql As String
Dim vCadena As String
Dim Produ As Integer

    On Error GoTo eModificar

    conn.BeginTrans

    
    If CLng(Text1(2).Text) <> CLng(VarieAnt) Then
        ' borramos la clasificacion de este campo
        Sql = "delete from rcampos_clasif "
        Sql = Sql & " where codcampo = " & DBSet(Text1(0).Text, "N")
        
        conn.Execute Sql
    End If
    
    b = True
    If CLng(Text1(2).Text) <> CLng(VarieAnt) Or CLng(Text1(1).Text) <> CLng(SocioAnt) Then
        b = ModificarEntradas(Text1(0).Text, SocioAnt, VarieAnt, Text1(1).Text, Text1(2).Text)
    End If
        
    ' modificamos los datos del campo
    If b Then
        If ModificaDesdeFormulario2(Me, 1) Then
            TerminaBloquear
            
            '[Monica]17/09/2013:en el campo ant en picassent ponemos otra cosa
            Produ = DevuelveValor("select codprodu from variedades where codvarie = " & VarieAnt)
            vCadena = CLng(SocioAnt) & "&" & CLng(Text1(0).Text) & "&" & Produ & "&" & VarieAnt
            
            CargarUnCampo CLng(Text1(0).Text), "U", vCadena
            
            PosicionarData
            CargaGrid 0, True
            CargaGrid 1, True
            CargaGrid 2, True
            CargaGrid 3, True
            CargaGrid 4, True
            CargaGrid 5, True
'[Monica]30/09/2013
'            CargaGrid 6, True
            CargarListaOrdenesRecogida Text1(23).Text
            VisualizaClasificacion
        End If
    
        conn.CommitTrans
        Exit Sub
    End If
    

eModificar:
    conn.RollbackTrans
    MuestraError Err.Number, "Modificando lineas"

End Sub


Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " rcampos_clasif.codcampo=" & Val(Text1(0).Text)
    vWhere = vWhere & " and rcampos_clasif.codvarie = " & Val(Text1(2).Text)
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function

Private Function ObtenerWhereCab2(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " rcampos_parcelas.codcampo=" & Val(Text1(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab2 = vWhere
End Function

Private Function ObtenerWhereCab3(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " rcampos_seguros.codcampo=" & Val(Text1(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab3 = vWhere
End Function


Private Function ObtenerWhereCab4(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " rcampos_hco.codcampo=" & Val(Text1(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab4 = vWhere
End Function


Private Function ObtenerWhereCab5(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " rcampos_gastos.codcampo=" & Val(Text1(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab5 = vWhere
End Function


Private Function ObtenerWhereCab6(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " rcampos_ordrec.codcampo=" & Val(Text1(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab6 = vWhere
End Function

Private Function ObtenerWhereCab7(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " rcampos_revision.codcampo=" & Val(Text1(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab7 = vWhere
End Function


' *** neteja els camps dels tabs de grid que
'estan fora d'este, i els camps de descripció ***
Private Sub LimpiarCamposFrame(Index As Integer)
'Dim I As Integer
'    On Error Resume Next
'
'    Select Case Index
'        Case 0 'telefonos
'            For I = 0 To txtAux.Count - 1
'                txtAux(I).Text = ""
'            Next I
'        Case 1 'secciones
'            For I = 0 To txtaux1.Count - 1
'                txtaux1(I).Text = ""
'            Next I
'    End Select
'
'    If Err.Number <> 0 Then Err.Clear
End Sub
' ***********************************************

Private Sub printNou()
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal

    '[Monica]05/02/2014: Personalizamos el informe de campos
    indRPT = 102 ' personalizacion del informe de campos
    
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
    
    cadNombreRPT = nomDocu

    With frmImprimir2
        .cadTabla2 = "rcampos"
        .Informe2 = cadNombreRPT '"rManCampos.rpt"
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
        .ConSubInforme2 = False
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
Dim Sql As String
Dim Rs As ADODB.Recordset

   ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 0 To Combo1.Count - 1
        Combo1(i).Clear
    Next i
    
    'tipo de parcela
    Combo1(0).AddItem "Rústica"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Urbana"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
    'tipo de recoleccion
    Combo1(1).AddItem "Cooperativa"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Socio"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1

    'tipo de campo
    Combo1(3).AddItem "Normal"
    Combo1(3).ItemData(Combo1(3).NewIndex) = 0
    Combo1(3).AddItem "Comercio"
    Combo1(3).ItemData(Combo1(3).NewIndex) = 1
    Combo1(3).AddItem "Industria"
    Combo1(3).ItemData(Combo1(3).NewIndex) = 2
    
    
    'TIPO DE SISTEMA DE RIEGO
    Sql = "select codriego, nomriego from rriego "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    While Not Rs.EOF
'        Sql = Replace(Rs.Fields(1).Value, "Factura", "Fac.")
        Sql = Rs.Fields(1).Value
'        Sql = Rs.Fields(0).Value & " - " & Sql
        Combo1(2).AddItem Sql 'campo del codigo
        Combo1(2).ItemData(Combo1(2).NewIndex) = i
        i = i + 1
        Rs.MoveNext
    Wend
    
    ' Entrega Ficha de Cultivo
    Sql = "select codtipo, descripcion from rfichculti "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    While Not Rs.EOF
        Sql = Rs.Fields(1).Value
        Combo1(4).AddItem Sql 'campo del codigo
        Combo1(4).ItemData(Combo1(4).NewIndex) = i
        i = i + 1
        Rs.MoveNext
    Wend
    
    
End Sub


Private Sub txtAux1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

    If Not PerderFocoGnral(txtAux1(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 1 ' calidad
            If PonerFormatoEntero(txtAux1(Index)) Then
                txtAux2(Index).Text = PonerNombreDeCod(txtAux1(Index), "rcalidad", "nomcalid", "codcalid", "N", , "codvarie", txtAux1(2).Text, "N")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe la Calidad: " & txtAux1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCalid = New frmManCalidades
                        frmCalid.DatosADevolverBusqueda = "0|1|"
                        frmCalid.NuevoCodigo = txtAux1(Index).Text
                        txtAux1(Index).Text = ""
                        TerminaBloquear
                        frmCalid.Show vbModal
                        Set frmCalid = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtAux1(Index).Text = ""
                    End If
                    PonerFoco txtAux1(Index)
                End If
            Else
                txtAux2(Index).Text = ""
            End If

        Case 3 ' muestra debe sumar el 100%
            If PonerFormatoDecimal(txtAux1(Index), 4) Then
                cmdAceptar.SetFocus
            End If

'        Case 2, 3 'fecha de alta y de baja
'            PonerFormatoFecha txtaux1(Index)

'        Case 4, 5 'cta Cliente y Proveedor
'            If txtaux1(Index).Text = "" Then Exit Sub
'
'            If Not vSeccion Is Nothing Then
'                txtAux2(Index).Text = PonerNombreCuenta(txtaux1(Index), Modo)
'                If txtaux1(Index).Text <> "" Then
'                    If Not vSeccion.CtaConRaizCorrecta(txtaux1(Index).Text, Index - 4) Then
'                        MsgBox "La cuenta no tiene la raiz correcta. Revise.", vbExclamation
'                    Else
'                        ' si la cuenta es correcta y no existe la insertamos en contabilidad
'                        txtAux2(Index).Text = PonerNombreCuenta(txtaux1(Index), 3, Text1(0))
'                    End If
'                End If
'            End If
'
'        Case 6 'codigo iva
'            If txtaux1(Index).Text = "" Then Exit Sub
'
'            If Not vSeccion Is Nothing Then
'                  txtAux2(0).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", txtaux1(Index).Text, "N")
'            End If
'            cmdAceptar.SetFocus

    End Select

    ' ******************************************************************************
End Sub

Private Sub txtAux1_GotFocus(Index As Integer)
   If Not txtAux1(Index).MultiLine Then ConseguirFocoLin txtAux1(Index)
End Sub

Private Sub txtAux1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not txtAux1(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub txtAux1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not txtAux1(Index).MultiLine Then
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
    End If
End Sub


Private Sub PonerDatosPartida()
Dim Zona As String
Dim OtroCampo As String
Dim CodPobla As String

    Zona = ""
'    Text3(3).Text = ""
    Text4(3).Text = ""
    Text5(3).Text = ""
    
    OtroCampo = "codpobla"
    Zona = DevuelveDesdeBDNew(cAgro, "rpartida", "codzonas", "codparti", Text1(3), "N", OtroCampo)
'    Text3(3).Text = Zona
    
    '[Monica]16/05/2013: si no es utxera ni escalona
    '[Monica]20/06/2012: si estamos en modo 2 no debo de mostrar la zona de la partida sino lo que hay grabado
    '[Monica]03/04/2014:  quito la condicion and modo <> 4
    If Modo <> 2 And vParamAplic.Cooperativa <> 8 And vParamAplic.Cooperativa <> 10 Then
        Text1(37).Text = Zona
    End If
    Zona = Text1(37).Text
    
    If Zona <> "" Then
        Text4(3).Text = DevuelveDesdeBDNew(cAgro, "rzonas", "nomzonas", "codzonas", Zona, "N")
        If OtroCampo <> "" Then
            CodPobla = OtroCampo
            If CodPobla <> "" Then Text5(3).Text = DevuelveDesdeBDNew(cAgro, "rpueblos", "despobla", "codpobla", CodPobla, "T")
        End If
    End If

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
        .Show vbModal
    End With
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    CadParam = ""
    numParam = 0
End Sub

Private Sub PonerClasificacionGrafica()
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim i As Integer
Dim arrData()
Dim TotalPorc As Currency
   
    Sql = "select count(*) from rcampos_clasif, rcalidad where rcampos_clasif.codcampo = " & Data1.Recordset!codcampo
    Sql = Sql & " and rcampos_clasif.codvarie = rcalidad.codvarie "
    Sql = Sql & " and rcampos_clasif.codcalid = rcalidad.codcalid "
    
    i = TotalRegistros(Sql)
    
    MSChart1.visible = True
    If i = 0 Then
        MSChart1.visible = False
        Exit Sub
    End If
    
    ReDim arrData(i - 1, 2)
   
    Sql = "select rcampos_clasif.muestra, rcalidad.nomcalid from rcampos_clasif, rcalidad where rcampos_clasif.codcampo = " & Data1.Recordset!codcampo
    Sql = Sql & " and rcampos_clasif.codvarie = rcalidad.codvarie "
    Sql = Sql & " and rcampos_clasif.codcalid = rcalidad.codcalid "
    Sql = Sql & " order by rcampos_clasif.codcalid "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    i = 0
    TotalPorc = 0
    While Not Rs.EOF
        arrData(i, 1) = DBLet(Rs!nomcalid, "T") '& " " & DBLet(Rs!muestra, "N")
        arrData(i, 2) = DBLet(Rs!Muestra, "N")
        
        TotalPorc = TotalPorc + DBLet(Rs!Muestra, "N")
        
        i = i + 1
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    MSChart1.ChartData = arrData
    MSChart1.ColumnLabel = "Porcentaje Total : " & TotalPorc & "%"
    
'    arrData(0, 1) = "Ene"   ' Establece las etiquetas en la primera serie.
'    arrData(1, 1) = "Feb"
'    arrData(2, 1) = "Mar"
'
'    arrData(0, 2) = 8
'    arrData(1, 2) = 4
'    arrData(2, 2) = 0.3
End Sub


Private Sub SumaTotalPorcentajes(numTab As Integer)
Dim Sql As String
Dim i As Currency
Dim Rs As ADODB.Recordset
   
   Select Case numTab
        Case 0 ' coopropietarios
            Sql = "select sum(porcentaje) from rcampos_cooprop where rcampos_cooprop.codcampo = " & Data1.Recordset!codcampo
            
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            i = 0
            If Not Rs.EOF Then
                i = DBLet(Rs.Fields(0).Value, "N")
            End If
            
            If i = 0 Then Exit Sub
            
            If i <> 100 Then
                NumTabMto = 0
                SituarTab numTab
                MsgBox "La suma de porcentajes es " & i & ". Debe de ser 100%. Revise.", vbExclamation
            End If
   
        Case 1 ' clasificaciones
            Sql = "select sum(muestra) from rcampos_clasif where rcampos_clasif.codcampo = " & Data1.Recordset!codcampo
            
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            i = 0
            If Not Rs.EOF Then
                i = DBLet(Rs.Fields(0).Value, "N")
            End If
            
            If i <> 100 Then
                MsgBox "La suma de muestras es " & i & ". Debe de ser 100%. Revise.", vbExclamation
            End If

        
   End Select

End Sub


Private Function VisualizaClasificacion() As Boolean
Dim Sql As String


    If Data1.Recordset.EOF Then
        VisualizaClasificacion = False
        Exit Function
    End If

    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "variedades", "tipoclasifica", "codvarie", Data1.Recordset!codvarie, "N")
    
    SSTab1.TabEnabled(3) = (Sql = "0")
    SSTab1.TabVisible(3) = (Sql = "0")
    
    VisualizaClasificacion = (Sql = "0")

End Function


Private Sub BloquearTodoExceptoDatosTecnicos(b As Boolean)
Dim i As Integer

    FrameAux0.Enabled = Not b
    FrameAux1.Enabled = Not b
    FrameAux2.Enabled = Not b
    Frame3.Enabled = Not b
    Frame6.Enabled = Not b
    Frame7.Enabled = Not b
    Frame8.Enabled = Not b And vParamAplic.BDAriges <> ""
    Frame9.Enabled = Not b
    FrameDatosDtoAdministracion.Enabled = Not b
    Frame4.Enabled = Not b
    For i = 1 To 3
        Text1(i).Enabled = Not b
    Next i
    For i = 21 To 23
        Text1(i).Enabled = Not b
    Next i
    imgZoom(0).Enabled = Not b
    For i = 1 To 4
        imgBuscar(i).Enabled = Not b
    Next i
    
End Sub


Private Sub TxtAux4_GotFocus(Index As Integer)
    If Not txtAux4(Index).MultiLine Then ConseguirFocoLin txtAux4(Index)
End Sub

Private Sub TxtAux4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not txtAux4(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub TxtAux4_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub TxtAux4_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

    If Not PerderFocoGnral(txtAux4(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 2, 3 'poligono y parcela
            PonerFormatoEntero txtAux4(Index)
        
        Case 5 'recinto
            PonerFormatoEntero txtAux4(Index)
        
        Case 6 'COD SIGPAC
            PonerFormatoEntero txtAux4(Index)
            
        Case 7, 8, 9, 10 'superficies en hectareas
            If Modo = 1 Then Exit Sub
            If PonerFormatoDecimal(txtAux4(Index), 7) Then
                If Index = 10 Then cmdAceptar.SetFocus
            Else
                If Index = 10 And txtAux4(Index) = "" Then cmdAceptar.SetFocus
            End If
            

    End Select

    ' ******************************************************************************
End Sub

'*******************************
Private Sub TxtAux5_GotFocus(Index As Integer)
    If Not txtAux5(Index).MultiLine Then ConseguirFocoLin txtAux5(Index)
End Sub

Private Sub TxtAux5_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not txtAux5(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub TxtAux5_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub TxtAux5_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

    If Not PerderFocoGnral(txtAux5(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 2, 6 'fecha y fecha de pago
            PonerFormatoFecha txtAux5(Index)
        
        Case 3 ' codigo de incidencia
            If PonerFormatoEntero(txtAux5(Index)) Then
                txtAux2(2).Text = PonerNombreDeCod(txtAux5(Index), "rincidencia", "nomincid", "codincid", "N")
                If txtAux2(2).Text = "" Then
                    cadMen = "No existe la Incidencia: " & txtAux5(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmInc = New frmManInciden
                        frmInc.DatosADevolverBusqueda = "0|1|"
                        frmInc.NuevoCodigo = txtAux1(Index).Text
                        txtAux5(Index).Text = ""
                        TerminaBloquear
                        frmInc.Show vbModal
                        Set frmInc = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtAux5(Index).Text = ""
                    End If
                    PonerFoco txtAux5(Index)
                End If
            Else
                txtAux2(2).Text = ""
            End If
        
        Case 4 'kilos
            PonerFormatoEntero txtAux5(Index)
        
        '[Monica]26/01/2016: nueva columna de kilos aportacion
        Case 7 ' kilos aportacion
            PonerFormatoEntero txtAux5(Index)
        
        Case 5 ' importe
            If Modo = 1 Then Exit Sub
            PonerFormatoDecimal txtAux5(Index), 1
        
    End Select

    ' ******************************************************************************
End Sub



'*********TXTAUX6
Private Sub TxtAux6_GotFocus(Index As Integer)
    If Not txtaux6(Index).MultiLine Then ConseguirFocoLin txtaux6(Index)
End Sub

Private Sub TxtAux6_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not txtaux6(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub TxtAux6_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub TxtAux6_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

    If Not PerderFocoGnral(txtaux6(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 3, 4 'fecha alta y fecha de baja
            PonerFormatoFecha txtaux6(Index)
        
        
        Case 2 ' codigo de socio
            If PonerFormatoEntero(txtaux6(Index)) Then
                txtAux2(4).Text = PonerNombreDeCod(txtaux6(Index), "rsocios", "nomsocio", "codsocio", "N")
                If txtAux2(4).Text = "" Then
                    cadMen = "No existe el Socio: " & txtaux6(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmSoc2 = New frmManSocios
                        frmSoc2.DatosADevolverBusqueda = "0|1|"
                        txtaux6(Index).Text = ""
                        TerminaBloquear
                        frmSoc2.Show vbModal
                        Set frmSoc2 = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtaux6(Index).Text = ""
                    End If
                    PonerFoco txtaux6(Index)
                End If
            Else
                txtAux2(4).Text = ""
            End If
        
        
        Case 5 ' codigo de incidencia
            If PonerFormatoEntero(txtaux6(Index)) Then
                txtAux2(3).Text = PonerNombreDeCod(txtaux6(Index), "rincidencia", "nomincid", "codincid", "N")
                If txtAux2(3).Text = "" Then
                    cadMen = "No existe la Incidencia: " & txtaux6(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmInc = New frmManInciden
                        frmInc.DatosADevolverBusqueda = "0|1|"
                        frmInc.NuevoCodigo = txtaux6(Index).Text
                        txtaux6(Index).Text = ""
                        TerminaBloquear
                        frmInc.Show vbModal
                        Set frmInc = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtaux6(Index).Text = ""
                    End If
                    PonerFoco txtaux6(Index)
                End If
            Else
                txtAux2(3).Text = ""
            End If
        
        
    End Select

    ' ******************************************************************************
End Sub


'*******************************



'*********TXTAUX7
Private Sub TxtAux7_GotFocus(Index As Integer)
    If Not txtAux7(Index).MultiLine Then ConseguirFocoLin txtAux7(Index)
End Sub

Private Sub TxtAux7_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not txtAux7(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub TxtAux7_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub TxtAux7_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

    If Not PerderFocoGnral(txtAux7(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 3 'fecha alta de gasto
            PonerFormatoFecha txtAux7(Index), True
        
        Case 2 ' codigo de concepto de gasto
            If PonerFormatoEntero(txtAux7(Index)) Then
                txtAux2(5).Text = PonerNombreDeCod(txtAux7(Index), "rconcepgasto", "nomgasto", "codgasto", "N")
                If txtAux2(5).Text = "" Then
                    cadMen = "No existe el Concepto de Gasto: " & txtAux7(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmGto = New frmManConcepGasto
                        frmGto.DatosADevolverBusqueda = "0|1|"
                        txtAux7(Index).Text = ""
                        TerminaBloquear
                        frmGto.Show vbModal
                        Set frmGto = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtAux7(Index).Text = ""
                    End If
                    PonerFoco txtAux7(Index)
                End If
            Else
                txtAux2(5).Text = ""
            End If
        
        Case 4 ' Importe
            If PonerFormatoDecimal(txtAux7(Index), 3) Then cmdAceptar.SetFocus
        
    End Select

    ' ******************************************************************************
End Sub


'*******************************

Private Sub BotonCambioSocio()
Dim Sql As String
Dim campo As String
Dim NroContadores As Long

    If Text1(11).Text <> "" Then
        MsgBox "Este campo tiene fecha de baja, no puede haber cambio de socio. Revise.", vbExclamation
        Exit Sub
    End If

    '[Monica]21/09/2012: si la cooperativa es Escalona y hay contadores de ese campo, cambiamos el socio
    '                    de los contadores que me indiquen
    cadHidrantes = ""
    If vParamAplic.Cooperativa = 10 Then
        NroContadores = DevuelveValor("select count(*) from rpozos where codcampo = " & DBSet(Text1(0).Text, "N"))
        If NroContadores <> 0 Then
            Set frmMens2 = New frmMensajes
            
            frmMens2.cadWHERE2 = "1"
            frmMens2.OpcionMensaje = 39
            frmMens2.cadWHERE = " and codcampo = " & DBSet(Text1(0).Text, "N")
            frmMens2.Show vbModal
            
            Set frmMens2 = Nothing
        End If
    End If

    campo = Text1(0).Text
    
'    If vParamAplic.Cooperativa = 10 Then
'        NroContadores = DevuelveValor("select count(*) from rpozos where codcampo = " & DBSet(Text1(0).Text, "N"))
'        If cadHidrantes <> "" Then
'            frmListado.CadTag = cadHidrantes
'        Else
'            NroContadores = DevuelveValor("select count(*) from rpozos where codcampo = " & DBSet(Text1(0).Text, "N"))
'            If NroContadores <> 0 Then Exit Sub ' me ha dado cancelar cuando he mostrado los contadores
'        End If
'    End If
    
    frmListado.NumCod = Text1(0).Text ' le pasamos el campo del que vamos a cambiar el socio
    frmListado.OpcionListado = 34
    frmListado.Show vbModal

    TerminaBloquear
        
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    Text1(0).Text = campo
    PosicionarData
    PonerCampos
    If Data1.Recordset.RecordCount > 1 Then PonerModo 2
    CargaGrid 0, True
    CargaGrid 1, True
    CargaGrid 2, True
    CargaGrid 3, True
    CargaGrid 4, True
    CargaGrid 5, True
    '[Monica]30/09/2013
    'CargaGrid 6, True
    CargarListaOrdenesRecogida Text1(23).Text
    VisualizaClasificacion
End Sub



'*******************************

Private Sub ChequeoNroOrden()
Dim Sql As String

    Sql = "and mid(right(concat('00000000',codcampo),8),1,6) <> nrocampo"

    cadCampos = ""

    Set frmMens = New frmMensajes
    
    frmMens.OpcionMensaje = 25
    frmMens.cadWHERE = Sql
    frmMens.Show vbModal
    
    Set frmMens = Nothing

    If cadCampos <> "" Then ModificarNroOrden (cadCampos)


End Sub


Private Sub ModificarNroOrden(vSQL As String)
Dim Sql As String
    
    If BloqueaRegistro("rcampos", vSQL) Then
        Sql = "update rcampos set nrocampo = mid(right(concat('00000000',codcampo),8),1,6) where " & vSQL
        conn.Execute Sql
        
        TerminaBloquear
        
        MsgBox "Proceso realizado correctamente.", vbExclamation
        Data1.Refresh
        PonerCampos
    Else
        MsgBox "No se ha podido realizar el proceso. Hay registros bloqueados por otro usuario", vbExclamation
    End If


End Sub



Private Sub EstablecerOrden(PorHectareas As Boolean)
    If PorHectareas Then
        Text2(5).TabIndex = 83
        Text2(6).TabIndex = 84
        Text2(7).TabIndex = 85
        Text2(33).TabIndex = 86
        
        Text1(5).TabIndex = 8
        Text1(6).TabIndex = 9
        Text1(7).TabIndex = 10
        Text1(33).TabIndex = 11
    Else
        Text2(5).TabIndex = 8
        Text2(6).TabIndex = 9
        Text2(7).TabIndex = 10
        Text2(33).TabIndex = 11
        
        Text1(5).TabIndex = 83
        Text1(6).TabIndex = 84
        Text1(7).TabIndex = 85
        Text1(33).TabIndex = 86
    End If
End Sub



'[Monica]14/02/2013: Calculamos la suma de superficies de parcelas (para todos Picassent y resto)
Private Sub CalcularTotalSuperficie(cadena As String)
Dim SigPac   As Currency
Dim CSigpac As Currency
Dim Catas As Currency
Dim CCatas As Currency

Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Sql As String
Dim Sql2 As String
Dim i As Integer

    On Error Resume Next
    
    
    Sql = "select sum(supsigpa) s1 , sum(supcultsigpa) s2, sum(supcatas) s3, sum(supcultcatas) s4 from (" & cadena & ") aaaaa"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SigPac = 0
    CSigpac = 0
    Catas = 0
    CCatas = 0
    For i = 6 To 12
        txtAux2(i).Text = ""
    Next i
    
    If TotalRegistrosConsulta(cadena) = 0 Then Exit Sub
    
    If Not Rs.EOF Then
        If Rs.Fields(0).Value <> 0 Then SigPac = DBLet(Rs.Fields(0).Value, "N") 'Solo es para saber que hay registros que mostrar
        If Rs.Fields(1).Value <> 0 Then CSigpac = DBLet(Rs.Fields(1).Value, "N") 'Solo es para saber que hay registros que mostrar
        If Rs.Fields(2).Value <> 0 Then Catas = DBLet(Rs.Fields(2).Value, "N") 'Solo es para saber que hay registros que mostrar
        If Rs.Fields(3).Value <> 0 Then CCatas = DBLet(Rs.Fields(3).Value, "N") 'Solo es para saber que hay registros que mostrar
        
        If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
            Sql2 = "select sum(supsigpa) s1, sum(supcatas) s3 from ("
            Sql2 = Sql2 & "select distinct rcampos_parcelas.poligono, rcampos_parcelas.parcela, rcampos_parcelas.supsigpa , rcampos_parcelas.supcatas from rcampos_parcelas where codcampo = " & DBSet(Text1(0).Text, "N") & ") aaaa"
            
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            'cojo los datos del primer registro
            If Not Rs2.EOF Then
                If Rs2.Fields(0).Value <> 0 Then SigPac = DBLet(Rs2.Fields(0).Value, "N")
                If Rs2.Fields(1).Value <> 0 Then Catas = DBLet(Rs2.Fields(1).Value, "N")
            End If
        
            Set Rs2 = Nothing
        
        End If
        
    
        txtAux2(6).Text = Format(SigPac, "###0.0000")
        txtAux2(7).Text = Format(CSigpac, "###0.0000")
        txtAux2(8).Text = Format(Catas, "###0.0000")
        txtAux2(9).Text = Format(CCatas, "###0.0000")
        txtAux2(10).Text = txtAux2(7).Text
        txtAux2(11).Text = txtAux2(8).Text
        txtAux2(12).Text = txtAux2(9).Text
        
    End If
    Rs.Close
    Set Rs = Nothing

    
    DoEvents
    
End Sub





'*********TXTAUX8
Private Sub TxtAux8_GotFocus(Index As Integer)
    If Not txtAux8(Index).MultiLine Then ConseguirFocoLin txtAux8(Index)
End Sub

Private Sub TxtAux8_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not txtAux8(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub TxtAux8_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub TxtAux8_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

    If Not PerderFocoGnral(txtAux8(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 2 'fecha de impresion
            If PonerFormatoFecha(txtAux8(Index), True) Then cmdAceptar.SetFocus
        
    End Select

    ' ******************************************************************************
End Sub


'*******************************
'[Monica]30/09/2013: cargamos las ordenes de recogida
Private Sub CargarListaOrdenesRecogida(NroCampo As String)
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim It As ListItem
Dim Sql4 As String
Dim TotalArray As Long

    Sql = "select rordrecogida.nroorden, rordrecogida.fecimpre, if(tijeralimp = 0, '','*') tijlim,  if(tijeradesinf = 0, '','*') tijdes,   if(capazolimp = 0, '','*') caplim,  if(capazodesinf = 0, '','*') capdes, "
    Sql = Sql & " CASE in1.nivel WHEN 0 THEN '<10' WHEN 1 THEN '10 a 20' WHEN 2 THEN '>20' ElSE '' END nind1, "
    Sql = Sql & " CASE in2.nivel WHEN 0 THEN '<10' WHEN 1 THEN '10 a 20' WHEN 2 THEN '>20' ElSE '' END nind2, "
    Sql = Sql & " CASE in3.nivel WHEN 0 THEN '<10' WHEN 1 THEN '10 a 20' WHEN 2 THEN '>20' ElSE '' END nind3, "
    Sql = Sql & " CASE in4.nivel WHEN 0 THEN '<10' WHEN 1 THEN '10 a 20' WHEN 2 THEN '>20' ElSE '' END nind4, "
    Sql = Sql & " CASE in5.nivel WHEN 0 THEN '<10' WHEN 1 THEN '10 a 20' WHEN 2 THEN '>20' ElSE '' END nind5, "
    Sql = Sql & " CASE in6.nivel WHEN 0 THEN '<10' WHEN 1 THEN '10 a 20' WHEN 2 THEN '>20' ElSE '' END nind6, "
    Sql = Sql & " CASE in7.nivel WHEN 0 THEN '<10' WHEN 1 THEN '10 a 20' WHEN 2 THEN '>20' ElSE '' END nind7, "
    Sql = Sql & " CASE in8.nivel WHEN 0 THEN '<10' WHEN 1 THEN '10 a 20' WHEN 2 THEN '>20' ElSE '' END nind8 "
    Sql = Sql & " from ((((((((rordrecogida left join rordrecogida_incid in1 on rordrecogida.nroorden = in1.nroorden and in1.idplaga = 1) "
    Sql = Sql & " left join rordrecogida_incid in2 on rordrecogida.nroorden = in2.nroorden and in2.idplaga = 2) "
    Sql = Sql & " left join rordrecogida_incid in3 on rordrecogida.nroorden = in3.nroorden and in3.idplaga = 3) "
    Sql = Sql & " left join rordrecogida_incid in4 on rordrecogida.nroorden = in4.nroorden and in4.idplaga = 4) "
    Sql = Sql & " left join rordrecogida_incid in5 on rordrecogida.nroorden = in5.nroorden and in5.idplaga = 5) "
    Sql = Sql & " left join rordrecogida_incid in6 on rordrecogida.nroorden = in6.nroorden and in6.idplaga = 6) "
    Sql = Sql & " left join rordrecogida_incid in7 on rordrecogida.nroorden = in7.nroorden and in7.idplaga = 7) "
    Sql = Sql & " left join rordrecogida_incid in8 on rordrecogida.nroorden = in8.nroorden and in8.idplaga = 8) "
    Sql = Sql & " where rordrecogida.nrocampo = " & DBSet(NroCampo, "N")
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    ListView4.ListItems.Clear
    
    ListView4.ColumnHeaders.Clear
    
    ListView4.ColumnHeaders.Add , , "Orden Rec.", 1000.0631
    ListView4.ColumnHeaders.Add , , "Fecha", 1100.2522, 0
    ListView4.ColumnHeaders.Add , , "Tij.L", 660.9371, 0
    ListView4.ColumnHeaders.Add , , "Tij.D", 660.9371, 0
    ListView4.ColumnHeaders.Add , , "Cap.L", 660.9371, 0
    ListView4.ColumnHeaders.Add , , "Cap.D", 660.9371, 0
    
    Sql4 = "select nomplaga from rplagasaux order by idplaga "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql4, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        ListView4.ColumnHeaders.Add , , Rs!nomplaga, 910.9371, 0
    
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
       
    TotalArray = 0
    While Not Rs.EOF
        Set It = ListView4.ListItems.Add

        It.Text = Format(DBLet(Rs!nroorden, "N"), "0000000")
        It.SubItems(1) = DBLet(Rs!fecimpre, "F")
        It.SubItems(2) = Rs!tijlim
        It.SubItems(3) = Rs!tijdes
        It.SubItems(4) = Rs!caplim
        It.SubItems(5) = Rs!capdes
        It.SubItems(6) = Rs!nind1
        It.SubItems(7) = Rs!nind2
        It.SubItems(8) = Rs!nind3
        It.SubItems(9) = Rs!nind4
        It.SubItems(10) = Rs!nind5
        It.SubItems(11) = Rs!nind6
        It.SubItems(12) = Rs!nind7
        It.SubItems(13) = Rs!nind8
        
        It.Checked = False

        Rs.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    Rs.Close

End Sub


'*********TXTAUX9
Private Sub TxtAux9_GotFocus(Index As Integer)
    If Not txtAux9(Index).MultiLine Then ConseguirFocoLin txtAux9(Index)
End Sub

Private Sub TxtAux9_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not txtAux9(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub TxtAux9_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub TxtAux9_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

    If Not PerderFocoGnral(txtAux9(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 2 'fecha de revision
            PonerFormatoFecha txtAux9(Index), True
        Case 4 ' observaciones
            cmdAceptar.SetFocus
    End Select

    ' ******************************************************************************
End Sub
