VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmManSocios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Socios"
   ClientHeight    =   10980
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   16425
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmManSocios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10980
   ScaleWidth      =   16425
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
      Left            =   13110
      TabIndex        =   173
      Top             =   255
      Width           =   1605
   End
   Begin VB.Frame FrameDesplazamiento 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   6345
      TabIndex        =   171
      Top             =   30
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   172
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
   Begin VB.Frame FrameBotonGnral2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   3900
      TabIndex        =   169
      Top             =   30
      Width           =   2370
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   170
         Top             =   180
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Baja de Socios"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imprimir Fases"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Socios de Baja"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Exportacion de Socios"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   240
      TabIndex        =   166
      Top             =   30
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   167
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   0
      Left            =   240
      TabIndex        =   70
      Top             =   780
      Width           =   15870
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
         Left            =   12735
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Código Socio|N|S|||rsocios|nroasociado|000000||"
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
         Tag             =   "Código Socio|N|N|1|999999|rsocios|codsocio|000000|S|"
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
         Left            =   11325
         TabIndex        =   164
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
         TabIndex        =   72
         Top             =   315
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
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
         TabIndex        =   71
         Top             =   315
         Width           =   705
      End
   End
   Begin VB.TextBox text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   25
      Left            =   9630
      TabIndex        =   88
      Top             =   1020
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   67
      Top             =   10200
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
         TabIndex        =   68
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
      Left            =   15135
      TabIndex        =   42
      Top             =   10320
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
      Left            =   13935
      TabIndex        =   41
      Top             =   10320
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4230
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
      Left            =   15135
      TabIndex        =   74
      Top             =   10320
      Visible         =   0   'False
      Width           =   1035
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8490
      Left            =   240
      TabIndex        =   69
      Top             =   1590
      Width           =   15900
      _ExtentX        =   28046
      _ExtentY        =   14975
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      Tab             =   3
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
      TabCaption(0)   =   "Datos básicos"
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
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label16"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label17"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "imgFec(3)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label44"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Image2"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Toolbar2"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Frame5"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Toolbar3"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "lw1"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Text3(0)"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Frame8"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Frame9"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).ControlCount=   12
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
      TabCaption(6)   =   "Movimientos"
      TabPicture(6)   =   "frmManSocios.frx":00B4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "FrameAux4"
      Tab(6).ControlCount=   1
      Begin VB.Frame Frame9 
         Height          =   690
         Left            =   135
         TabIndex        =   197
         Top             =   810
         Width           =   13335
         Begin VB.OptionButton Option1 
            Caption         =   "Alta/Baja Campo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   6
            Left            =   11070
            TabIndex        =   204
            Top             =   270
            Width           =   2085
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Imágenes"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   9540
            TabIndex        =   203
            Top             =   270
            Width           =   1275
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Impresión"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   8085
            TabIndex        =   202
            Top             =   270
            Width           =   1410
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Facturas Cliente"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   5895
            TabIndex        =   201
            Top             =   270
            Width           =   1950
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Detalle Entradas"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   3825
            TabIndex        =   200
            Top             =   270
            Width           =   1905
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Entradas por Huerto"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   1470
            TabIndex        =   199
            Top             =   270
            Width           =   2310
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Facturas"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   225
            TabIndex        =   198
            Top             =   270
            Width           =   1410
         End
      End
      Begin VB.Frame FrameAux4 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7275
         Left            =   -74760
         TabIndex        =   182
         Top             =   510
         Width           =   14360
         Begin VB.TextBox txtAux5 
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
            Index           =   7
            Left            =   10950
            MaxLength       =   30
            TabIndex        =   193
            Tag             =   "Nro Docum|N|S|||rsocios_movim|numerodoc|||"
            Text            =   "Doc"
            Top             =   3480
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtAux5 
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
            Index           =   6
            Left            =   8760
            MaxLength       =   30
            TabIndex        =   192
            Tag             =   "Causa|T|S|||rsocios_movim|causa|||"
            Text            =   "Causa"
            Top             =   3450
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtAux5 
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
            TabIndex        =   190
            Tag             =   "Superficie|N|N|||rsocios_movim|supcoope|###,##0.00||"
            Text            =   "Superficie"
            Top             =   3450
            Visible         =   0   'False
            Width           =   3165
         End
         Begin VB.TextBox txtAux5 
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
            TabIndex        =   189
            Tag             =   "Fecha|F|N|||rsocios_movim|fecmovim|dd/mm/yyyy||"
            Text            =   "fecha"
            Top             =   3450
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtAux5 
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
            TabIndex        =   188
            Tag             =   "Código Socio|N|N|1|999999|rsocios_miembros|codsocio|000000|S|"
            Text            =   "Socio"
            Top             =   3450
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtAux5 
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
            TabIndex        =   187
            Tag             =   "Campo|N|N|||rsocios_movim|codcampo|0000000||"
            Text            =   "campo"
            Top             =   3450
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.TextBox txtAux5 
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
            TabIndex        =   186
            Tag             =   "Importe|N|S|||rsocios_movim|importe|###,###,##0.00||"
            Text            =   "Importe"
            Top             =   3450
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtAux5 
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
            TabIndex        =   185
            Tag             =   "Lineal|N|N|||rsocios_miembros|numlinea|#####0|S|"
            Text            =   "Linea"
            Top             =   3450
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Frame FrameToolAux4 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   60
            TabIndex        =   183
            Top             =   0
            Visible         =   0   'False
            Width           =   1500
            Begin MSComctlLib.Toolbar ToolAux 
               Height          =   330
               Index           =   4
               Left            =   150
               TabIndex        =   184
               Top             =   180
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               AllowCustomize  =   0   'False
               Style           =   1
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
         End
         Begin MSAdodcLib.Adodc AdoAux 
            Height          =   375
            Index           =   4
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
            Bindings        =   "frmManSocios.frx":00D0
            Height          =   5985
            Index           =   4
            Left            =   60
            TabIndex        =   191
            Top             =   720
            Width           =   13175
            _ExtentX        =   23230
            _ExtentY        =   10557
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
      Begin VB.Frame FrameAux3 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7365
         Left            =   -74955
         TabIndex        =   153
         Top             =   405
         Width           =   14360
         Begin VB.Frame FrameToolAux3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   60
            TabIndex        =   180
            Top             =   0
            Width           =   1500
            Begin MSComctlLib.Toolbar ToolAux 
               Height          =   330
               Index           =   3
               Left            =   150
               TabIndex        =   181
               Top             =   180
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               AllowCustomize  =   0   'False
               Style           =   1
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
            Index           =   1
            Left            =   720
            MaxLength       =   6
            TabIndex        =   160
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
            TabIndex        =   159
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
            Tag             =   "Código Socio|N|N|1|999999|rsocios_miembros|codsocio|000000|S|"
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
            TabIndex        =   158
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
            TabIndex        =   156
            Tag             =   "Nombre|T|N|||rsocios_miembros|nommiembro|||"
            Text            =   "Nombre"
            Top             =   3450
            Visible         =   0   'False
            Width           =   3165
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
            Bindings        =   "frmManSocios.frx":00E8
            Height          =   5985
            Index           =   3
            Left            =   60
            TabIndex        =   157
            Top             =   720
            Width           =   13175
            _ExtentX        =   23230
            _ExtentY        =   10557
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
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4395
         Left            =   9945
         TabIndex        =   148
         Top             =   1710
         Visible         =   0   'False
         Width           =   3465
         Begin VB.CommandButton cmdAccCRM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   0
            Picture         =   "frmManSocios.frx":0100
            Style           =   1  'Graphical
            TabIndex        =   151
            ToolTipText     =   "Insertar Imágen"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton cmdAccCRM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1080
            Picture         =   "frmManSocios.frx":0B02
            Style           =   1  'Graphical
            TabIndex        =   150
            ToolTipText     =   "Ver Documento"
            Top             =   0
            Width           =   375
         End
         Begin VB.CommandButton cmdAccCRM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   480
            Picture         =   "frmManSocios.frx":108C
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
            Top             =   405
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
         Height          =   2205
         Left            =   -74850
         TabIndex        =   142
         Top             =   1095
         Width           =   7245
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
            Left            =   1545
            MaxLength       =   35
            TabIndex        =   5
            Tag             =   "Domicilio|T|N|||rsocios|dirsocio|||"
            Top             =   465
            Width           =   5310
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
            Left            =   1545
            MaxLength       =   6
            TabIndex        =   6
            Tag             =   "C.Postal|T|N|||rsocios|codpostal|||"
            Top             =   975
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
            Left            =   2370
            MaxLength       =   30
            TabIndex        =   7
            Tag             =   "Población|T|N|||rsocios|pobsocio|||"
            Top             =   975
            Width           =   4470
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
            Left            =   1545
            MaxLength       =   30
            TabIndex        =   8
            Tag             =   "Provincia|T|N|||rsocios|prosocio|||"
            Top             =   1470
            Width           =   5295
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   1260
            ToolTipText     =   "Buscar Población"
            Top             =   1035
            Width           =   240
         End
         Begin VB.Label Label6 
            Caption         =   "Dirección"
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
            Top             =   495
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
            Top             =   1530
            Width           =   945
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
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
            Top             =   1005
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
         Height          =   2205
         Left            =   -74850
         TabIndex        =   138
         Top             =   3360
         Width           =   7245
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
            Left            =   1545
            MaxLength       =   35
            TabIndex        =   9
            Tag             =   "Domicilio|T|N|||rsocios|dirsociocorreo|||"
            Top             =   495
            Width           =   5280
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
            Left            =   1545
            MaxLength       =   6
            TabIndex        =   10
            Tag             =   "C.Postal|T|N|||rsocios|codpostalcorreo|||"
            Top             =   975
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
            Left            =   2310
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "Población|T|N|||rsocios|pobsociocorreo|||"
            Top             =   975
            Width           =   4500
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
            Left            =   1545
            MaxLength       =   30
            TabIndex        =   12
            Tag             =   "Provincia|T|N|||rsocios|prosociocorreo|||"
            Top             =   1470
            Width           =   5265
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   1260
            ToolTipText     =   "Buscar Población"
            Top             =   990
            Width           =   240
         End
         Begin VB.Label Label6 
            Caption         =   "Dirección"
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
            Top             =   495
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
            Top             =   1500
            Width           =   945
         End
         Begin VB.Label Label1 
            Caption         =   "Población"
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
            Top             =   1005
            Width           =   975
         End
      End
      Begin VB.Frame FrameAux2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7380
         Left            =   -74955
         TabIndex        =   132
         Top             =   405
         Width           =   14360
         Begin VB.Frame FrameToolAux2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   60
            TabIndex        =   176
            Top             =   0
            Width           =   1500
            Begin MSComctlLib.Toolbar ToolAux 
               Height          =   330
               Index           =   2
               Left            =   150
               TabIndex        =   177
               Top             =   180
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               AllowCustomize  =   0   'False
               Style           =   1
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
            Index           =   2
            Left            =   1290
            MaxLength       =   9
            TabIndex        =   134
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
            TabIndex        =   135
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
            TabIndex        =   136
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
            TabIndex        =   133
            Tag             =   "Numero Fases|N|N|||rsocios_pozos|numfases|000|S|"
            Text            =   "Fases"
            Top             =   3405
            Visible         =   0   'False
            Width           =   465
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
            Bindings        =   "frmManSocios.frx":1A8E
            Height          =   5985
            Index           =   2
            Left            =   60
            TabIndex        =   137
            Top             =   720
            Width           =   9995
            _ExtentX        =   17621
            _ExtentY        =   10557
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
         Left            =   11295
         TabIndex        =   125
         Text            =   "Text4"
         Top             =   2040
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
         Left            =   -69195
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Fecha Nacimiento|F|S|||rsocios|fechanac|dd/mm/yyyy||"
         Top             =   660
         Width           =   1260
      End
      Begin VB.Frame FrameAux1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7245
         Left            =   -74955
         TabIndex        =   101
         Top             =   405
         Width           =   15165
         Begin VB.Frame FrameToolAux1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   30
            TabIndex        =   174
            Top             =   0
            Width           =   1500
            Begin MSComctlLib.Toolbar ToolAux 
               Height          =   330
               Index           =   1
               Left            =   150
               TabIndex        =   175
               Top             =   180
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               AllowCustomize  =   0   'False
               Style           =   1
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
         End
         Begin VB.Frame Frame3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1230
            Left            =   30
            TabIndex        =   113
            Top             =   5895
            Width           =   14865
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
               Left            =   11460
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   116
               Text            =   "nomiva"
               Top             =   510
               Width           =   2970
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
               Left            =   6645
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   115
               Text            =   "nomCuenta Proveedor"
               Top             =   510
               Width           =   4250
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
               TabIndex        =   114
               Text            =   "nomCuenta Cliente"
               Top             =   510
               Width           =   4260
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
               TabIndex        =   119
               Top             =   540
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
               Left            =   11070
               TabIndex        =   118
               Top             =   540
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
               Left            =   5625
               TabIndex        =   117
               Top             =   540
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
            TabIndex        =   112
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
            TabIndex        =   48
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
            TabIndex        =   111
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
            TabIndex        =   110
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
            Index           =   0
            Left            =   5760
            TabIndex        =   107
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
            TabIndex        =   47
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
            TabIndex        =   105
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
            TabIndex        =   46
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
            TabIndex        =   104
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
            TabIndex        =   45
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
            TabIndex        =   44
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
            TabIndex        =   43
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
            TabIndex        =   103
            Tag             =   "Nombre|N|N|||rsocios_seccion|codsocio|000000|S|"
            Text            =   "socio"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
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
            Bindings        =   "frmManSocios.frx":1AA6
            Height          =   4545
            Index           =   1
            Left            =   30
            TabIndex        =   102
            Top             =   750
            Width           =   14360
            _ExtentX        =   25321
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
      Begin VB.Frame Frame4 
         Caption         =   "Teléfonos"
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
         Height          =   1710
         Left            =   -74850
         TabIndex        =   96
         Top             =   5610
         Width           =   7245
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
            Left            =   5310
            MaxLength       =   10
            TabIndex        =   16
            Tag             =   "Móvil|T|S|||rsocios|movsocio|||"
            Top             =   990
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
            Left            =   1530
            MaxLength       =   10
            TabIndex        =   14
            Tag             =   "Telfno 2|T|S|||rsocios|telsoci2|||"
            Text            =   "123456789012345"
            Top             =   1005
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
            Left            =   5310
            MaxLength       =   10
            TabIndex        =   15
            Tag             =   "Teléfono 3|T|S|||rsocios|telsoci3|||"
            Text            =   "0000000000"
            Top             =   420
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
            Left            =   1530
            MaxLength       =   10
            TabIndex        =   13
            Tag             =   "Teléfono 1|T|S|||rsocios|telsoci1|||"
            Text            =   "123456789012345"
            Top             =   405
            Width           =   1545
         End
         Begin VB.Label Label11 
            Caption         =   "Número 3"
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
            Left            =   4080
            TabIndex        =   100
            Top             =   450
            Width           =   1050
         End
         Begin VB.Label Label20 
            Caption         =   "Número 2"
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
            TabIndex        =   99
            Top             =   1050
            Width           =   990
         End
         Begin VB.Label Label25 
            Caption         =   "Móvil"
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
            Left            =   4065
            TabIndex        =   98
            Top             =   1050
            Width           =   495
         End
         Begin VB.Label Label41 
            Caption         =   "Número 1"
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
            TabIndex        =   97
            Top             =   450
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
         Left            =   -73605
         MaxLength       =   50
         TabIndex        =   17
         Tag             =   "E-mail|T|S|||rsocios|maisocio|||"
         Top             =   7770
         Width           =   5580
      End
      Begin VB.Frame FrameAux0 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5490
         Left            =   -74955
         TabIndex        =   81
         Top             =   405
         Width           =   12360
         Begin VB.Frame FrameToolAux0 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   60
            TabIndex        =   178
            Top             =   0
            Width           =   1500
            Begin MSComctlLib.Toolbar ToolAux 
               Height          =   330
               Index           =   0
               Left            =   150
               TabIndex        =   179
               Top             =   180
               Width           =   1290
               _ExtentX        =   2275
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               AllowCustomize  =   0   'False
               Style           =   1
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
            Height          =   1065
            Index           =   16
            Left            =   7425
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   66
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
            TabIndex        =   55
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
            TabIndex        =   59
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
            TabIndex        =   50
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
            TabIndex        =   49
            Tag             =   "Código Socio|N|N|1|999999|rsocios_telefonos|codsocio|000000|S|"
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
            TabIndex        =   52
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
            TabIndex        =   53
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
            TabIndex        =   56
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
            TabIndex        =   57
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
            TabIndex        =   63
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
            TabIndex        =   64
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
            TabIndex        =   65
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
            TabIndex        =   51
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
            TabIndex        =   62
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
            TabIndex        =   54
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
            TabIndex        =   61
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
            TabIndex        =   60
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
            TabIndex        =   58
            Tag             =   "Teléfono|T|S|||rsocios_telefonos|telefono1|||"
            Text            =   "1234567890"
            Top             =   2235
            Width           =   1410
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
            Bindings        =   "frmManSocios.frx":1ABE
            Height          =   4545
            Index           =   0
            Left            =   60
            TabIndex        =   82
            Top             =   735
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
            Caption         =   "Dirección"
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
            TabIndex        =   109
            Top             =   870
            Width           =   1140
         End
         Begin VB.Image imgZoom 
            Height          =   240
            Index           =   1
            Left            =   9000
            Tag             =   "-1"
            ToolTipText     =   "Zoom descripción"
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
            TabIndex        =   92
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
            TabIndex        =   91
            Top             =   1815
            Width           =   1230
         End
         Begin VB.Label Label2 
            Caption         =   "Población"
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
            TabIndex        =   90
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
            TabIndex        =   87
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
            TabIndex        =   86
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
            TabIndex        =   85
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
            TabIndex        =   84
            Top             =   2730
            Width           =   555
         End
         Begin VB.Label Label3 
            Caption         =   "Teléfono"
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
            Top             =   2280
            Width           =   975
         End
      End
      Begin VB.Frame FrameDatosDtoAdministracion 
         Caption         =   "Datos Relacionados Dto.Administración"
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
         Height          =   6900
         Left            =   -67440
         TabIndex        =   76
         Top             =   450
         Width           =   8100
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
            Index           =   32
            Left            =   1770
            MaxLength       =   25
            TabIndex        =   38
            Tag             =   "Nro.REGEPA|T|S|||rsocios|nroregepa|||"
            Top             =   6210
            Width           =   3375
         End
         Begin VB.CheckBox chkAbonos 
            Caption         =   "Hay embargo"
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
            Left            =   6075
            TabIndex        =   39
            Tag             =   "Hay embargo|N|N|||rsocios|hayembargo||N|"
            Top             =   6210
            Width           =   1620
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
            Left            =   5355
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Tag             =   "Tipo Contrato|N|N|0|3|rsocios|tipcontrato||N|"
            Top             =   5625
            Width           =   2340
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
            Index           =   31
            Left            =   1770
            MaxLength       =   10
            TabIndex        =   36
            Tag             =   "Fecha Renovacion|F|S|||rsocios|fecrenovacion|dd/mm/yyyy||"
            Top             =   5625
            Width           =   1260
         End
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
            Left            =   3750
            TabIndex        =   165
            Tag             =   "No Devolver Aportaciones|N|N|||rsocios|nodevolverapor||N|"
            Top             =   1005
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
            Top             =   3960
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
            Top             =   3960
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
            Top             =   4485
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
            Top             =   3450
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
            Top             =   5070
            Width           =   5280
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
            Tag             =   "Código Globalgap|T|S|||rsocios|codigoggap|||"
            Top             =   5070
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
            Left            =   2820
            TabIndex        =   26
            Tag             =   "Fact.Interna ADV|N|N|0|1|rsocios|esfactadvinterna||N|"
            Top             =   3450
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
            Left            =   5355
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Tag             =   "Tipo Relacion|N|N|0|2|rsocios|tiporelacion||N|"
            Top             =   2880
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
            TabIndex        =   120
            Top             =   1560
            Width           =   5445
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
            Tag             =   "Código Cooperativa|N|N|0|99|rsocios|codcoope|00||"
            Top             =   1560
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
            Left            =   6630
            TabIndex        =   27
            Tag             =   "Correo|N|N|||rsocios|correo||N|"
            Top             =   3450
            Width           =   1260
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
            Left            =   2820
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Tag             =   "Tipo Produccion|N|N|0|3|rsocios|tipoprod||N|"
            Top             =   2880
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
            Left            =   210
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Tag             =   "Tipo IRPF|N|N|0|2|rsocios|tipoirpf||N|"
            Top             =   2880
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
            Top             =   4485
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
            Top             =   4485
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
            Top             =   4485
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
            Top             =   4485
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
            Tag             =   "Código Situacion|N|N|0|99|rsocios|codsitua|00||"
            Top             =   2085
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
            TabIndex        =   77
            Top             =   2085
            Width           =   5445
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
            Top             =   1035
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
            Top             =   510
            Width           =   1260
         End
         Begin VB.Label Label42 
            Caption         =   "Nro.REGEPA"
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
            TabIndex        =   196
            Top             =   6240
            Width           =   1305
         End
         Begin VB.Image imgAyuda 
            Height          =   240
            Index           =   1
            Left            =   5715
            MousePointer    =   4  'Icon
            Tag             =   "-1"
            ToolTipText     =   "Ayuda"
            Top             =   6255
            Width           =   240
         End
         Begin VB.Label Label40 
            Caption         =   "Tipo de Contrato"
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
            Left            =   3420
            TabIndex        =   195
            Top             =   5670
            Width           =   1830
         End
         Begin VB.Label Label37 
            Caption         =   "F.Renovación"
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
            TabIndex        =   194
            Top             =   5655
            Width           =   1305
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   4
            Left            =   1515
            Picture         =   "frmManSocios.frx":1AD6
            ToolTipText     =   "Buscar fecha"
            Top             =   5670
            Width           =   240
         End
         Begin VB.Image imgDoc 
            Height          =   435
            Index           =   0
            Left            =   7200
            ToolTipText     =   "Aportaciones del Socio"
            Top             =   975
            Width           =   495
         End
         Begin VB.Image imgDoc 
            Height          =   465
            Index           =   1
            Left            =   7200
            ToolTipText     =   "Impresión Documento Alta/Baja"
            Top             =   360
            Width           =   510
         End
         Begin VB.Label Label14 
            Caption         =   "Documentos Alta/Baja/Transmisión"
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
            Left            =   3720
            TabIndex        =   163
            Top             =   510
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
            TabIndex        =   162
            Top             =   3990
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
            TabIndex        =   161
            Top             =   4020
            Width           =   705
         End
         Begin VB.Image imgAyuda 
            Height          =   240
            Index           =   0
            Left            =   180
            MousePointer    =   4  'Icon
            Tag             =   "-1"
            ToolTipText     =   "Ayuda"
            Top             =   3510
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   2
            Left            =   1485
            ToolTipText     =   "Buscar globalgap"
            Top             =   5100
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
            Top             =   5100
            Width           =   1005
         End
         Begin VB.Label Label21 
            Caption         =   "Relación Cooperativa"
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
            Left            =   5355
            TabIndex        =   130
            Top             =   2595
            Width           =   2175
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
            TabIndex        =   121
            Top             =   1620
            Width           =   1185
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   1455
            ToolTipText     =   "Buscar Cooperativa"
            Top             =   1590
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   1
            Left            =   1455
            Picture         =   "frmManSocios.frx":1B61
            ToolTipText     =   "Buscar fecha"
            Top             =   1035
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   0
            Left            =   1455
            Picture         =   "frmManSocios.frx":1BEC
            ToolTipText     =   "Buscar fecha"
            Top             =   510
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
            Left            =   2790
            TabIndex        =   94
            Top             =   2595
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
            TabIndex        =   93
            Top             =   2595
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
            TabIndex        =   89
            Top             =   4530
            Width           =   1005
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   0
            Left            =   1455
            ToolTipText     =   "Buscar Situación"
            Top             =   2115
            Width           =   240
         End
         Begin VB.Label Label23 
            Caption         =   "Situación"
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
            TabIndex        =   80
            Top             =   2145
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
            TabIndex        =   79
            Top             =   1065
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
            TabIndex        =   78
            Top             =   540
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
         Height          =   540
         Index           =   20
         Left            =   -67275
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Tag             =   "Observaciones|T|S|||rsocios|observaciones|||"
         Top             =   7635
         Width           =   7815
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
         Left            =   -73275
         MaxLength       =   15
         TabIndex        =   3
         Tag             =   "NIF / CIF|T|N|||rsocios|nifsocio|||"
         Top             =   660
         Width           =   1920
      End
      Begin MSComctlLib.ListView lw1 
         Height          =   6690
         Left            =   135
         TabIndex        =   123
         Top             =   1665
         Width           =   9330
         _ExtentX        =   16457
         _ExtentY        =   11800
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
         Left            =   225
         TabIndex        =   129
         Top             =   1755
         Visible         =   0   'False
         Width           =   480
         _ExtentX        =   847
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
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   10395
         TabIndex        =   131
         Top             =   2565
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
            TabIndex        =   127
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
            TabIndex        =   126
            Top             =   270
            Width           =   2220
         End
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   4350
         Left            =   180
         TabIndex        =   122
         Top             =   1710
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   7673
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Appearance      =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   13
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
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Altas/Bajas de Campos"
               Object.Tag             =   "6"
            EndProperty
         EndProperty
      End
      Begin VB.Image Image2 
         Height          =   375
         Left            =   180
         Top             =   405
         Width           =   420
      End
      Begin VB.Label Label44 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   300
         Left            =   10035
         TabIndex        =   205
         Top             =   1665
         Width           =   2865
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   10815
         Picture         =   "frmManSocios.frx":1C77
         ToolTipText     =   "Buscar fecha"
         Top             =   2085
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
         Left            =   10095
         TabIndex        =   128
         Top             =   2085
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
         Left            =   720
         TabIndex        =   124
         Top             =   450
         Width           =   7590
      End
      Begin VB.Label Label10 
         Caption         =   "F.Nacimiento"
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
         Left            =   -70785
         TabIndex        =   106
         Top             =   660
         Width           =   1275
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   -69495
         Picture         =   "frmManSocios.frx":1D02
         ToolTipText     =   "Buscar fecha"
         Top             =   660
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
         Left            =   -74625
         TabIndex        =   95
         Top             =   7800
         Width           =   705
      End
      Begin VB.Image imgMail 
         Height          =   240
         Index           =   0
         Left            =   -73890
         Top             =   7800
         Width           =   240
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   -65655
         Tag             =   "-1"
         ToolTipText     =   "Zoom descripción"
         Top             =   7395
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
         Left            =   -67275
         TabIndex        =   75
         Top             =   7365
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
         TabIndex        =   73
         Top             =   660
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
      Left            =   15150
      TabIndex        =   168
      Top             =   195
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
      Top             =   10260
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

Private Const IdPrograma = 2004

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
Private WithEvents frmSocPrev As frmBasico2 ' Socios vista previa
Attribute frmSocPrev.VB_VarHelpID = -1
Private WithEvents frmPue As frmManPueblos
Attribute frmPue.VB_VarHelpID = -1
Private WithEvents frmPue2 As frmManPueblos
Attribute frmPue2.VB_VarHelpID = -1

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
Dim cadB As String

Dim vSeccion As CSeccion
Dim b As Boolean

Private BuscaChekc As String
Private Const CarpetaIMG = "ImgFicFT"

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
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


Dim Tel1Ant As String
Dim Tel2Ant As String


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
Dim cadena1 As String

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
                    
                    
                    
                    
                    If vParamAplic.HayAppAriagro Then PushSocio Text1(0).Text, "I"
                    
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
                    
                    '[Monica]13/08/2018: para el caso de modificacion de un socio no lo habiamos lanzado
                    If vParamAplic.HayAppAriagro Then PushSocio Text1(0).Text, "M"
                    
                    
                    
                    '------------------------------------------------------------------------------
                    '  LOG de acciones
                    Set LOG = New cLOG
                    LOG.InsertarCambiosRegistros 14, vUsu, "Cambio Socios-U", "rsocios", ObtenerWhereCab(False)
                    Set LOG = Nothing
                    '-----------------------------------------------------------------------------
                    
                    '[Monica]11/06/2018: si es un socio < 10000 se comunica a otra cooperativa
                    If Text1(0).Text < cMaxSocio And TieneCamposVariedadComercializada(Text1(0).Text) Then
                        cadena1 = GeneraCadenaUpdate
                        ComunicaCooperativa "rsocios", cadena1, "U", ""
                    End If
                    
                    
                    '[Monica]10/07/2013: Si han cambiado nombre o CCC pregunto si quieren cambiar los datos de la cuenta en la seccion de horto
                    ModificarDatosCuentaContable
                    
                    PosicionarData
                    PonerCampos
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


Private Function GeneraCadenaUpdate() As String
Dim Sql As String

    Sql = "UPDATE rsocios SET "
    Sql = Sql & "nomsocio=" & DBSet(Text1(1), "T")
    Sql = Sql & ",dirsocio=" & DBSet(Text1(3), "T")
    Sql = Sql & ",codpostal=" & DBSet(Text1(4), "T")
    Sql = Sql & ",pobsocio=" & DBSet(Text1(5), "T")
    Sql = Sql & ",prosocio=" & DBSet(Text1(6), "T")
    Sql = Sql & ",nifsocio= " & DBSet(Text1(2), "T")
    Sql = Sql & " WHERE codsocio=" & DBSet(ComprobarCero(Text1(0).Text) + cMaxSocio, "N")

    GeneraCadenaUpdate = Sql

End Function



Private Function PushSocio(Codigo As String, Operacion As String) As Boolean
Dim Sql As String


    On Error GoTo ePushSocio

    PushSocio = False

'    If Not AbrirConexionApp Then Exit Function
    
    If Not ExisteAripush Then Exit Function

    Select Case Operacion
        Case "I" ' insertar
            Sql = DevuelveDesdeBDNew(cAgro, "aripush.usuariospush", "ariagroid", "ariagroid", Text1(0).Text, "N")
            If Sql = "" Then
                Sql = "insert into aripush.usuariospush (nif,nombre,comunId,ariagroId,tiendaId,gasolineraId,telefoniaId,tratamientosId,login,"
                Sql = Sql & "password,email,playerId,direccion,codPostal,poblacion,provincia,telefono1,telefono2,iban)"
                Sql = Sql & " values ("
                Sql = Sql & DBSet(Text1(2), "T") & "," ' nif
                Sql = Sql & DBSet(Text1(1), "T") & "," 'nombre
                Sql = Sql & DBSet(Text1(0), "N") & "," 'codsocio
                Sql = Sql & DBSet(Text1(0), "N") & "," 'ariagroid
                Sql = Sql & DBSet(Text1(0), "N") & "," 'tiendaid
                Sql = Sql & DBSet(Text1(0), "N") & "," 'gasolineraid
                Sql = Sql & DBSet(Text1(0), "N") & "," 'telefoniaid
                Sql = Sql & DBSet(Text1(0), "N") & "," 'tratamientosid
                Sql = Sql & DBSet(Text1(0), "N") & "," 'login
                Sql = Sql & DBSet(Text1(2), "T") & "," 'password el nif
                Sql = Sql & DBSet(Text1(12), "T") & "," 'email
                Sql = Sql & ValorNulo & "," 'playerid
                Sql = Sql & DBSet(Text1(3), "T") & "," 'dirsocio (direccion fiscal)
                Sql = Sql & DBSet(Text1(4), "T") & "," 'codigo postal
                Sql = Sql & DBSet(Text1(5), "T") & "," 'poblacion
                Sql = Sql & DBSet(Text1(6), "T") & "," 'provincia
                Sql = Sql & DBSet(Text1(8), "T") & "," 'telefono1
                Sql = Sql & DBSet(Text1(9), "T")  'telefono2
                '[Monica]10/10/2018: solo en el caso de que no sea picassent se introduce el iban
                If vParamAplic.Cooperativa <> 2 Then
                    Sql = Sql & "," & DBSet(Text1(27).Text & Text1(16).Text & Text1(17).Text & Text1(18).Text & Text1(19).Text, "T")
                Else
                    Sql = Sql & "," & ValorNulo
                End If
                Sql = Sql & ")"
            Else
                Sql = "update aripush.usuariospush set "
                Sql = Sql & " nif = " & DBSet(Text1(2), "T")
                Sql = Sql & ",nombre = " & DBSet(Text1(1).Text, "T")
                Sql = Sql & ",email = " & DBSet(Text1(12).Text, "T")
                Sql = Sql & ",direccion = " & DBSet(Text1(3), "T")  'dirsocio (direccion fiscal)
                Sql = Sql & ",codpostal = " & DBSet(Text1(4), "T")  'codigo postal
                Sql = Sql & ",poblacion = " & DBSet(Text1(5), "T")  'poblacion
                Sql = Sql & ",provincia = " & DBSet(Text1(6), "T")  'provincia
                Sql = Sql & ",telefono1 = " & DBSet(Text1(8), "T")  'telefono1
                Sql = Sql & ",telefono2 = " & DBSet(Text1(9), "T")  'telefono2
                '[Monica]10/10/2018: solo en el caso de no ser picassent se actualiza el iban
                If vParamAplic.Cooperativa <> 2 Then
                    Sql = Sql & ",iban = " & DBSet(Text1(27).Text & Text1(16).Text & Text1(17).Text & Text1(18).Text & Text1(19).Text, "T")
                End If
                Sql = Sql & " where ariagroId = " & DBSet(Text1(0).Text, "N")
            End If

            conn.Execute Sql
            
        
        Case "M" ' modificar

            '[Monica]18/01/2019: si no existe que lo de de alta
            Sql = DevuelveDesdeBDNew(cAgro, "aripush.usuariospush", "ariagroid", "ariagroid", Text1(0).Text, "N")
            If Sql = "" Then
                Sql = "insert into aripush.usuariospush (nif,nombre,comunId,ariagroId,tiendaId,gasolineraId,telefoniaId,tratamientosId,login,"
                Sql = Sql & "password,email,playerId,direccion,codPostal,poblacion,provincia,telefono1,telefono2,iban)"
                Sql = Sql & " values ("
                Sql = Sql & DBSet(Text1(2), "T") & "," ' nif
                Sql = Sql & DBSet(Text1(1), "T") & "," 'nombre
                Sql = Sql & DBSet(Text1(0), "N") & "," 'codsocio
                Sql = Sql & DBSet(Text1(0), "N") & "," 'ariagroid
                Sql = Sql & DBSet(Text1(0), "N") & "," 'tiendaid
                Sql = Sql & DBSet(Text1(0), "N") & "," 'gasolineraid
                Sql = Sql & DBSet(Text1(0), "N") & "," 'telefoniaid
                Sql = Sql & DBSet(Text1(0), "N") & "," 'tratamientosid
                Sql = Sql & DBSet(Text1(0), "N") & "," 'login
                Sql = Sql & DBSet(Text1(2), "T") & "," 'password el nif
                Sql = Sql & DBSet(Text1(12), "T") & "," 'email
                Sql = Sql & ValorNulo & "," 'playerid
                Sql = Sql & DBSet(Text1(3), "T") & "," 'dirsocio (direccion fiscal)
                Sql = Sql & DBSet(Text1(4), "T") & "," 'codigo postal
                Sql = Sql & DBSet(Text1(5), "T") & "," 'poblacion
                Sql = Sql & DBSet(Text1(6), "T") & "," 'provincia
                Sql = Sql & DBSet(Text1(8), "T") & "," 'telefono1
                Sql = Sql & DBSet(Text1(9), "T")  'telefono2
                '[Monica]10/10/2018: solo en el caso de que no sea picassent se introduce el iban
                If vParamAplic.Cooperativa <> 2 Then
                    Sql = Sql & "," & DBSet(Text1(27).Text & Text1(16).Text & Text1(17).Text & Text1(18).Text & Text1(19).Text, "T")
                Else
                    Sql = Sql & "," & ValorNulo
                End If
                Sql = Sql & ")"
            Else
                Sql = "update aripush.usuariospush set "
                Sql = Sql & " nif = " & DBSet(Text1(2), "T")
                Sql = Sql & ",nombre = " & DBSet(Text1(1).Text, "T")
                Sql = Sql & ",email = " & DBSet(Text1(12).Text, "T")
                Sql = Sql & ",direccion = " & DBSet(Text1(3), "T")  'dirsocio (direccion fiscal)
                Sql = Sql & ",codpostal = " & DBSet(Text1(4), "T")  'codigo postal
                Sql = Sql & ",poblacion = " & DBSet(Text1(5), "T")  'poblacion
                Sql = Sql & ",provincia = " & DBSet(Text1(6), "T")  'provincia
                Sql = Sql & ",telefono1 = " & DBSet(Text1(8), "T")  'telefono1
                Sql = Sql & ",telefono2 = " & DBSet(Text1(9), "T")  'telefono2
                '[Monica]10/10/2018: solo en el caso de no ser picassent se actualiza el iban
                If vParamAplic.Cooperativa <> 2 Then
                    Sql = Sql & ",iban = " & DBSet(Text1(27).Text & Text1(16).Text & Text1(17).Text & Text1(18).Text & Text1(19).Text, "T")
                End If
                Sql = Sql & " where ariagroId = " & DBSet(Text1(0).Text, "N")
            End If

            conn.Execute Sql
            
'            Sql = "update aripush.usuariospush set "
'            Sql = Sql & " nif = " & DBSet(Text1(2), "T")
'            Sql = Sql & ",nombre = " & DBSet(Text1(1).Text, "T")
'            Sql = Sql & ",email = " & DBSet(Text1(12).Text, "T")
'            Sql = Sql & ",direccion = " & DBSet(Text1(3), "T") 'dirsocio (direccion fiscal)
'            Sql = Sql & ",codpostal = " & DBSet(Text1(4), "T") 'codigo postal
'            Sql = Sql & ",poblacion = " & DBSet(Text1(5), "T") 'poblacion
'            Sql = Sql & ",provincia = " & DBSet(Text1(6), "T") 'provincia
'            Sql = Sql & ",telefono1 = " & DBSet(Text1(8), "T") 'telefono1
'            Sql = Sql & ",telefono2 = " & DBSet(Text1(9), "T") 'telefono2
'            '[Monica]10/10/2018: solo en el caso de no ser picassent se actualiza el iban
'            If vParamAplic.Cooperativa <> 2 Then
'                Sql = Sql & ",iban = " & DBSet(Text1(27).Text & Text1(16).Text & Text1(17).Text & Text1(18).Text & Text1(19).Text, "T")
'            End If
'            Sql = Sql & " where ariagroId = " & DBSet(Text1(0).Text, "N")
'
'            conn.Execute Sql
        
        Case "E" ' eliminar
    End Select

'    CerrarConexionApp = True
    
    Exit Function

ePushSocio:
    MuestraError Err.Number, "Push Socio en App", Err.Description
End Function


Private Function ExisteAripush() As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error GoTo eExisteAripush

    Sql = "select * from aripush.usuariospush "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Set Rs = Nothing
    
    ExisteAripush = True
    Exit Function

eExisteAripush:
    ExisteAripush = False
End Function

Private Sub cmdAux_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 4 'Secciones
            Set frmSec = New frmManSeccion
            frmSec.DatosADevolverBusqueda = "0|1|"
            frmSec.CodigoActual = txtAux1(1).Text
            frmSec.Show vbModal
            Set frmSec = Nothing
            PonerFoco txtAux1(1)
            
            If txtAux1(1).Text <> "" Then txtAux1_LostFocus (1)
        
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
           
           PonerFormatoFecha txtAux1(indice)
           If txtAux1(indice).Text <> "" Then frmC.NovaData = CDate(txtAux1(indice).Text)
        
           Screen.MousePointer = vbDefault
           frmC.Show vbModal
           Set frmC = Nothing
           PonerFoco txtAux1(indice)
        
        Case 2, 3 'cuentas contables de cliente y proveedor
            If vSeccion Is Nothing Then Exit Sub
            
            indice = Index + 2
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = txtAux1(indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco txtAux1(indice)
        
        
        Case 5 'codigo de iva
            Set frmTIva = New frmTipIVAConta
            frmTIva.DeConsulta = True
            frmTIva.DatosADevolverBusqueda = "0|1|"
            frmTIva.CodigoActual = txtAux1(6).Text
            frmTIva.Show vbModal
            Set frmTIva = Nothing
            PonerFoco txtAux1(6)

    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

' *** si n'hi han combos a la capçalera ***
Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbLightBlue
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Combo1(Index).BackColor = vbLightBlue Then Combo1(Index).BackColor = vbWhite
    
    If Index = 0 And (Modo = 3 Or Modo = 4) Then
        '[Monica]12/03/2015: solo si no es Alzira lo dejamos como estaba, para el caso de Alzira siempre abierto
        If vParamAplic.Cooperativa <> 4 Then
            '[Monica]26/03/2014: añadimos que puedan ser tb ED (Combo1(0).ListIndex = 1) antes solo entidad
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

Private Sub ImprimirDocALTABaja(NroDoc As Long, esAlta As Boolean)
Dim vAux As String
Dim vCampos As String
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim indRPT As Byte
Dim nomDocu As String

    If NroDoc <> 0 Then
        vCampos = ""
        '[Monica]08/06/2018: ahora miramos de aportaciones
        'Sql = "select codcampo from rsocios_movim where numerodoc = " & DBSet(NroDoc, "N")
        Sql = "select numfactu from raportacion where nrodocum = " & DBSet(NroDoc, "N")
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            vCampos = vCampos & Rs.Fields(0).Value & ","
            Rs.MoveNext
        Wend
        
        Set Rs = Nothing
        If vCampos <> "" Then
            vCampos = Mid(vCampos, 1, Len(vCampos) - 1)
            
            InicializarVbles
            
            'Añadir el parametro de Empresa
            cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
            numParam = numParam + 1
            
            
            cadFormula = "{rcampos.codcampo} in [" & vCampos & "]  and {rsocios.codsocio} = " & DBSet(Text1(0).Text, "N")
            
            If esAlta Then
                cadTitulo = "Documento Alta de Campos"
            Else
                cadTitulo = "Documento Baja de Campos"
            End If
             
            cadParam = cadParam & "pFecha=""" & lw1.SelectedItem.SubItems(1) & """|"
            numParam = numParam + 1
        
            cadParam = cadParam & "pCausas=""" & lw1.SelectedItem.SubItems(3) & """|"
            numParam = numParam + 1
            
            If esAlta Then
                indRPT = 18 ' alta campos
            Else
                indRPT = 19 ' baja campos
            End If
            
            If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
            
            cadNombreRPT = nomDocu
            
            LlamarImprimir
        End If
        
    End If

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
        .Buttons(1).Image = 20  'baja de socio
        .Buttons(2).Image = 25  'imprimir fases
        .Buttons(3).Image = 16  'socios de baja
        .Buttons(4).Image = 34  'exportacion de socios
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
        Label10.Caption = "Fecha Revisión"
        Text1(7).Tag = "Fecha Revisión|F|S|||rsocios|fechanac|dd/mm/yyyy||"
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
    ' ******* si n'hi han llínies *******
    DataGridAux(0).ClearFields
    DataGridAux(1).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "rsocios"
    Ordenacion = " ORDER BY codsocio"
    '************************************************
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    '***** canviar el nom de la PK de la capçalera; repasar codEmpre *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where codsocio=-1"
    Data1.Refresh
       
    ModoLineas = 0
       
    ' **** si n'hi ha algun frame que no te datagrids ***
'    CargaFrame 3, False
    ' *************************************************
         
    CargaCombo
    
    '[Monica]08/04/2015: sólo Catadau ve el nro de asociado
    Label36.visible = (vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19)
    Text1(30).visible = (vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19)
    Text1(30).Enabled = (vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19)
    
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1 'búsqueda
        ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
        Text1(0).BackColor = vbLightBlue 'codclien
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
        cadB = "codsocio= " & Socio
        Text1(0).Text = Socio
        
        HacerBusqueda
        SSTab1.Tab = 3
        Toolbar2_ButtonClick Toolbar2.Buttons(11)
    End If

    '[Monica]04/09/2019: en el caso de Monasterios no tiene telefonos, pozos, miembros y movimientos
    If vParamAplic.Cooperativa = 17 Then
        Me.SSTab1.TabVisible(2) = False
        Me.SSTab1.TabVisible(4) = False
        Me.SSTab1.TabVisible(5) = False
        Me.SSTab1.TabVisible(6) = False
        
        Me.SSTab1.TabEnabled(2) = False
        Me.SSTab1.TabEnabled(4) = False
        Me.SSTab1.TabEnabled(5) = False
        Me.SSTab1.TabEnabled(6) = False
    End If

    '[Monica]26/07/2018: para todos quitamos los telefonos
    Me.SSTab1.TabVisible(2) = False
    Me.SSTab1.TabEnabled(2) = False
    
    '[Monica]25/02/2019: por defecto
    Me.Option1(0).Value = True

End Sub


Private Sub LimpiarCampos()
Dim i As Integer

    On Error Resume Next
    
    Timer1.Enabled = False
    Label31.visible = False
    
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
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
    DesplazamientoVisible b And Data1.Recordset.RecordCount > 1
    
    '---------------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    CmdAceptar.visible = b
       
    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a més neteja els camps Text1
    BloquearText1 Me, Modo
    BloquearCombo Me, Modo
    For i = 0 To chkAbonos.Count - 1
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
    DoEvents
    
    
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
            
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    PonerLongCampos

    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 0, False
        CargaGrid 1, False
        CargaGrid 2, False
        CargaGrid 3, False
        CargaGrid 4, False
    End If
    
    Me.imgDoc(0).Enabled = (Modo = 2 And (vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 16 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19))
    Me.imgDoc(0).visible = (Modo = 2 And (vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 16 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19))
    
    
    b = (Modo = 4) Or (Modo = 2)
    DataGridAux(0).Enabled = b
    DataGridAux(1).Enabled = b
    DataGridAux(2).Enabled = b
    DataGridAux(3).Enabled = b
    DataGridAux(4).Enabled = b
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
    For i = 1 To txtAux1.Count - 1
        BloquearTxt txtAux1(i), Not b
    Next i
    b = (Modo = 5) And (NumTabMto = 1) And ModoLineas = 2
    BloquearTxt txtAux1(1), b
    BloquearBtn cmdAux(4), b
    
    'pozos
    b = (Modo = 5) And NumTabMto = 2
    For i = 1 To 3
        BloquearTxt txtAux3(i), Not b
    Next i
    b = (Modo = 5) And NumTabMto = 2 And ModoLineas = 2
    BloquearTxt txtAux3(1), b
    
    'miembros
    b = (Modo = 5) And NumTabMto = 3
    For i = 2 To 5
        BloquearTxt txtAux4(i), Not b
    Next i
    
'    ' capital y votos enabled solo si el socio es entidad
'    Text1(28).Enabled = ((Modo = 1) Or ((Combo1(0).ListIndex = 2) And (Modo = 3 Or Modo = 4)))
'    Text1(29).Enabled = ((Modo = 1) Or ((Combo1(0).ListIndex = 2) And (Modo = 3 Or Modo = 4)))
     
     '-----------------------------
    PonerModoOpcionesMenu (Modo) 'Activar opcions menú según modo
    PonerOpcionesMenu   'Activar opcions de menú según nivell
                        'de permisos de l'usuari

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

Dim vCodSoc As String
    
    'Barra de CAPÇALERA
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
    
    vCodSoc = ""
    If Not Data1.Recordset.EOF Then vCodSoc = Data1.Recordset.Fields(0).Value
    
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And DatosADevolverBusqueda = "" 'And Not DeConsulta and DatosADevolverBusqueda = ""
    'Modificar
    Toolbar1.Buttons(2).Enabled = b And EsSocioCooperativa(vCodSoc)

    Me.mnModificar.Enabled = b And Not EsSocioCooperativa(vCodSoc)
    'eliminar
    Toolbar1.Buttons(3).Enabled = b And EsSocioCooperativa(vCodSoc)
    Me.mnEliminar.Enabled = b
    'dar de baja un socio
    Toolbar5.Buttons(1).Enabled = b And EsSocioCooperativa(vCodSoc)
    Me.mnBajaSocio.Enabled = b
    
    '[Monica]13/03/2018: solo para Coopic
    'informe de socios de baja
    Toolbar5.Buttons(3).Enabled = (vParamAplic.Cooperativa = 16)
    
    'Imprimir
    'Toolbar1.Buttons(12).Enabled = (b Or Modo = 0)
    Toolbar1.Buttons(8).Enabled = b
'    Toolbar1.Buttons(14).Enabled = b
       
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And DatosADevolverBusqueda = ""
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.Adoaux(i).Recordset.RecordCount > 0)
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i
    ' ****************************************
    ToolAux(4).Buttons(1).Enabled = False
    ToolAux(4).Buttons(3).Enabled = False
    
    
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
    DesplazamientoData Data1, Index, True
    PonerModoOpcionesMenu Modo
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
            
       Case 4 ' movimientos
            tabla = "rsocios_movim"
            Sql = "SELECT rsocios_movim.codsocio, rsocios_movim.numlinea, rsocios_movim.codcampo, rsocios_movim.supcoope, rsocios_movim.fecmovim, rsocios_movim.importe, rsocios_movim.causa, rsocios_movim.numerodoc "
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
        cadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        cadB = Aux
        '   Com la clau principal es única, en posar el sql apuntant
        '   al valor retornat sobre la clau ppal es suficient
        ' *** canviar o llevar el WHERE; repasar codEmpre ***
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
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

'[Monica]12/11/2018: ayuda para el codigo de poblacion
Private Sub frmPue_DatoSeleccionado(CadenaSeleccion As String)
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo pueblo
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre pueblo
End Sub

Private Sub frmPue2_DatoSeleccionado(CadenaSeleccion As String)
    Text1(24).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo pueblo
    Text1(23).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre pueblo
End Sub

Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
    txtAux1(1).Text = RecuperaValor(CadenaSeleccion, 1) 'codseccion
    FormateaCampo txtAux1(1)
    txtAux2(1).Text = RecuperaValor(CadenaSeleccion, 2) 'nomseccion
End Sub


Private Sub frmSit_DatoSeleccionado(CadenaSeleccion As String)
    Text1(15).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo situacion
    FormateaCampo Text1(15)
    Text2(15).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre situacion
End Sub

Private Sub frmSocPrev_DatoSeleccionado(CadenaSeleccion As String)
Dim cadB As String
    
    If CadenaSeleccion <> "" Then
        cadB = "codsocio = " & DBSet(RecuperaValor(CadenaSeleccion, 1), "N")
        
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmTIva_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento Tipo de iva
    txtAux1(6).Text = RecuperaValor(CadenaSeleccion, 1) 'codiva
    FormateaCampo txtAux1(6)
    txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'nomiva
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
End Sub

Private Sub Image1_DblClick()
Dim L As Long
Dim C As String
    
    
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
                vCadena = "Sólo está activo si el socio es una entidad o ED. En este caso, cuando " & vbCrLf & _
                          "esté marcado todas las facturas que se generen para este socio, se" & vbCrLf & _
                          "generarán como contabilizadas con el número que le corresponda del" & vbCrLf & _
                          "tipo de movimiento." & vbCrLf & vbCrLf & _
                          "Cuando se reciba el documento se podrá cambiar el nro de factura y " & vbCrLf & _
                          "contabilizar en un proceso específico. " & vbCrLf & vbCrLf
            Else
                vCadena = "Cuando esté marcado todas las facturas que se generen para este socio," & vbCrLf & _
                          "se generarán como contabilizadas con el número que le corresponda " & vbCrLf & _
                          "del tipo de movimiento." & vbCrLf & vbCrLf & _
                          "Cuando se reciba el documento se podrá cambiar el nro de factura y " & vbCrLf & _
                          "contabilizar en un proceso específico. " & vbCrLf & vbCrLf
            End If
                      
                      
        Case 1
           ' "____________________________________________________________"
            '[Monica]15/05/2018:
                vCadena = "Si está marcado cuando se integren las facturas de anticipo y liquidación no " & vbCrLf & _
                          "se grabará el IBAN en el pago." & vbCrLf & _
                          "" & vbCrLf & vbCrLf
                      
                      
                      
    End Select
    MsgBoxA vCadena, vbInformation, "Descripción de Ayuda"
    
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
            Case 4
                indice = 31
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
        '[Monica]15/01/2018: faltaba RTA, [Monica]19/04/2018: y todas las demas rectificativas
        If lw1.SelectedItem.Text = "RCP" Or lw1.SelectedItem.Text = "RMP" Or lw1.SelectedItem.Text = "TAL" Or lw1.SelectedItem.Text = "RVP" Or lw1.SelectedItem.Text = "RMT" Or lw1.SelectedItem.Text = "RTA" Or _
            lw1.SelectedItem.Text = "RRC" Or lw1.SelectedItem.Text = "RRM" Or lw1.SelectedItem.Text = "RRV" Or lw1.SelectedItem.Text = "FIN" Then
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
    
    
    Case 6
        'documento de baja de campos
        ImprimirDocALTABaja lw1.SelectedItem.Text, lw1.SelectedItem.SubItems(2) > 0
    
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

Private Sub mnSociosBaja_Click()
    AbrirListado (52)
End Sub

Private Sub mnExportacionSocios_Click()
    AbrirListado (53)
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



Private Sub Option1_Click(Index As Integer)
   
'    If Button.Tag = "" Then Exit Sub
'
'    If Button.Index = 9 Then
'        ImprimirDocumentos
'        Exit Sub
'    End If
'
'
'    Label16.Caption = ""
'    'Levantamos todos los botones y dejamos pulsado el de ahora
'    For NumRegElim = 1 To Toolbar2.Buttons.Count
'        If Toolbar2.Buttons(NumRegElim).Tag <> "" Then
'            If Toolbar2.Buttons(NumRegElim).Index <> Button.Index Then Toolbar2.Buttons(NumRegElim).Value = tbrUnpressed
'        End If
'    Next NumRegElim
'    CargaColumnas CByte(Button.Tag)
    
    If Index = 4 Then
        ImprimirDocumentos
        Exit Sub
    End If
    
    
    CargaColumnas CByte(Index)
        
    
    'Hacemos las acciones
    If Modo = 2 Then CargaDatosLW
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
        Case 5  'Búscar
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
' ***** Si la clau primaria de la capçalera no es Text1(0), canviar-ho en <=== *****
    Timer1.Enabled = False
    Label31.visible = False
    
    
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonerFoco Text1(0) ' <===
        Text1(0).BackColor = vbLightBlue ' <===
        ' *** si n'hi han combos a la capçalera ***
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbLightBlue
            PonerFoco Text1(kCampo)
        End If
    End If
' ******************************************************************************
End Sub

Private Sub HacerBusqueda()

    cadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        ' *** foco al 1r camp visible de la capçalera que siga clau primaria ***
        PonerFoco Text1(0)
        ' **********************************************************************
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
    
    Set frmSocPrev = New frmBasico2
    
    AyudaSociosPrev frmSocPrev, , cadB
    
    Set frmSocPrev = Nothing

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
    cadB = ""
    
    PonerModo 0
    
    If chkVistaPrevia.Value = 1 Then
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
    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
        Text1(0).Text = DevuelveValor("select max(coalesce(codsocio,0))+ 1 from rsocios where codsocio < " & DBSet(cMaxSocio, "N"))
    End If
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

    '[Monica]20/04/2018: faltaba los valores de telefono
    Tel1Ant = Text1(8).Text
    Tel2Ant = Text1(9).Text


    FecBajaAnt = Text1(14).Text

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

    '[Monica]21/09/2016:
    If Text1(14).Text <> "" Then
        MsgBox "Este socio ya está dado de baja.", vbExclamation
        Exit Sub
    End If


    Screen.MousePointer = vbHourglass
    Text1(0).Text = Data1.Recordset!Codsocio
    
    frmListado.NumCod = Text1(0).Text
    frmListado.Opcionlistado = 23
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
'        If i <> 4 Then
            CargaGrid i, True
            If Not Adoaux(i).Recordset.EOF Then _
                PonerCamposForma2 Me, Adoaux(i), 2, "FrameAux" & i
'        End If
    Next i
    ' *******************************************

    ' *** si n'hi han llínies sense datagrid ***
    ' ***************************************
    
    ' ************* configurar els camps de les descripcions de la capçalera *************
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
                    If NumTabMto = 0 Or NumTabMto = 1 Or NumTabMto = 2 Or NumTabMto = 4 Or NumTabMto = 3 Then
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
                                For i = 0 To txtAux1.Count - 1
                                    txtAux1(i).Text = ""
                                    BloquearTxt txtAux1(i), True
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
                                For i = 0 To txtAux4.Count
                                    BloquearTxt txtAux4(i), True
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
                    
                    If Not Adoaux(NumTabMto).Recordset.EOF Then
                        Adoaux(NumTabMto).Recordset.MoveFirst
                    End If

                Case 2 'modificar llínies
                    ModoLineas = 0
                    
                    ' *** si n'hi han tabs ***
                    SituarTab (NumTabMto)
                    'SSTab1.Tab = 1
                    'SSTab2.Tab = NumTabMto
                    ' ***********************
                    
                    PonerModo 4
                    If Not Adoaux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de llínies ***
                        V = Adoaux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
                        Adoaux(NumTabMto).Recordset.Find (Adoaux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
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
            If Not Adoaux(NumTabMto).Recordset.EOF Then
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
    
    '[Monica]12/06/2018: en el caso de que sean coopic-picassent no pueden insertar socios superiores a cMaxSocio=10000
    If b And Modo = 3 And (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16) And CLng(Text1(0).Text) > cMaxSocio Then
        MsgBox "No puede insertar un socio con codigo superior a " & cMaxSocio & ".", vbExclamation
        PonerFoco Text1(0)
        b = False
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
            Else
'                '[Monica]20/11/2013: añadimos el tema de la comprobacion del IBAN
'                If Not Comprueba_CC_IBAN(cta, Text1(42).Text) Then
'                    cadMen = "La cuenta IBAN del cliente no es correcta. ¿ Desea continuar ?."
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
                '[Monica]12/04/2017: solo en el caso de que no sea natural dejamos introducir un codigo, añadida la condicion
                If vParamAplic.Cooperativa <> 9 Then
                    Sql = "select count(*) from rsocios where codsocio <> " & DBSet(Text1(0).Text, "N") & " and codigoggap = " & DBSet(Text1(26).Text, "T")
                    
                    If TotalRegistros(Sql) <> 0 Then
                        MsgBox "Este código de GlobalGap ya está asignado a otro socio. Revise.", vbExclamation
                        PonerFoco Text1(26)
                    End If
                End If
            End If
        End If
        
        '[Monica]29/09/2014: comprobamos si vamos a dar de baja que no tenga fecha de alta en programa operativo en ninguno de sus campos
        If b Then
            If Text1(14).Text <> "" Then
                Sql = "select count(*) from rcampos where codsocio = " & DBSet(Text1(0).Text, "N") & " and not fecaltapropera is null "
                If TotalRegistros(Sql) <> 0 Then
                    MsgBox "Este socio tiene algún campo dado de alta en el programa operativo." & vbCrLf & vbCrLf & "No puede darlo de baja. Revise.", vbExclamation
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


Private Function Eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar

    conn.BeginTrans
    ' ***** canviar el nom de la PK de la capçalera, repasar codEmpre *******
    vWhere = " WHERE codsocio=" & Data1.Recordset!Codsocio
        ' ***********************************************************************
        
    ' ***** elimina les llínies ****
    conn.Execute "DELETE FROM rsocios_seccion " & vWhere
        
    conn.Execute "DELETE FROM rsocios_telefonos " & vWhere
    
    '[Monica]10/05/2017: faltarian eliminar la lineas
    conn.Execute "DELETE FROM rsocios_miembros " & vWhere
        
    conn.Execute "DELETE FROM rsocios_pozos " & vWhere
    
    conn.Execute "DELETE FROM rsocios_movim " & vWhere
    
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
                    cadMen = "No existe la Situación: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBoxA(cadMen, vbQuestion + vbYesNo) = vbYes Then
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
            If Modo = 1 Then Exit Sub
            '[Monica]24/10/2013: no tenia que dar el mensaje de dentro de campaña
            PonerFormatoFecha Text1(Index) ', True
            
        Case 13, 14, 31 'Fechas
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
                        '[Monica]12/04/2017: solo en el caso de que no sea natural dejamos introducir un codigo, añadida la condicion
                        If vParamAplic.Cooperativa <> 9 Then
                    
                            Sql = "select count(*) from rsocios where codsocio <> " & DBSet(Text1(0).Text, "N") & " and codigoggap = " & DBSet(Text1(Index).Text, "T")
                            
                            If TotalRegistros(Sql) <> 0 Then
                                MsgBox "Este código de GlobalGap ya está asignado a otro socio. Revise.", vbExclamation
                                PonerFoco Text1(Index)
                            End If
                            
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
            imgZoom_Click (0)
        Else
            KEYpress KeyAscii
        End If
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub


Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
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
Dim Eliminar As Boolean

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
    Eliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 0 'telefonos
            Sql = "¿Seguro que desea eliminar el telefono?"
            Sql = Sql & vbCrLf & "Teléfono: " & Adoaux(Index).Recordset!idtelefono & " - " & Adoaux(Index).Recordset!imei
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM rsocios_telefonos"
                Sql = Sql & vWhere & " AND idtelefono= " & DBLet(Adoaux(Index).Recordset!idtelefono, "T")
                
                
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
            Sql = Sql & vbCrLf & "Sección: " & Adoaux(Index).Recordset!codsecci & " - " & Adoaux(Index).Recordset!nomsecci
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM rsocios_seccion"
                Sql = Sql & vWhere & " AND codsecci= " & DBLet(Adoaux(Index).Recordset!codsecci, "N")
            
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
            Sql = Sql & vbCrLf & "Numero Fase: " & Adoaux(Index).Recordset!numfases
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM rsocios_pozos"
                Sql = Sql & vWhere & " AND numfases= " & DBLet(Adoaux(Index).Recordset!numfases, "N")
                
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
            Sql = "¿Seguro que desea eliminar el miembro?"
            Sql = Sql & vbCrLf & "NIF: " & Adoaux(Index).Recordset!nifmiembro
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM rsocios_miembros"
                Sql = Sql & vWhere & " AND numlinea= " & DBLet(Adoaux(Index).Recordset!NumLinea, "N")
                
                CadenaCambio = "DELETE FROM rsocios_miembros " & vWhere
                '------------------------------------------------------------------------------
                '  LOG de acciones
                ValorAnterior = ""
                Set LOG = New cLOG
                LOG.InsertarCambiosRegistros 14, vUsu, "Cambio Socios-D", "rsocios_miembros", ObtenerWhereCab(False)
                Set LOG = Nothing
                '-----------------------------------------------------------------------------
            End If
        
        Case 4 'movimientos
            Sql = "¿Seguro que desea eliminar el campo de baja?"
            Sql = Sql & vbCrLf & "Campo: " & Adoaux(Index).Recordset!codCampo
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM rsocios_movim"
                Sql = Sql & vWhere & " AND numlinea= " & DBLet(Adoaux(Index).Recordset!NumLinea, "N")
                
                CadenaCambio = "DELETE FROM rsocios_miembros " & vWhere
                '------------------------------------------------------------------------------
                '  LOG de acciones
                ValorAnterior = ""
                Set LOG = New cLOG
                LOG.InsertarCambiosRegistros 14, vUsu, "Cambio Socios-D", "rsocios_movim", ObtenerWhereCab(False)
                Set LOG = Nothing
                '-----------------------------------------------------------------------------
            End If
        
        
        
    End Select

    If Eliminar Then
        NumRegElim = Adoaux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        conn.Execute Sql
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        If Index <> 4 Then _
            CargaGrid Index, True
        ' ***************************************************
        If Not SituarDataTrasEliminar(Adoaux(Index), NumRegElim, True) Then
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
        Case 0: vtabla = "rsocios_telefonos"
        Case 1: vtabla = "rsocios_seccion"
        Case 2: vtabla = "rsocios_pozos"
        Case 3: vtabla = "rsocios_miembros"
        Case 4: vtabla = "rsocios_movim"
    End Select
    ' ********************************************************
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            If Index <> 4 Then ' *** els index als que no volem sugerir-li un codi ***
                NumF = SugerirCodigoSiguienteStr(vtabla, "idtelefono", vWhere)
            Else
                NumF = ""
            End If
            ' ***************************************************************

            AnyadirLinea DataGridAux(Index), Adoaux(Index)
    
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
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
'            If Index <> 4 Then ' *** els index als que no volem sugerir-li un codi ***
'                NumF = SugerirCodigoSiguienteStr(vTabla, "codsecci", vWhere)
'            Else
                NumF = ""
'            End If
            ' ***************************************************************

            AnyadirLinea DataGridAux(Index), Adoaux(Index)
    
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
                    For i = 0 To txtAux1.Count - 1
                        txtAux1(i).Text = ""
                    Next i
                    txtAux1(0).Text = Text1(0).Text 'codsocio
                    txtAux1(1).Text = NumF 'codseccion
                    txtAux2(1).Text = ""
                    txtAux2(4).Text = ""
                    txtAux2(5).Text = ""
                    txtAux2(0).Text = ""
                    PonerFoco txtAux1(1)
                    
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

            AnyadirLinea DataGridAux(Index), Adoaux(Index)
    
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 240
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
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            If Index <> 4 Then ' *** els index als que no volem sugerir-li un codi ***
                NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)
            Else
                NumF = ""
            End If
            ' ***************************************************************
            
            AnyadirLinea DataGridAux(Index), Adoaux(Index)
    
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 240
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineas Index, ModoLineas, anc
        
            For i = 0 To txtAux4.Count - 1
                txtAux4(i).Text = ""
            Next i
            
            txtAux4(0).Text = Text1(0).Text 'codsocio
            txtAux4(1).Text = NumF 'numero de fase
            PonerFoco txtAux4(2) 'nif
            
            
        Case 4
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)
            ' ***************************************************************
            
            AnyadirLinea DataGridAux(Index), Adoaux(Index)
    
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 240
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineas Index, ModoLineas, anc
        
            For i = 0 To txtAux4.Count - 1
                txtAux5(i).Text = ""
            Next i
            
            txtAux5(0).Text = Text1(0).Text 'codsocio
            txtAux5(1).Text = NumF 'numero de fase
            PonerFoco txtAux5(2) 'nif
            
            
    End Select
End Sub


Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim i As Integer
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
    BloquearTxt Text1(0), True
    ' *********************************
  
    Select Case Index
        Case 0, 1, 2, 3, 4 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
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
                txtAux1(i).Text = DataGridAux(Index).Columns(i).Text
            Next i
            txtAux2(1).Text = DataGridAux(Index).Columns(2).Text
            For i = 3 To 7
                txtAux1(i - 1).Text = DataGridAux(Index).Columns(i).Text
            Next i
        
            CargarValoresAnteriores Me, 2, "FrameAux1"
        
        
        Case 2 'pozos
            For i = 0 To 3
                txtAux3(i).Text = DataGridAux(Index).Columns(i).Text
            Next i
        
            CargarValoresAnteriores Me, 2, "FrameAux2"
        
        Case 3 'miembros
            For i = 0 To 5
                txtAux4(i).Text = DataGridAux(Index).Columns(i).Text
            Next i
        
            CargarValoresAnteriores Me, 2, "FrameAux3"
        
        Case 4 'movimientos
            For i = 0 To 7
                txtAux5(i).Text = DataGridAux(Index).Columns(i).Text
            Next i
        
            CargarValoresAnteriores Me, 2, "FrameAux4"
        
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'telefonos
            PonerFoco txtAux(2)
        Case 1 'secciones
            PonerFoco txtAux1(2)
            If txtAux1(1).Text <> "" Then
                Set vSeccion = New CSeccion
                If vSeccion.LeerDatos(txtAux1(1)) Then
                    If vSeccion.AbrirConta Then
                        If txtAux1(4).Text <> "" Then
                            txtAux2(4).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", txtAux1(4).Text, "T")
                        End If
                        If txtAux1(5).Text <> "" Then
                            txtAux2(5).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", txtAux1(5).Text, "T")
                        End If
                        If txtAux1(6).Text <> "" Then
                            txtAux2(0).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", txtAux1(6).Text, "N")
                        End If
                    End If
                End If
            End If
        Case 2 ' pozos
            PonerFoco txtAux3(2)
        Case 3 ' miembros
            PonerFoco txtAux4(2)
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
            For jj = 1 To txtAux1.Count - 1
                txtAux1(jj).visible = b
                txtAux1(jj).Top = alto
            Next jj
            txtAux2(1).visible = b
            txtAux2(1).Top = alto
        
            For jj = 0 To cmdAux.Count - 1
                cmdAux(jj).visible = b
                cmdAux(jj).Top = txtAux1(3).Top
                cmdAux(jj).Height = txtAux1(3).Height
            Next jj
            
        Case 2 ' pozos
            For jj = 1 To 3
                txtAux3(jj).visible = b
                txtAux3(jj).Top = alto
            Next jj
    
        Case 3 ' miembros
            For jj = 2 To 5
                txtAux4(jj).visible = b
                txtAux4(jj).Top = alto
            Next jj
            
        Case 4 ' movim
            For jj = 2 To 5
                txtAux5(jj).visible = b
                txtAux5(jj).Top = alto
            Next jj
    
    End Select
End Sub






Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 ' Baja de socios
            mnBajaSocio_Click
        Case 2 'Imprimir fases
            mnImprimirFases_Click
        Case 3 ' socios dados de baja
            mnSociosBaja_Click
        Case 4 ' exportacion de socios a csv/xls
            mnExportacionSocios_Click
            
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
        Sql = DevuelveDesdeBDNew(cAgro, "rsocios_pozos", "acciones", "codsocio", txtAux3(0).Text, "N", , "numfases", txtAux3(1).Text, "N")
        If Sql <> "" Then
            MsgBox "El número de fase ya existe. Reintroduzca.", vbExclamation
            b = False
            PonerFoco txtAux3(1)
        End If
    End If
    
    If b And NumTabMto = 3 And ModoLineas = 1 Then
        Sql = ""
        Sql = DevuelveDesdeBDNew(cAgro, "rsocios_miembros", "nifmiembro", "codsocio", txtAux4(0).Text, "N", , "nifmiembro", txtAux4(2).Text)
        If Sql <> "" Then
            MsgBox "El nif ya existe para este socio. Reintroduzca.", vbExclamation
            b = False
            PonerFoco txtAux4(2)
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
            
            frmGlo.DeConsulta = False
            frmGlo.DatosADevolverBusqueda = "0|1|"
            frmGlo.CodigoActual = Text1(indice).Text
            frmGlo.Show vbModal
            
            Set frmGlo = Nothing

            Set frmGlo = Nothing
            PonerFoco Text1(indice)
            
         '[Monica]12/11/2018: buscar poblacion
         Case 3 ' buscar poblacion
            indice = 4
            Set frmPue = New frmManPueblos
            frmPue.DeConsulta = False
            frmPue.DatosADevolverBusqueda = "0|1|"
            frmPue.CodigoActual = Text1(indice).Text
            frmPue.Show vbModal
            Set frmPue = Nothing
            
            PonerFoco Text1(indice)
            
         Case 4 ' buscar poblacion
            indice = 24
            Set frmPue2 = New frmManPueblos
            frmPue2.DeConsulta = False
            frmPue2.DatosADevolverBusqueda = "0|1|"
            frmPue2.CodigoActual = Text1(indice).Text
            frmPue2.Show vbModal
            Set frmPue2 = Nothing
    
            PonerFoco Text1(indice)
    End Select
    
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    txtAux1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codiva
    FormateaCampo txtAux1(indice)
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
                    If vSeccion.LeerDatos(Adoaux(1).Recordset!codsecci) Then
                        If vSeccion.AbrirConta Then
                            If DBLet(Adoaux(1).Recordset!codmaccli, "T") <> "" Then
                                txtAux2(4).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Adoaux(1).Recordset!codmaccli, "T")
                            End If
                            If DBLet(Adoaux(1).Recordset!codmacpro, "T") <> "" Then
                                txtAux2(5).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Adoaux(1).Recordset!codmacpro, "T")
                            End If
                            
                            txtAux2(0).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", CStr(Adoaux(1).Recordset!CodIva), "N")
                            
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

    Adoaux(Index).ConnectionString = conn
    Adoaux(Index).RecordSource = MontaSQLCarga(Index, enlaza)
    Adoaux(Index).CursorType = adOpenDynamic
    Adoaux(Index).LockType = adLockPessimistic
    Adoaux(Index).Refresh
    
    If Not Adoaux(Index).Recordset.EOF Then
        PonerCamposForma2 Me, Adoaux(Index), 2, "FrameAux" & Index
        ' *** si n'hi han tabs sense datagrids, li pose els valors als camps ***
        If (Index = 4) Then 'datos facturacion
            tip = Adoaux(Index).Recordset!tipclien
            If (tip = 1) Then 'persona
                txtAux2(27).Text = Adoaux(Index).Recordset!ape_raso & "," & Adoaux(Index).Recordset!Nom_Come
            ElseIf (tip = 2) Then 'empresa
                txtAux2(27).Text = Adoaux(Index).Recordset!Nom_Come
            End If
            txtAux2(28).Text = DBLet(Adoaux(Index).Recordset!desforpa, "T")
            txtAux2(29).Text = DBLet(Adoaux(Index).Recordset!desrutas, "T")
            'txtAux2(31).Text = DBLet(AdoAux(Index).Recordset!comision, "T") & " %"
            txtAux2(32).Text = DBLet(Adoaux(Index).Recordset!nomrapel, "T")
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
    
    CargaGridGnral Me.DataGridAux(Index), Me.Adoaux(Index), tots, PrimeraVez
    
    
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

            If (enlaza = True) And (Not Adoaux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
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
            tots = "N||||0|;S|txtaux1(1)|T|Código|800|;S|cmdAux(4)|B|||;" 'codsocio,codsecci
            tots = tots & "S|txtAux2(1)|T|Nombre|4400|;"
            tots = tots & "S|txtaux1(2)|T|Fecha Alta|1900|;S|cmdAux(0)|B|||;"
            tots = tots & "S|txtaux1(3)|T|Fecha Baja|1900|;S|cmdAux(1)|B|||;"
            tots = tots & "S|txtaux1(4)|T|Cta.Cliente|2000|;S|cmdAux(2)|B|||;"
            tots = tots & "S|txtaux1(5)|T|Cta.Proveedor|2000|;S|cmdAux(3)|B|||;"
            tots = tots & "S|txtaux1(6)|T|Iva|800|;S|cmdAux(5)|B|||;"
            arregla tots, DataGridAux(Index), Me, 350
        
            DataGridAux(Index).Columns(2).Alignment = dbgLeft
            DataGridAux(Index).Columns(5).Alignment = dbgLeft
            DataGridAux(Index).Columns(6).Alignment = dbgLeft
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
'            BloquearTxt txtAux(14), Not b
'            BloquearTxt txtAux(15), Not b

            If (enlaza = True) And (Not Adoaux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
'                txtAux2(4).Text = PonerNombreCuenta(txtaux1(4), Modo)
'                txtAux2(5).Text = PonerNombreCuenta(txtaux1(5), Modo)
'                txtAux2(0).Text = PonerNombreDeCod(txtaux1(6), "tiposiva", "nombriva", "codigiva", "N", cConta)
            Else
                For i = 0 To 6
                    txtAux1(i).Text = ""
                Next i
                txtAux2(0).Text = ""
                txtAux2(1).Text = ""
                txtAux2(4).Text = ""
                txtAux2(5).Text = ""
            End If
        
        Case 2 'pozos
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;S|txtAux3(1)|T|Fases|1400|;" 'codsocio,numfase
            tots = tots & "S|txtAux3(2)|T|Acciones|1700|;"
            tots = tots & "S|txtAux3(3)|T|Observaciones|6280|;"
            arregla tots, DataGridAux(Index), Me, 350
        
            DataGridAux(Index).Columns(3).Alignment = dbgLeft
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))

        Case 3 'miembros de entidades
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;S|txtAux4(2)|T|NIF|1450|;" 'codsocio,numfase
            tots = tots & "S|txtAux4(3)|T|Nombre|7450|;"
            tots = tots & "S|txtAux4(4)|T|Votos|1800|;"
            tots = tots & "S|txtAux4(5)|T|Acciones|1900|;"
            arregla tots, DataGridAux(Index), Me, 350
        
            DataGridAux(Index).Columns(3).Alignment = dbgLeft
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
        
        Case 4 'movimientos
            'si es visible|control|tipo campo|nombre campo|ancho control
            tots = "N||||0|;N||||0|;S|txtAux5(2)|T|Campo|1500|;" 'codsocio,numfase
            tots = tots & "S|txtAux5(3)|T|Superficie|1700|;"
            tots = tots & "S|txtAux5(4)|T|Fecha|1650|;"
            tots = tots & "S|txtAux5(5)|T|Importe|1750|;"
            tots = tots & "S|txtAux5(6)|T|Causa|6000|;"
            tots = tots & "N|txtAux5(7)|T|Doc|1000|;"
            
            arregla tots, DataGridAux(Index), Me, 350
        
            DataGridAux(Index).Columns(3).Alignment = dbgLeft
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))



    End Select
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
    ' **** si n'hi han llínies en grids i camps fora d'estos ****
    If Not Adoaux(Index).Recordset.EOF Then
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
               txtAux2(4).Text = PonerNombreCuenta(txtAux1(4), 3, Text1(0))
               txtAux2(5).Text = PonerNombreCuenta(txtAux1(5), 3, Text1(0))
            End If
            
            If NumTabMto = 1 Then
                If Not vSeccion Is Nothing Then
                    vSeccion.CerrarConta
                    Set vSeccion = Nothing
                End If
            End If
            
            
            Select Case NumTabMto
                Case 0, 1, 2, 3, 4 ' *** els index de les llinies en grid (en o sense tab) ***
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
                   txtAux2(4).Text = PonerNombreCuenta(txtAux1(4), 4, Text1(0))
                   txtAux2(5).Text = PonerNombreCuenta(txtAux1(5), 4, Text1(0))
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
                V = Adoaux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
                CargaGrid NumTabMto, True
            End If

            ' *** si n'hi han tabs ***
            SituarTab (NumTabMto)

            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            If NumTabMto <> 3 Then
                DataGridAux(NumTabMto).SetFocus
                Adoaux(NumTabMto).Recordset.Find (Adoaux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
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
            For i = 0 To txtAux1.Count - 1
                txtAux1(i).Text = ""
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
        If cadB <> "" Then
            .cadRegSelec = SQL2SF(cadB)
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
    Combo1(1).AddItem "No productor"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 4
   
    'tipo de relacion con la cooperativa
    Combo1(2).AddItem "Socio"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 0
    Combo1(2).AddItem "Asociado"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 1
    Combo1(2).AddItem "Tercero"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 2
  
    'tipo de contrato
    Combo1(3).AddItem "Sin contrato"
    Combo1(3).ItemData(Combo1(3).NewIndex) = 0
    Combo1(3).AddItem "3 años"
    Combo1(3).ItemData(Combo1(3).NewIndex) = 1
    Combo1(3).AddItem "5 años"
    Combo1(3).ItemData(Combo1(3).NewIndex) = 2
  
  
End Sub


Private Sub txtAux1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim cadena As String
    
    If Not PerderFocoGnral(txtAux1(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 1 ' seccion
                If PonerFormatoEntero(txtAux1(Index)) Then
                    Set vSeccion = New CSeccion
                    If vSeccion.LeerDatos(txtAux1(Index)) Then
                        txtAux2(Index).Text = vSeccion.Nombre
                        If vSeccion.AbrirConta Then
                        
                            ' si estamos insertando montamos las cuentas contables con las raices
                            ' y el codigo
                            
                            cadena = String(vEmpresa.DigitosUltimoNivel - vEmpresa.DigitosNivelAnterior, "0")
                            
                            '18/09/2009
                            txtAux1(4).Text = vSeccion.RaizSocio & Format(txtAux1(0).Text, cadena)
                            txtAux1(5).Text = vSeccion.RaizProv & Format(txtAux1(0).Text, cadena)
                        End If
                    Else
                        Set vSeccion = Nothing
                        cadMen = "No existe la Sección: " & txtAux1(Index).Text & vbCrLf
                        cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                        If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                            Set frmSec = New frmManSeccion
                            frmSec.DatosADevolverBusqueda = "0|1|"
                            frmSec.NuevoCodigo = txtAux1(Index).Text
                            txtAux1(Index).Text = ""
                            TerminaBloquear
                            frmSec.Show vbModal
                            Set frmSec = Nothing
                            If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                        Else
                            txtAux1(Index).Text = ""
                        End If
                    End If
                Else
                    txtAux1(Index).Text = ""
                End If
        
            
        Case 2, 3 'fecha de alta y de baja
            PonerFormatoFecha txtAux1(Index), True
            
        Case 4, 5 'cta Cliente y Proveedor
            If txtAux1(Index).Text = "" Then Exit Sub
            
            If Not vSeccion Is Nothing Then
                txtAux2(Index).Text = PonerNombreCuenta(txtAux1(Index), Modo)
                If txtAux1(Index).Text <> "" Then
                    If Not vSeccion.CtaConRaizCorrecta(txtAux1(Index).Text, Index - 4) Then
                        MsgBox "La cuenta no tiene la raiz correcta. Revise.", vbExclamation
                    Else
                        ' si la cuenta es correcta y no existe la insertamos en contabilidad
                        txtAux2(Index).Text = PonerNombreCuenta(txtAux1(Index), 3, Text1(0))
                    End If
                End If
            End If

        Case 6 'codigo iva
            If txtAux1(Index).Text = "" Then Exit Sub
            
            If Not vSeccion Is Nothing Then
                  txtAux2(0).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", txtAux1(Index).Text, "N")
            End If
            CmdAceptar.SetFocus

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
            CmdAceptar.SetFocus

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
        .Buttons(13).Image = 27
    End With
    With Me.Toolbar3
        .ImageList = frmPpal.imgListImages16
        .Buttons(1).Image = 5
        .Buttons(3).Image = 7
        .Buttons(5).Image = 6
        .Buttons(7).Image = 8
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
Dim C As ColumnHeader

    Frame8.visible = False

    Select Case OpcionList
    Case 0
        'Facturas
        Label16.Caption = "Facturas"
        Columnas = "Tipo|Numero|Fecha|Importe|"
        Ancho = "1000|2000|1400|3600|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|1|"
        'Formatos
        Formato = "|00000000|dd/mm/yyyy|" & FormatoImporte & "|"
        Ncol = 4
               
    Case 1
        'Entradas por Huerto
        Label16.Caption = "Entradas por Huerto"
        Columnas = "Huerto|Nro.Orden|Partida|Variedad|Kilos Netos|"
        Ancho = "1200|1200|2400|2200|1700|"
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
            Ancho = "1400|1400|2000|1500|2100|"
            'vwColumnRight =1  left=0   center=2
            Alinea = "0|0|0|1|1|"
            'Formatos
            Formato = "|dd/mm/yyyy|0|###,##0|###,###,##0|"
            Ncol = 5
        Else
            Columnas = "Albarán|Fecha|Campo|N.Orden|Partida|Variedad|Kilos Netos|"
            Ancho = "1100|1400|1100|1000|1500|1500|1400|"
            'vwColumnRight =1  left=0   center=2
            Alinea = "0|0|0|0|0|0|1|"
            'Formatos
            Formato = "|dd/mm/yyyy|0|0|0|0|###,###,##0|"
            Ncol = 7
        End If
        
    Case 3
        'Facturas adv , retirada almazara y bodega , recibos de pozos
        Label16.Caption = "Facturas Cliente"
        Columnas = "Tipo|Factura|Fecha|Importe|Cobrado|"
        Ancho = "1000|1500|1500|2600|2000|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|0|1|0|"
        'Formatos
        Formato = "|00000000|dd/mm/yyyy|" & FormatoImporte & "||"
        Ncol = 5
        
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
        
    Case 6
        'Documentos de baja de campos
        Label16.Caption = "Documentos de alta/baja de campos"
        Columnas = "Documento|Fecha|Importe|Concepto|"
        Ancho = "1800|1400|1600|3600|"
        'vwColumnRight =1  left=0   center=2
        Alinea = "0|0|1|0|"
        'Formatos
        Formato = "00000000|dd/mm/yyyy|" & FormatoImporte & "||"
        Ncol = 4
        
    End Select
    
    
'    'Fecha incio busquedas
'    Text1(46).Text = Format(imgFecha(3).Tag, "dd/mm/yyyy")
    'Guardo la opcion en el tag
    lw1.Tag = OpcionList & "|" & Ncol & "|"
    
    lw1.ColumnHeaders.Clear
    
    For NumRegElim = 1 To Ncol
         Set C = lw1.ColumnHeaders.Add()
         C.Text = RecuperaValor(Columnas, CInt(NumRegElim))
         C.Width = RecuperaValor(Ancho, CInt(NumRegElim))
         C.Alignment = Val(RecuperaValor(Alinea, CInt(NumRegElim)))
         C.Tag = RecuperaValor(Formato, CInt(NumRegElim))
    Next NumRegElim
End Sub


Private Sub CargaDatosLW()
Dim C As String
Dim bs As Byte
    bs = Screen.MousePointer
    C = Me.lblIndicador.Caption
    lblIndicador.Caption = "Leyendo " & Label16.Caption
    lblIndicador.Refresh
    CargaDatosLW2
    Me.lblIndicador.Caption = C
    Screen.MousePointer = bs
End Sub

Private Sub CargaDatosLW2()
Dim Cad As String
Dim Rs As ADODB.Recordset
Dim It As ListItem
Dim ElIcono As Integer
Dim GroupBy As String
Dim Orden As String
Dim C As String
Dim i As Integer

    On Error GoTo ECargaDatosLW
    
    If Modo <> 2 Then Exit Sub
    
'    For NumRegElim = 1 To Toolbar2.Buttons.Count
'        If Toolbar2.Buttons(NumRegElim).Value = tbrPressed Then
'            ElIcono = Toolbar3.Buttons(NumRegElim).Image
'            Exit For
'        End If
'    Next
    
    For NumRegElim = 0 To 6
        If Option1(NumRegElim).Value Then
            ElIcono = Toolbar3.Buttons(NumRegElim + 1).Image
            i = NumRegElim
            Select Case i
                Case 0
                    Me.Image2.Picture = frmPpal.imgListPpal.ListImages(23).Picture
                Case 2
                    Me.Image2.Picture = frmPpal.imgListPpal.ListImages(30).Picture
                Case 3
                    Me.Image2.Picture = frmPpal.imgListPpal.ListImages(22).Picture
                Case Else
                    Me.Image2.Picture = LoadPicture("")
            End Select
        
            Exit For
        End If
    Next
    'Fecha incio busquedas
    Text3(0).Text = Format(imgFec(3).Tag, "dd/mm/yyyy")
    
    
    
    Select Case CByte(i) 'CByte(RecuperaValor(lw1.Tag, 1))
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
        
    
    Case 3
        'FACTURAS de cliente (advfacturas, rbodfacturas, rrecibpozos)
        'advfacturas
        Dim ContaADV As Integer
        ContaADV = DevuelveValor("select empresa_conta from rseccion where codsecci = " & vParamAplic.SeccionADV)
        
        If ContaADV <> 0 Then
            '[Monica]19/04/2018: contabilidad nueva
            If vParamAplic.ContabilidadNueva Then
                Cad = "select h.codtipom,h.numfactu,h.fecfactu,h.totalfac totalfac, if(sum(coalesce(c1.impvenci,0)+coalesce(c1.gastos,0)-coalesce(c1.impcobro,0))=0,'Cobrado','')  "
                Cad = Cad & " from (advfacturas h inner join usuarios.stipom u1 on h.codtipom = u1.codtipom) left join ariconta" & ContaADV & ".cobros c1 on c1.numserie = u1.letraser "
                Cad = Cad & " and c1.numfactu = h.numfactu "
                Cad = Cad & " and c1.fecfactu = h.fecfactu "
                Cad = Cad & " WHERE 1=1"
                Cad = Cad & " and h.codsocio=" & Data1.Recordset!Codsocio
            Else
                Cad = "select h.codtipom,h.numfactu,h.fecfactu,h.totalfac totalfac, if(sum(coalesce(c1.impvenci,0)+coalesce(c1.gastos,0)-coalesce(c1.impcobro,0))=0,'Cobrado','')  "
                Cad = Cad & " from (advfacturas h inner join usuarios.stipom u1 on h.codtipom = u1.codtipom) left join conta" & ContaADV & ".scobro c1 on c1.numserie = u1.letraser "
                Cad = Cad & " and c1.codfaccl = h.numfactu "
                Cad = Cad & " and c1.fecfaccl = h.fecfactu "
                Cad = Cad & " WHERE 1=1"
                Cad = Cad & " and h.codsocio=" & Data1.Recordset!Codsocio
            End If
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
            '[Monica] 19/04/2018: contabilidad nueva
            If vParamAplic.ContabilidadNueva Then
                Cad = Cad & "select i.codtipom,i.numfactu,i.fecfactu,i.totalfac totalfac, if(sum(coalesce(c2.impvenci,0)+coalesce(c2.gastos,0)-coalesce(c2.impcobro,0))=0,'Cobrado','')  "
                Cad = Cad & " from (rbodfacturas i inner join usuarios.stipom u2 on i.codtipom = u2.codtipom) left join ariconta" & ContaBOD & ".cobros c2 on c2.numserie = u2.letraser "
                Cad = Cad & " and c2.numfactu = i.numfactu "
                Cad = Cad & " and c2.fecfactu = i.fecfactu "
                Cad = Cad & " WHERE 1=1"
                Cad = Cad & " and i.codsocio=" & Data1.Recordset!Codsocio
            Else
                Cad = Cad & "select i.codtipom,i.numfactu,i.fecfactu,i.totalfac totalfac, if(sum(coalesce(c2.impvenci,0)+coalesce(c2.gastos,0)-coalesce(c2.impcobro,0))=0,'Cobrado','')  "
                Cad = Cad & " from (rbodfacturas i inner join usuarios.stipom u2 on i.codtipom = u2.codtipom) left join conta" & ContaBOD & ".scobro c2 on c2.numserie = u2.letraser "
                Cad = Cad & " and c2.codfaccl = i.numfactu "
                Cad = Cad & " and c2.fecfaccl = i.fecfactu "
                Cad = Cad & " WHERE 1=1"
                Cad = Cad & " and i.codsocio=" & Data1.Recordset!Codsocio
            End If
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
        
            '[Monica]19/04/2018: contabilidad nueva
            If vParamAplic.ContabilidadNueva Then
                Cad = Cad & "select j.codtipom,j.numfactu,j.fecfactu,j.totalfact totalfac, if(sum(coalesce(c.impvenci,0)+coalesce(c.gastos,0)-coalesce(c.impcobro,0))=0,'Cobrado','') "
                Cad = Cad & " from (rrecibpozos j inner join usuarios.stipom u on j.codtipom = u.codtipom)  left join ariconta" & ContaPOZOS & ".cobros c on c.numserie = u.letraser "
                Cad = Cad & " and c.numfactu = j.numfactu "
                Cad = Cad & " and c.fecfactu = j.fecfactu "
                Cad = Cad & " WHERE 1=1"
                Cad = Cad & " and j.codsocio=" & Data1.Recordset!Codsocio
            Else
                Cad = Cad & "select j.codtipom,j.numfactu,j.fecfactu,j.totalfact totalfac, if(sum(coalesce(c.impvenci,0)+coalesce(c.gastos,0)-coalesce(c.impcobro,0))=0,'Cobrado','') "
                Cad = Cad & " from (rrecibpozos j inner join usuarios.stipom u on j.codtipom = u.codtipom)  left join conta" & ContaPOZOS & ".scobro c on c.numserie = u.letraser "
                Cad = Cad & " and c.codfaccl = j.numfactu "
                Cad = Cad & " and c.fecfaccl = j.fecfactu "
                Cad = Cad & " WHERE 1=1"
                Cad = Cad & " and j.codsocio=" & Data1.Recordset!Codsocio
            End If
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
        
        
    Case 6
        'FACTURAS de cliente (advfacturas, rbodfacturas, rrecibpozos)
'[Monica]08/06/2018: lo sacamos de aportaciones
'        cad = "select h.numerodoc,h.fecmovim,sum(coalesce(h.importe,0)),h.causa  "
'        cad = cad & " from rsocios_movim h "
'        cad = cad & " WHERE 1=1"
'        cad = cad & " and h.codsocio=" & Data1.Recordset!Codsocio

        Cad = "select h.nrodocum numerodoc,h.fecaport fecmovim,sum(coalesce(h.importe,0)),h.descripcion causa  "
        Cad = Cad & " from raportacion h "
        Cad = Cad & " WHERE 1=1"
        Cad = Cad & " and h.codsocio=" & Data1.Recordset!Codsocio

        BuscaChekc = "h.fecaport"
        'La fecha
        If BuscaChekc <> "" Then Cad = Cad & " and " & BuscaChekc & " >='" & Format(imgFec(3).Tag, FormatoFecha) & "'"
        Cad = Cad & " GROUP BY 1,2,4 "
        
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
'                It.SmallIcon = ElIcono
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
Dim C As String
Dim L As Long
Dim Rs As ADODB.Recordset
Dim nFile As Long


    ProcesarCarpetaImagenes


    C = "Select * from rfichdocs where codsocio=" & DBSet(Text1(0).Text, "N") & " ORDER BY orden"
'    Me.lblCarga2.Caption = "Leyendo desde BD "
'    Me.lblCarga2.Refresh
    adodc1.ConnectionString = conn
    adodc1.RecordSource = C
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
            C = App.Path & "\" & CarpetaIMG & "\" & L
            If DBLet(adodc1.Recordset!Docum) <> "0" Then
                C = App.Path & "\" & CarpetaIMG & "\" & adodc1.Recordset!Docum
            End If
            If Dir(C) <> "" Then
                AnyadirAlListview C, True
            Else
                If LeerBinary(adodc1.Recordset!campo, C) Then
                    AnyadirAlListview C, True
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
    cadselect = ""
    cadParam = ""
    numParam = 0
End Sub

Private Sub ProcesarCarpetaImagenes()
Dim C As String
Dim MiNombre As String

    On Error GoTo EProcesarCarpetaImagenes
    C = App.Path & "\" & CarpetaIMG
    If Dir(C, vbDirectory) = "" Then
        MkDir C
    Else
        On Error Resume Next
        If Dir(C & "\*.*", vbArchive) <> "" Then 'Kill c & "\*.*"
            MiNombre = Dir(C & "\*.*")   ' Recupera la primera entrada.
            Do While MiNombre <> ""   ' Inicia el bucle.
               ' Ignora el directorio actual y el que lo abarca.
               If MiNombre <> "." And MiNombre <> ".." Then
                    Kill C & "\" & MiNombre
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
Dim SQL1 As String
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
        If PoblaAnt <> Text1(5).Text Then Cad = Cad & " Población,"
        If ProviAnt <> Text1(6).Text Then Cad = Cad & " Provincia,"
        If NifAnt <> Text1(2).Text Then Cad = Cad & " NIF,"
        If EMaiAnt <> Text1(12).Text Then Cad = Cad & " EMail,"
        
        If IbanAnt <> Text1(27).Text Then Cad = Cad & " Iban,"
        
        If BancoAnt <> Text1(16).Text Then Cad = Cad & " Banco,"
        If SucurAnt <> Text1(17).Text Then Cad = Cad & " Sucursal,"
        If DigitoAnt <> Text1(18).Text Then Cad = Cad & " Dig.Control,"
        If CuentaAnt <> Text1(19).Text Then Cad = Cad & " Cuenta banco,"
        
        Cad = Mid(Cad, 1, Len(Cad) - 1)
        
        Cad = Cad & vbCrLf & vbCrLf & "¿ Desea actualizar los datos en la Contabilidad de la Sección Horto ?" & vbCrLf & vbCrLf
        
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
'[Monica]09/06/2017: el codigo de pais no lo voy a modificar se queda el de la cuenta contable
'                            Sql = Sql & ", codpais = 'ES' "
                        
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
                            
                                SQL1 = "select sum(total) from ("
                                SQL1 = SQL1 & "select  count(*) total "
                                SQL1 = SQL1 & " from scobro  cc "
                                '[Monica]14/11/2014: si es escalona o utxera no miramos la situacion de la remesa
                                If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                                    SQL1 = SQL1 & " where cc.codmacta  in (" & Cuentas & ")  "
                                Else
                                    '[Monica]09/09/2014: la remesa no tiene que estar contabilizada. Lo cambiamos
                                    'Sql1 = Sql1 & " where cc.codmacta  in (" & Cuentas & ")  and (cc.codrem is null or cc.codrem = 0) and (cc.transfer is null or cc.transfer = 0)"
                                    SQL1 = SQL1 & " where cc.codmacta  in (" & Cuentas & ")  and (cc.siturem is null or cc.siturem <= 'Q')"
                                End If
                                SQL1 = SQL1 & " union "
                                SQL1 = SQL1 & " select count(*) total "
                                SQL1 = SQL1 & " from spagop pp "
                                '[Monica]09/09/2014: en los pagos cuando se contabiliza la transferencia desaparece el pago. Lo cambiamos
                                'Sql1 = Sql1 & " where pp.ctaprove in (" & Cuentas & ") and  (pp.transfer is null or pp.transfer = 0)"
                                SQL1 = SQL1 & " where pp.ctaprove in (" & Cuentas & ") "
                                SQL1 = SQL1 & ") aaaaaa "

                            Else
                                SQL1 = "select sum(total) from ("
                                SQL1 = SQL1 & "select  count(*) total "
                                SQL1 = SQL1 & " from cobros  cc "
                                '[Monica]14/11/2014: si es escalona o utxera no miramos la situacion de la remesa
                                If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                                    SQL1 = SQL1 & " where cc.codmacta  in (" & Cuentas & ")  "
                                Else
                                    '[Monica]09/09/2014: la remesa no tiene que estar contabilizada. Lo cambiamos
                                    'Sql1 = Sql1 & " where cc.codmacta  in (" & Cuentas & ")  and (cc.codrem is null or cc.codrem = 0) and (cc.transfer is null or cc.transfer = 0)"
                                    SQL1 = SQL1 & " where cc.codmacta  in (" & Cuentas & ")  and (cc.siturem is null or cc.siturem <= 'Q')"
                                End If
                                
                                If Not (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10) Then
                                    SQL1 = SQL1 & " and (impvenci + coalesce(gastos,0) - coalesce(impcobro, 0)) <> 0 "
                                End If
                                
                                SQL1 = SQL1 & " union "
                                SQL1 = SQL1 & " select count(*) total "
                                SQL1 = SQL1 & " from pagos pp "
                                '[Monica]09/09/2014: en los pagos cuando se contabiliza la transferencia desaparece el pago. Lo cambiamos
                                'Sql1 = Sql1 & " where pp.ctaprove in (" & Cuentas & ") and  (pp.transfer is null or pp.transfer = 0)"
                                SQL1 = SQL1 & " where pp.codmacta in (" & Cuentas & ") "
                                
                                If Not (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10) Then
                                    SQL1 = SQL1 & " and impefect - coalesce(imppagad,0) <> 0 "
                                    SQL1 = SQL1 & " and codforpa in (select codforpa from formapago where tipforpa = 1) "
                                End If
                                
                                SQL1 = SQL1 & ") aaaaaa "
                            End If

                            Nregs = 0

                            Set Rs = New ADODB.Recordset
                            Rs.Open SQL1, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
                            If Not Rs.EOF Then
                                If DBLet(Rs.Fields(0).Value) <> 0 Then Nregs = Rs.Fields(0).Value
                            End If
                            Set Rs = Nothing
                            
                            If Nregs <> 0 Then
                                Cad = "Se han producido cambios en la Cta.Bancaria del Socio."
                                Cad = Cad & vbCrLf & vbCrLf & "¿ Desea actualizar los Cobros y Pagos pendientes en Tesoreria ?" & vbCrLf & vbCrLf
                                
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
                                            Sql2 = Sql2 & " and (impvenci + coalesce(gastos,0) - coalesce(impcobro, 0)) <> 0 "
                                            
                                         End If
                                         
                                         ConnConta.Execute Sql2
                                         
                                         Sql2 = "update pagos set iban = " & DBSet(vIban, "T")
                                         Sql2 = Sql2 & " where codmacta in (" & Cuentas & ") "
                                         
                                        If Not (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10) Then
                                            Sql2 = Sql2 & " and impefect - coalesce(imppagad,0) <> 0 "
                                            Sql2 = Sql2 & " and codforpa in (select codforpa from formapago where tipforpa = 1) "
                                        End If
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
Dim cadena As String
    
    If Not PerderFocoGnral(txtAux4(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 2 ' nif
            txtAux4(Index).Text = UCase(txtAux4(Index).Text)
            ValidarNIF txtAux4(Index).Text
        
        Case 4 ' votos
            PonerFormatoDecimal txtAux4(Index), 1
        
        Case 5 ' capital
            If PonerFormatoDecimal(txtAux4(Index), 1) Then
                CmdAceptar.SetFocus
            End If

    End Select
    
    ' ******************************************************************************
End Sub

