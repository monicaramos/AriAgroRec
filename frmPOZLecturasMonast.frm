VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPOZLecturasMonast 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Toma de Lecturas"
   ClientHeight    =   12570
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   11160
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPOZLecturasMonast.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12570
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Cuenta Principal|N|N|0|1|cltebanc|ctaprpal|0||"
   Begin VB.Frame Frame3 
      Caption         =   "Lecturas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4125
      Left            =   45
      TabIndex        =   5
      Top             =   8460
      Width           =   10975
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   5175
         MaxLength       =   10
         TabIndex        =   32
         Tag             =   "Contador Actual|N|S|||rpozos|lect_act|######0||"
         Text            =   "0000000000"
         Top             =   1620
         Width           =   2085
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Guardar Lectura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   3420
         TabIndex        =   26
         Top             =   3105
         Width           =   3870
      End
      Begin VB.CommandButton CC 
         Caption         =   "<--"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   8505
         TabIndex        =   25
         Top             =   3150
         Width           =   1980
      End
      Begin VB.CommandButton C0 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   7470
         TabIndex        =   24
         Top             =   3150
         Width           =   945
      End
      Begin VB.CommandButton C3 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   9540
         TabIndex        =   23
         Top             =   2205
         Width           =   945
      End
      Begin VB.CommandButton C2 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   8505
         TabIndex        =   22
         Top             =   2205
         Width           =   945
      End
      Begin VB.CommandButton C1 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   7470
         TabIndex        =   21
         Top             =   2205
         Width           =   945
      End
      Begin VB.CommandButton C6 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   9540
         TabIndex        =   20
         Top             =   1260
         Width           =   945
      End
      Begin VB.CommandButton C5 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   8505
         TabIndex        =   19
         Top             =   1260
         Width           =   945
      End
      Begin VB.CommandButton C4 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   7470
         TabIndex        =   18
         Top             =   1260
         Width           =   945
      End
      Begin VB.CommandButton C9 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   9540
         TabIndex        =   17
         Top             =   315
         Width           =   945
      End
      Begin VB.CommandButton C8 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   8505
         TabIndex        =   16
         Top             =   315
         Width           =   945
      End
      Begin VB.CommandButton C7 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   7470
         TabIndex        =   15
         Top             =   315
         Width           =   945
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         ItemData        =   "frmPOZLecturasMonast.frx":000C
         Left            =   1845
         List            =   "frmPOZLecturasMonast.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Tag             =   "Comunidad|N|N|||rpozos|codpozo|||"
         Top             =   2385
         Width           =   5460
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Ausente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   180
         TabIndex        =   12
         Tag             =   "Cobrar Cuota|N|S|||rpozos|cobrarcuota|0|N|"
         Top             =   3285
         Width           =   2130
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   1860
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   8
         Tag             =   "Lectura Anterior|N|S|||rpozos|lect_ant|######0||"
         Text            =   "123456789"
         Top             =   720
         Width           =   1860
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   5175
         MaxLength       =   9
         TabIndex        =   7
         Tag             =   "Contador Actual|N|S|||rpozos|lect_act|######0||"
         Text            =   "123456789"
         Top             =   765
         Width           =   2085
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   1860
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   6
         Tag             =   "Consumo|N|S|||rpozos|consumo|########0||"
         Text            =   "123456789"
         Top             =   1620
         Width           =   1860
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   4905
         Picture         =   "frmPOZLecturasMonast.frx":0010
         ToolTipText     =   "Buscar fecha"
         Top             =   1755
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3915
         TabIndex        =   33
         Top             =   1650
         Width           =   1710
      End
      Begin VB.Label Label2 
         Caption         =   "Concepto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   180
         TabIndex        =   13
         Top             =   2430
         Width           =   1890
      End
      Begin VB.Label Label23 
         Caption         =   "Anterior"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   180
         TabIndex        =   11
         Top             =   735
         Width           =   1440
      End
      Begin VB.Label Label9 
         Caption         =   "Actual"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3915
         TabIndex        =   10
         Top             =   795
         Width           =   1710
      End
      Begin VB.Label Label14 
         Caption         =   "Consumo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   180
         TabIndex        =   9
         Top             =   1635
         Width           =   1620
      End
      Begin VB.Line Line1 
         X1              =   150
         X2              =   3870
         Y1              =   1380
         Y2              =   1380
      End
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
      Height          =   1845
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   90
      Width           =   11010
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         ItemData        =   "frmPOZLecturasMonast.frx":009B
         Left            =   2115
         List            =   "frmPOZLecturasMonast.frx":009D
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Tag             =   "Calle|N|N|||rpozos|codparti|||"
         Top             =   1035
         Width           =   8205
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         ItemData        =   "frmPOZLecturasMonast.frx":009F
         Left            =   2115
         List            =   "frmPOZLecturasMonast.frx":00A1
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Tag             =   "Comunidad|N|N|||rpozos|codpozo|||"
         Top             =   315
         Width           =   8205
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   330
         Left            =   10395
         TabIndex        =   29
         Top             =   360
         Width           =   450
         _ExtentX        =   794
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
      Begin VB.Label Label3 
         Caption         =   "Calle"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   180
         TabIndex        =   28
         Top             =   1035
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Comunidad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   180
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   780
      Top             =   6300
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
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
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
      Index           =   0
      Left            =   690
      MaxLength       =   40
      TabIndex        =   1
      Top             =   930
      Width           =   1035
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmPOZLecturasMonast.frx":00A3
      Height          =   6135
      Left            =   45
      TabIndex        =   4
      Top             =   2025
      Width           =   10590
      _ExtentX        =   18680
      _ExtentY        =   10821
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   29
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   15.75
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
      Index           =   0
      Left            =   225
      MaxLength       =   7
      TabIndex        =   30
      Tag             =   "Propietario|N|S|||rpozos|codcampo|0000||"
      Text            =   "1234567"
      Top             =   1530
      Width           =   1185
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
      Index           =   1
      Left            =   1530
      MaxLength       =   7
      TabIndex        =   31
      Tag             =   "Piso|T|S|||rcampos|observac|||"
      Text            =   "1234567"
      Top             =   1530
      Width           =   5325
   End
   Begin VB.CommandButton cmdActualizar 
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   10350
      TabIndex        =   35
      Top             =   1260
      Width           =   495
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
      Index           =   2
      Left            =   6930
      MaxLength       =   7
      TabIndex        =   34
      Tag             =   "Piso|T|S|||rcampos|observac|||"
      Text            =   "1234567"
      Top             =   1530
      Width           =   3255
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
      Begin VB.Menu mnBarra3 
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
Attribute VB_Name = "frmPOZLecturasMonast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MONICA                   -+-+
' +-+- Men�: Hidrantes de Pozos        -+-+
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

'Dim T1 As Single

Public DatosADevolverBusqueda As String    'Tindr� el n� de text que vol que torne, empipat
Public Event DatoSeleccionado(CadenaSeleccion As String)

' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmC1 As frmCal 'calendario fecha
Attribute frmC1.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1

Private WithEvents frmPar As frmManPartidas 'partidas
Attribute frmPar.VB_VarHelpID = -1
Private WithEvents frmPoz As frmPOZPozos 'tipos de Pozos
Attribute frmPoz.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmSoc1 As frmManSocios 'socios
Attribute frmSoc1.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'mensajes
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmCam As frmManCamposMonast 'campos
Attribute frmCam.VB_VarHelpID = -1
Private WithEvents frmMen2 As frmMensajes ' orden de printnou
Attribute frmMen2.VB_VarHelpID = -1

Private WithEvents frmHidPrev As frmPOZHidrantesMonastPrev ' campos vista previa
Attribute frmHidPrev.VB_VarHelpID = -1


' *****************************************************
Dim CodTipoMov As String

Dim Orden As String

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
Dim Indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim CadB As String

Dim vSeccion As CSeccion
Dim B As Boolean

Private BuscaChekc As String
Private NumCajas As Currency
Private NumCajasAnt As Currency
Private NumKilosAnt As Currency

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Public ImpresoraDefecto As String

Dim Lineas As Collection
Dim NF As Integer

Dim UltimaLectura As String


Private Sub C0_Click()
    Text1(kCampo).Text = Text1(kCampo).Text & "0"
End Sub

Private Sub C1_Click()
    Text1(kCampo).Text = Text1(kCampo).Text & "1"
End Sub

Private Sub C2_Click()
    Text1(kCampo).Text = Text1(kCampo).Text & "2"
End Sub

Private Sub C3_Click()
    Text1(kCampo).Text = Text1(kCampo).Text & "3"
End Sub

Private Sub C4_Click()
    Text1(kCampo).Text = Text1(kCampo).Text & "4"
End Sub

Private Sub C5_Click()
    Text1(kCampo).Text = Text1(kCampo).Text & "5"
End Sub

Private Sub C6_Click()
    Text1(kCampo).Text = Text1(kCampo).Text & "6"
End Sub

Private Sub C7_Click()
    Text1(kCampo).Text = Text1(kCampo).Text & "7"
End Sub

Private Sub C8_Click()
    Text1(kCampo).Text = Text1(kCampo).Text & "8"
End Sub

Private Sub C9_Click()
    Text1(kCampo).Text = Text1(kCampo).Text & "9"
End Sub

Private Sub CC_Click()
    If Text1(kCampo).Text <> "" Then
        Text1(kCampo).Text = Mid(Text1(kCampo).Text, 1, Len(Text1(kCampo).Text) - 1)
    End If
End Sub

Private Sub Check1_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "check1(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "check1(" & Index & ")|"
    End If
End Sub

Private Sub cmd1_Click()

End Sub

Private Sub cmdAceptar_Click()
Dim Sql As String

    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    If DatosOK Then
        Sql = "update rpozos set fech_act = " & DBSet(Text1(0).Text, "F") & ", lect_act = " & DBSet(Text1(9).Text, "N")
        Sql = Sql & ", consumo = " & DBSet(Text1(4).Text, "N")
        Sql = Sql & ", calibre = " & DBSet(Combo1(1).ListIndex, "N")
        Sql = Sql & " where hidrante = " & DBSet(DataGrid1.Columns(2), "T")
        
        conn.Execute Sql
    End If
    Screen.MousePointer = vbDefault
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Function DatosOK() As Boolean
'Dim Datos As String
Dim B As Boolean
Dim Sql As String
Dim Mens As String
Dim FechaAnt As Date
Dim NroDig As Integer
Dim Inicio As Long
Dim Fin As Long
Dim Consumo As Long
Dim Limite As Long

    If DataGrid1.Columns(2).Value = "" Then
        MsgBox "No se ha seleccionado contador. Revise.", vbExclamation
        B = False
    End If
    
    Hidrante = DataGrid1.Columns(2).Value
    
    If B Then
        If Text1(9).Text <> "" Then
             Inicio = 0
             Fin = 0
             NroDig = DevuelveValor("select digcontrol from rpozos where hidrante = " & DBSet(Hidrante, "T"))
             Limite = 10 ^ NroDig
             
             If Text1(7).Text <> "" Then Inicio = CLng(Text1(7).Text)
             If Text1(9).Text <> "" Then Fin = CLng(Text1(9).Text)
             
             
             If Fin >= Inicio Then
                Consumo = Fin - Inicio
             Else
                If MsgBox("� Es un reinicio de contador ? ", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                    Consumo = (Limite - Inicio) + Fin
                Else
                    Consumo = Fin - Inicio
                End If
             End If
             
             If Consumo > Limite - 1 Or Consumo < 0 Then
                MsgBox "Error en la lectura. Revise", vbExclamation
                PonerFoco Text1(9)
                B = False
             Else
                
                If Text1(0).Text = "" Then
                    MsgBox "La fecha de lectura debe tener un valor. Revise.", vbExclamation
                    PonerFoco Text1(0)
                    B = False
                Else
                    FechaAnt = DevuelveValor("select fech_ant from rpozos where hidrante = " & DBSet(Hidrante, "T"))
                    If CDate(Text1(9).Text) < FechaAnt Then
                        MsgBox "La fecha de lectura actual es inferior a la de �ltima lectura. Revise.", vbExclamation
                        PonerFoco Text1(9)
                        B = False
                    End If
                End If
             End If
        Else
            If Text1(9).Text = "" And Text1(0).Text = "" Then
                Text1(4).Text = ""
                B = True
            Else
                B = False
            End If
        End If
    
    End If
    
    If B Then Text1(4).Text = Consumo
    
    
    DatosOK = B
End Function




Private Sub cmdActualizar_Click()
    If Combo1(2).ListIndex <> -1 And Combo1(0).ListIndex <> -1 Then
        CargaGrid "rpozos.codparti = " & DBSet(Combo1(2).ItemData(Combo1(2).ListIndex), "N") & " and rpozos.codpozo = " & DBSet(Combo1(0).ItemData(Combo1(0).ListIndex), "N")
    Else
        CargaGrid "rpozos.codparti = -1 and rpozos.codpozo = - 1"
    End If
End Sub


Private Sub Combo1_Click(Index As Integer)
    If Index = 1 Then Exit Sub
    
    cmdActualizar_Click
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    If Me.Data1.Recordset.EOF Then Exit Sub

    Text1(7).Text = Data1.Recordset.Fields(3)
    Text1(9).Text = Data1.Recordset.Fields(4)
    Text1(4).Text = Data1.Recordset.Fields(5)
    
    If ComprobarCero(Text1(9).Text) = "0" Then Text1(9).Text = ""
    If ComprobarCero(Text1(4).Text) = "0" Then Text1(4).Text = ""
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then PrimeraVez = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim I As Integer


    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 12
    End With
    
    LimpiarCampos   'Neteja els camps TextBox
    
    '*** canviar el nom de la taula i l'ordenaci� de la cap�alera ***
    NombreTabla = "rpozos"
    Ordenacion = " ORDER BY hidrante "
    '************************************************
    
    'Mirem com est� guardat el valor del check
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    '***** canviar el nom de la PK de la cap�alera; repasar codEmpre *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where hidrante is null"
    Data1.Refresh
       
    ModoLineas = 0
         
    CargaCombo
    
    DatosaMemorizar True
    
    PonerModo 1 'b�squeda
    
    Text1(0).Text = Format(Now, "dd/mm/yyyy")

End Sub


Private Sub LimpiarCampos()
Dim I As Integer

    On Error Resume Next
    
    limpiar Me   'M�tode general: Neteja els controls TextBox
    Me.Check1(0).Value = 0
    
    ' *** si n'hi han combos a la cap�alera ***
    ' *****************************************

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub LimpiarCamposLin(FrameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, FrameAux  'M�tode general: Neteja els controls TextBox

    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO s'habiliten, o no, els diversos camps del
'   formulari en funci� del modo en que anem a treballar
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)

    
    ' el concepto
    Combo1(1).ListIndex = 1
    
    
    If UltimaLectura <> "" Then
        PosicionarCombo2 Combo1(0), Format(RecuperaValor(UltimaLectura, 1), "0000"), 4
        PosicionarCombo2 Combo1(2), Format(RecuperaValor(UltimaLectura, 2), "0000"), 4
    
        CargaGrid "rpozos.codparti = " & DBSet(Combo1(0).ItemData(Combo1(0).ListIndex), "N") & " and rpozos.codpozo = " & DBSet(Combo1(2).ItemData(Combo1(2).ListIndex), "N")
    Else
        Combo1(0).ListIndex = -1
        Combo1(2).ListIndex = -1
        
        CargaGrid "rpozos.codparti = -1 and rpozos.codpozo = -1"
    End If


EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
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



Private Sub Desplazamiento(Index As Integer)
'Botons de Despla�ament; per a despla�ar-se pels registres de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index, True
    PonerCampos
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


Private Sub frmC1_Selec(vFecha As Date)
Dim Indice As Byte
    Indice = CByte(Me.imgFec(0).Tag)
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmMen2_DatoSeleccionado(CadenaSeleccion As String)
    Orden = CadenaSeleccion
    If CadenaSeleccion = "" Then Orden = "pOrden={rpozos.hidrante}"
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo de campo
    FormateaCampo Text1(5)
    PonerDatosCampo Text1(5).Text
End Sub

Private Sub frmPar_DatoSeleccionado(CadenaSeleccion As String)
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo de partida
    FormateaCampo Text1(3)
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre de partida
End Sub

Private Sub frmPoz_DatoSeleccionado(CadenaSeleccion As String)
    Text1(13).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo de pozo
    FormateaCampo Text1(13)
    Text2(13).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre de pozo
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codsocio
    FormateaCampo Text1(2)
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2) 'nomsocio
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(Indice).Text = vCampo
End Sub


Private Sub imgZoom_Click(Index As Integer)
    Set frmZ = New frmZoom

    Select Case Index
        Case 0
            Indice = 11
            frmZ.pTitulo = "Observaciones del Contador"
            frmZ.pValor = Text1(Indice).Text
            frmZ.pModo = Modo
            frmZ.Show vbModal
            Set frmZ = Nothing
            PonerFoco Text1(Indice)
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
            Case 0
                Indice = 0
       End Select
       
       Me.imgFec(0).Tag = Indice
       
       PonerFormatoFecha Text1(Indice)
       If Text1(Indice).Text <> "" Then frmC1.NovaData = CDate(Text1(Indice).Text)
    
       Screen.MousePointer = vbDefault
       frmC1.Show vbModal
       Set frmC1 = Nothing
       PonerFoco Text1(Indice)
    

End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub


Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
Dim NroCopias As String
Dim Lin As String

    printNou
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
        Case 8 'Imprimir
'            AbrirListado (10)
            mnImprimir_Click
    End Select
End Sub

Private Sub BotonBuscar()
Dim I As Integer
' ***** Si la clau primaria de la cap�alera no es Text1(0), canviar-ho en <=== *****
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonerFoco Text1(0) ' <===
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
    
    If CadB <> "" Then
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        ' *** foco al 1r camp visible de la cap�alera que siga clau primaria ***
        PonerFoco Text1(0)
        ' **********************************************************************
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
    Dim cad As String
    Dim NombreTabla1 As String
'
'    'Cridem al form
'    ' **************** arreglar-ho per a vore lo que es desije ****************
'    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
'    cad = ""
'    cad = cad & "Contadores|rpozos.hidrante|N||18�"
'    cad = cad & "Socio|rpozos.codsocio|N|000000|12�"
'    cad = cad & "Nombre|rsocios.nomsocio|T||55�"
'    cad = cad & "Nro.Orden|rpozos.nroorden|T||15�"
'
'    NombreTabla1 = "(rpozos inner join rsocios on rpozos.codsocio = rsocios.codsocio)"
'
'    If cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vtabla = NombreTabla1
'        frmB.vSQL = CadB
'        HaDevueltoDatos = False
'        frmB.vDevuelve = "0|" '*** els camps que volen que torne ***
'        frmB.vTitulo = "Hidrantes" ' ***** repasa a��: t�tol de BuscaGrid *****
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

    Set frmHidPrev = New frmPOZHidrantesMonastPrev
    
    frmHidPrev.cWhere = CadB
    frmHidPrev.DatosADevolverBusqueda = "0|1|2|"
    frmHidPrev.Show vbModal
    
    Set frmHidPrev = Nothing


End Sub


Private Sub cmdRegresar_Click()
Dim cad As String
Dim Aux As String
Dim I As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ning�n registro devuelto.", vbExclamation
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
    
    ' *** canviar o llevar, si cal, el WHERE; repasar codEmpre ***
    CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
    'CadenaConsulta = "Select * from " & NombreTabla & " where codempre = " & codEmpre & Ordenacion
    ' ******************************************
    PonerCadenaBusqueda
    ' *** si n'hi han ll�nies sense grids ***
'        CargaFrame 0, True
    ' ************************************
End Sub


Private Sub BotonAnyadir()

    LimpiarCampos 'Huida els TextBox
    PonerModo 3
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la cap�alera *******
'    Text1(0).Text = SugerirCodigoSiguienteStr("rentradas", "numnotac")
'    FormateaCampo Text1(0)
       
    PonerFoco Text1(0) '*** 1r camp visible que siga PK ***
    ' ***********************************************************
    ' *** si n'hi han camps de descripci� a la cap�alera ***
    PosarDescripcions
    ' ******************************************************

End Sub


Private Sub BotonModificar()

    PonerModo 4

    ' *** bloquejar els camps visibles de la clau primaria de la cap�alera ***
    BloquearTxt Text1(0), True
    ' *************************************************************************
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(1)
    ' *********************************************************
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
    cad = "�Seguro que desea eliminar el Hidrante?"
    cad = cad & vbCrLf & "Hidrante: " & Data1.Recordset.Fields(0)
    ' **************************************************************************
    
    'borrem
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
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
Dim I As Integer
Dim CodPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1  'opcio=1: posa el format o els camps de la cap�alera
    
    ' *** si n'hi han ll�nies en datagrids ***
    
    PosarDescripcions


' lo he quitado de aqui pq el consumo est� almacenado en un campo de la tabla rpozos
'    CalcularConsumo
    ' ********************************************************************************
    
    
    
End Sub

Private Sub CalcularConsumo()
Dim Sql As String
Dim Inicio As Long
Dim Fin As Long
Dim Consumo As Long
Dim NroDig As Integer
Dim Limite As Long

    If Text1(9).Text = "" Then
        Text1(4).Text = "0"
        Exit Sub
    End If

    Inicio = 0
    Fin = 0
    
    If Text1(7).Text <> "" Then Inicio = CLng(Text1(7).Text)
    If Text1(9).Text <> "" Then Fin = CLng(Text1(9).Text)
    
    NroDig = CCur(Text1(12).Text)
    Limite = (10 ^ NroDig)
    
    If Fin >= Inicio Then
        Consumo = Fin - Inicio
    Else
        If MsgBox("� Es un reinicio de contador ? ", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            Consumo = (Limite - Inicio) + Fin
        Else
            Consumo = Fin - Inicio
        End If
    End If
    
    If Consumo > (Limite - 1) Or Consumo < 0 Then
        MsgBox "Error en la lectura.", vbExclamation
        PonerFoco Text1(9)
    End If
    
   
'    Text2(0).Text = Format(Consumo, "#,###,##0")
    Text1(4).Text = Format(Consumo, "#,###,##0")

End Sub


Private Sub cmdCancelar_Click()
Dim I As Integer
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
          
          
    End Select
End Sub



Private Sub PosicionarData()
Dim cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la cap�alera, no llevar els () ***
    cad = "(hidrante='" & Text1(0).Text & "')"
    ' ***************************************
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarData(Data1, cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
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
    vWhere = " WHERE hidrante='" & Trim(Data1.Recordset!Hidrante) & "'"
        ' ***********************************************************************
        
    conn.Execute "Delete from rpozos_cooprop " & vWhere
    conn.Execute "Delete from rpozos_campos " & vWhere
        
    conn.Execute "Delete from " & NombreTabla & vWhere
       
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

    Select Case Index
        Case 0 ' fecha
            PonerFormatoFecha Text1(0)
    
        Case 9 ' CONTADORES
            PonerFormatoEntero Text1(Index)
            CalcularConsumo
                        
    End Select
End Sub

Private Sub PonerDatosCampo(campo As String)
Dim cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim Rs As ADODB.Recordset
Dim Sql As String


    If campo = "" Then Exit Sub
    
    If Not (Modo = 3 Or Modo = 4) Then Exit Sub

    '[Monica]22/11/2012: Preguntamos si quiere traer los datos del socio del campo
    If (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10) And Modo = 4 Then
        Sql = "select rcampos.codsocio, rsocios.nomsocio from rcampos inner join rsocios on rcampos.codsocio = rsocios.codsocio where rcampos.codcampo = " & DBSet(Text1(5).Text, "N")
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
       
        If DBLet(Rs.Fields(0)) <> CLng(ComprobarCero(Text1(2).Text)) Then
            Text1(2).Text = Format(DBLet(Rs!Codsocio, "N"), "000000") ' codigo de socio del campo
            Text2(2).Text = DBLet(Rs!nomsocio, "T") ' nombre de socio
           
           'If MsgBox("� Desea traer los datos de RAE al contador ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
        End If
        
        Set Rs = Nothing
        
        Exit Sub
        
    End If

    cad = "rcampos.codcampo = " & DBSet(campo, "N") & " and rcampos.fecbajas is null"
     
    Cad1 = "select rcampos.codparti, rpartida.nomparti, rcampos.poligono, rcampos.parcela, rcampos.supcoope, rpueblos.despobla, rcampos.subparce, rcampos.codsocio "
    Cad1 = Cad1 & " from rcampos, rpartida, rpueblos "
    Cad1 = Cad1 & " where " & cad
    Cad1 = Cad1 & " and rcampos.codparti = rpartida.codparti "
    Cad1 = Cad1 & " and rpartida.codpobla = rpueblos.codpobla"
     
    Set Rs = New ADODB.Recordset
    Rs.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Text1(3).Text = ""
        Text2(1).Text = ""
        
        Text1(5).Text = campo
        PonerFormatoEntero Text1(5)
        Text1(3).Text = DBLet(Rs.Fields(0).Value, "N") ' codigo de partida
        If Text1(3).Text <> "" Then Text1(3).Text = Format(Text1(3).Text, "0000")
        Text2(3).Text = DBLet(Rs.Fields(1).Value, "T") ' nombre de partida
        Text2(1).Text = DBLet(Rs.Fields(5).Value, "T") ' nombre de poblacion
'[Monica]03/08/2012: quito el formato de poligono y parcela
'        If Text1(4).Text <> "" Then Text1(4).Text = Format(Text1(4).Text, "0000")
        
        If vParamAplic.Cooperativa = 10 Then Text1(5).Text = Text1(5).Text & " " & DBLet(Rs.Fields(6).Value)
        
'        If Text1(5).Text <> "" Then Text1(5).Text = Format(Text1(5).Text, "000000")
        
    End If
    
    Set Rs = Nothing
    
End Sub




Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 2: KEYBusqueda KeyAscii, 0 'socio
                Case 3: KEYBusqueda KeyAscii, 1 'partida
            End Select
        End If
    Else
        KEYpress KeyAscii
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


Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
End Sub

' **** si n'hi han camps de descripci� a la cap�alera ****
Private Sub PosarDescripcions()
Dim NomEmple As String
Dim CodPobla As String
Dim Sql As String

    On Error GoTo EPosarDescripcions

    Text2(2).Text = PonerNombreDeCod(Text1(2), "rsocios", "nomsocio", "codsocio", "N")
    Text2(3).Text = PonerNombreDeCod(Text1(3), "rpartida", "nomparti", "codparti", "N")
    Text2(13).Text = PonerNombreDeCod(Text1(13), "rtipopozos", "nompozo", "codpozo", "N")
        
        
    If Text1(3).Text <> "" Then
        Sql = "select despobla from rpueblos, rpartida where rpartida.codparti = " & DBSet(Text1(3).Text, "N")
        Sql = Sql & " and rpueblos.codpobla = rpartida.codpobla "
        
        Text2(1).Text = DevuelveValor(Sql) ' nombre de poblacion
    End If
    
EPosarDescripcions:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo descripciones", Err.Description
End Sub
' ************************************************************



' *** si n'hi han formularis de buscar codi a les ll�nies ***
Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
    
    Select Case Index
       Case 0 'Socios
            Set frmSoc = New frmManSocios
'            frmSoc.DeConsulta = True
            frmSoc.DatosADevolverBusqueda = "0|1|"
'            frmSoc.CodigoActual = Text1(1).Text
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco Text1(2)
    
       Case 1 'Partidas
            Set frmPar = New frmManPartidas
            frmPar.DeConsulta = True
            frmPar.DatosADevolverBusqueda = "0|1|"
            frmPar.CodigoActual = Text1(3).Text
            frmPar.Show vbModal
            Set frmPar = Nothing
            PonerFoco Text1(3)
    
       Case 2 'Tipo de Pozos
            Set frmPoz = New frmPOZPozos
            frmPoz.DeConsulta = True
            frmPoz.DatosADevolverBusqueda = "0|1|"
            frmPoz.CodigoActual = Text1(3).Text
            frmPoz.Show vbModal
            Set frmPoz = Nothing
            PonerFoco Text1(13)
    
       Case 3 'Campo
            Set frmCam = New frmManCamposMonast
            frmCam.DatosADevolverBusqueda = "0|1|"
'            frmCam.CodigoActual = Text1(5).Text
            frmCam.Show vbModal
            Set frmCam = Nothing
            PonerFoco Text1(5)
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub



Private Sub printNou()
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
    
    
    ' pedimos el orden del informe
    Set frmMen2 = New frmMensajes
    
    frmMen2.OpcionMensaje = 38
    frmMen2.Show vbModal
    
    Set frmMen2 = Nothing
    
    indRPT = 78 ' personalizacion del informe de hidrantes
    
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
    
    '[Monica]12/03/2013: solo si es quatretonda tiene una impresion expandida de rpozos_campos
    If vParamAplic.Cooperativa = 7 Then
        If MsgBox("� Desea imprimir en formato expandido ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            nomDocu = Replace(nomDocu, ".rpt", "1.rpt")
        End If
    End If
    
    
    With frmImprimir2
        .cadTabla2 = "rpozos"
        .Informe2 = nomDocu
        If CadB <> "" Then
            .cadRegSelec = SQL2SF(CadB)
        Else
            .cadRegSelec = ""
        End If
        .cadRegActua = POS2SF(Data1, Me)
        '[Monica]13/07/2012: falla si hay un solo registro seleccionado y apretamos registros buscados
        If Data1.Recordset.RecordCount = 1 Then .cadRegSelec = .cadRegActua
        .cadTodosReg = ""
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|" & Orden
        .NumeroParametros2 = 2
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = True
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

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub



Private Sub CargaCombo()
Dim miRsAux As ADODB.Recordset

    Combo1(0).Clear
    Combo1(1).Clear
    
    'Comunidades
    Set miRsAux = New ADODB.Recordset

    miRsAux.Open "Select * from rtipopozos order by nompozo", conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not miRsAux.EOF
        Combo1(0).AddItem Format(miRsAux!codpozo, "0000") & " " & miRsAux!nompozo
        Combo1(0).ItemData(Combo1(0).NewIndex) = miRsAux!codpozo
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    'Calles
    Set miRsAux = New ADODB.Recordset

    miRsAux.Open "Select * from rpartida order by nomparti", conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not miRsAux.EOF
        Combo1(2).AddItem Format(miRsAux!codparti, "0000") & " " & miRsAux!nomparti
        Combo1(2).ItemData(Combo1(2).NewIndex) = miRsAux!codparti
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    'Conceptos
    Set miRsAux = New ADODB.Recordset

    miRsAux.Open "Select * from rriego order by nomriego", conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not miRsAux.EOF
        Combo1(1).AddItem Format(miRsAux!codriego, "00") & " " & miRsAux!nomriego
        Combo1(1).ItemData(Combo1(1).NewIndex) = miRsAux!codriego
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    
End Sub


Private Sub CargaGrid(Optional vSQL As String)
    Dim Sql As String
    Dim tots As String


    CadenaConsulta = "select rpozos.codcampo, rcampos.observac, rpozos.hidrante, rpozos.lect_ant, rpozos.lect_act,  "
    CadenaConsulta = CadenaConsulta & " if(rpozos.lect_act is null or rpozos.lect_act = 0, 0,rpozos.lect_act - rpozos.lect_ant) consumo "
    CadenaConsulta = CadenaConsulta & " from rpozos inner join rcampos on rpozos.codcampo = rcampos.codcampo "
    
    If vSQL <> "" Then
        CadenaConsulta = CadenaConsulta & " where " & vSQL
    End If

    Sql = CadenaConsulta

    '********************* canviar el ORDER BY *********************++
    Sql = Sql & " ORDER BY rcampos.observac"
    '**************************************************************++

    CargaGridGnral Me.DataGrid1, Me.Data1, Sql, PrimeraVez

    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "S|txtAux(0)|T|Propiedad|1800|;"
    tots = tots & "S|txtAux(1)|T|Piso|6400|;N||||0|;N||||0|;N||||0|;S|txtAux(2)|T|Consumo|1800|;"

    arregla tots, DataGrid1, Me, 510 '350

    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgCenter
    
End Sub

Private Sub DatosaMemorizar(Leer As Boolean)
Dim NF As Integer
Dim cad As String
On Error GoTo ENumeroEmpresaMemorizar


        
    cad = App.Path & "\ultLect.dat"
    If Leer Then
        If Dir(cad) <> "" Then
            NF = FreeFile
            Open cad For Input As #NF
            Line Input #NF, cad
            Close #NF
            cad = Trim(cad)
            
                'El primer pipe es el usuario
                UltimaLectura = cad
    
        End If
    Else 'Escribir
        NF = FreeFile
        Open cad For Output As #NF
        cad = Text1(0).Text
        Print #NF, UltimaLectura
        Close #NF
    End If
ENumeroEmpresaMemorizar:
    Err.Clear
End Sub

