VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEntBasculaNat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entrada en báscula"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13095
   Icon            =   "frmEntBasculaNat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   13095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Cuenta Principal|N|N|0|1|cltebanc|ctaprpal|0||"
   Begin VB.Frame Frame2 
      Height          =   765
      Index           =   0
      Left            =   240
      TabIndex        =   27
      Top             =   480
      Width           =   12465
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   22
         Left            =   6660
         MaxLength       =   10
         TabIndex        =   2
         Top             =   360
         Width           =   1200
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   6660
         MaxLength       =   20
         TabIndex        =   49
         Tag             =   "Hora|FH|N|||rentradas|horaentr|yyyy-mm-dd hh:mm:ss||"
         Top             =   360
         Width           =   1200
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   4380
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Entrada|F|N|||rentradas|fechaent|dd/mm/yyyy||"
         Top             =   360
         Width           =   1200
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1320
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Numero de Nota|N|S|1|9999999|rentradas|numnotac|0000000|S|"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label26 
         Caption         =   "Hora"
         Height          =   255
         Left            =   6075
         TabIndex        =   50
         Top             =   360
         Width           =   570
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   4095
         Picture         =   "frmEntBasculaNat.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label18 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   3465
         TabIndex        =   48
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Nota"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   25
      Top             =   6225
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
         TabIndex        =   26
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   11640
      TabIndex        =   22
      Top             =   6330
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   10380
      TabIndex        =   21
      Top             =   6360
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4185
      Top             =   6315
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
      TabIndex        =   30
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
         NumButtons      =   20
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
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Tara Tractor"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Paletización"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
         TabIndex        =   31
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   11640
      TabIndex        =   29
      Top             =   6330
      Visible         =   0   'False
      Width           =   1035
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4830
      Left            =   240
      TabIndex        =   33
      Top             =   1305
      Width           =   12465
      _ExtentX        =   21987
      _ExtentY        =   8520
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   6
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
      TabCaption(0)   =   "Datos entrada"
      TabPicture(0)   =   "frmEntBasculaNat.frx":0097
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(26)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label28"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "imgBuscar(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "imgBuscar(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label11"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label23"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "imgBuscar(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label9"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "imgBuscar(3)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label12"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "imgBuscar(4)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "imgBuscar(5)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label20"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label22"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "imgAyuda(0)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "FrameDatosDtoAdministracion"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text5(3)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text4(3)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text3(3)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text1(2)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text2(2)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Combo1(0)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text1(1)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text2(1)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text1(5)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text2(3)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text2(4)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Combo1(1)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text2(12)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text1(12)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text2(6)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text1(6)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text2(7)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text1(7)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Combo1(2)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Text2(0)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).ControlCount=   41
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   4710
         MaxLength       =   4
         TabIndex        =   90
         Top             =   1260
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Tag             =   "Transportado por|N|N|0|1|rentradas|transportadopor||N|"
         Top             =   4110
         Width           =   1560
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1380
         MaxLength       =   2
         TabIndex        =   8
         Tag             =   "Código Tarifa|N|S|0|99|rentradas|codtarif|00||"
         Top             =   3480
         Width           =   555
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   1995
         TabIndex        =   59
         Top             =   3480
         Width           =   3540
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1380
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Código Transporte|T|S|||rentradas|codtrans|||"
         Top             =   3120
         Width           =   1005
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   2430
         TabIndex        =   57
         Top             =   3120
         Width           =   3090
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   1380
         MaxLength       =   4
         TabIndex        =   6
         Tag             =   "Código Capataz|N|S|0|9999|rentradas|codcapat|0000||"
         Top             =   2760
         Width           =   555
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   1995
         TabIndex        =   55
         Top             =   2760
         Width           =   3540
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Tag             =   "Recolectado|N|N|0|1|rentradas|recolect||N|"
         Top             =   4110
         Width           =   1560
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   1380
         MaxLength       =   4
         TabIndex        =   53
         Top             =   1648
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2340
         MaxLength       =   30
         TabIndex        =   52
         Top             =   1648
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1380
         MaxLength       =   8
         TabIndex        =   5
         Tag             =   "Código Campo|N|N|1|99999999|rentradas|codcampo|00000000|N|"
         Top             =   1260
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2280
         MaxLength       =   40
         TabIndex        =   46
         Top             =   896
         Width           =   3315
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1380
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "Código Socio|N|N|1|999999|rentradas|codsocio|000000|N|"
         Top             =   896
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Tag             =   "Tipo Entrada|N|N|0|3|rentradas|tipoentr||N|"
         Top             =   4110
         Width           =   1560
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2280
         MaxLength       =   30
         TabIndex        =   42
         Top             =   520
         Width           =   3315
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1410
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Variedad|N|N|1|9999|rentradas|codvarie|0000||"
         Top             =   520
         Width           =   840
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1380
         MaxLength       =   6
         TabIndex        =   37
         Top             =   2385
         Width           =   855
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2325
         MaxLength       =   30
         TabIndex        =   36
         Top             =   2385
         Width           =   3255
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1380
         MaxLength       =   30
         TabIndex        =   35
         Top             =   2040
         Width           =   4200
      End
      Begin VB.Frame FrameDatosDtoAdministracion 
         Caption         =   "Pesos y Taras"
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
         Height          =   4065
         Left            =   5895
         TabIndex        =   34
         Top             =   405
         Width           =   6420
         Begin VB.ComboBox Combo15 
            Height          =   315
            Index           =   4
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   96
            Tag             =   "Tipo Cajon 5|N|S|||rentradas|tipocajo5||N|"
            Top             =   2490
            Width           =   1860
         End
         Begin VB.ComboBox Combo15 
            Height          =   315
            Index           =   0
            Left            =   270
            Style           =   2  'Dropdown List
            TabIndex        =   92
            Tag             =   "Tipo Cajon 1|N|S|||rentradas|tipocajo1||N|"
            Top             =   1080
            Width           =   1860
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   24
            Left            =   4995
            MaxLength       =   7
            TabIndex        =   20
            Tag             =   "Otras Taras|N|S|0|999999|rentradas|otrastaras|###,##0||"
            Top             =   3360
            Width           =   1155
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   23
            Left            =   2160
            MaxLength       =   7
            TabIndex        =   23
            Tag             =   "Peso Trasnportista|N|S|0|999999|rentradas|kilostra|###,##0||"
            Top             =   3660
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   2160
            MaxLength       =   5
            TabIndex        =   18
            Top             =   2970
            Width           =   1215
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Height          =   2310
            Left            =   135
            TabIndex        =   62
            Top             =   585
            Width           =   6135
            Begin VB.ComboBox Combo15 
               Height          =   315
               Index           =   3
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   95
               Tag             =   "Tipo Cajon 4|N|S|||rentradas|tipocajo4||N|"
               Top             =   1560
               Width           =   1860
            End
            Begin VB.ComboBox Combo15 
               Height          =   315
               Index           =   2
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   94
               Tag             =   "Tipo Cajon 3|N|S|||rentradas|tipocajo3||N|"
               Top             =   1200
               Width           =   1860
            End
            Begin VB.ComboBox Combo15 
               Height          =   315
               Index           =   1
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   93
               Tag             =   "Tipo Cajon 2|N|S|||rentradas|tipocajo2||N|"
               Top             =   840
               Width           =   1860
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000018&
               Height          =   285
               Index           =   13
               Left            =   2025
               MaxLength       =   5
               TabIndex        =   13
               Tag             =   "Nro.Cajas 1|N|S|||rentradas|numcajo1|#,##0||"
               Top             =   480
               Width           =   1215
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000018&
               Height          =   285
               Index           =   18
               Left            =   4845
               MaxLength       =   7
               TabIndex        =   67
               Tag             =   "Tara 1|N|S|0|999999|rentradas|taracaja1|###,##0||"
               Top             =   480
               Width           =   1155
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000018&
               Height          =   285
               Index           =   19
               Left            =   4845
               MaxLength       =   7
               TabIndex        =   66
               Tag             =   "Tara 2|N|S|0|999999|rentradas|taracaja2|###,##0||"
               Top             =   840
               Width           =   1155
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000018&
               Height          =   285
               Index           =   20
               Left            =   4845
               MaxLength       =   7
               TabIndex        =   65
               Tag             =   "Tara 3|N|S|0|999999|rentradas|taracaja3|###,##0||"
               Top             =   1200
               Width           =   1155
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000018&
               Height          =   285
               Index           =   8
               Left            =   4845
               MaxLength       =   7
               TabIndex        =   64
               Tag             =   "Tara 4|N|S|0|999999|rentradas|taracaja4|###,##0||"
               Top             =   1560
               Width           =   1155
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000018&
               Height          =   285
               Index           =   9
               Left            =   4845
               MaxLength       =   7
               TabIndex        =   63
               Tag             =   "Tara 5|N|S|0|999999|rentradas|taracaja5|###,##0||"
               Top             =   1920
               Width           =   1155
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000018&
               Height          =   285
               Index           =   14
               Left            =   2025
               MaxLength       =   5
               TabIndex        =   14
               Tag             =   "Nro.Cajas 2|N|S|||rentradas|numcajo2|#,##0||"
               Top             =   840
               Width           =   1215
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000018&
               Height          =   285
               Index           =   15
               Left            =   2025
               MaxLength       =   5
               TabIndex        =   15
               Tag             =   "Nro.Cajas 3|N|S|||rentradas|numcajo3|#,##0||"
               Top             =   1200
               Width           =   1215
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000018&
               Height          =   285
               Index           =   16
               Left            =   2025
               MaxLength       =   5
               TabIndex        =   16
               Tag             =   "Nro.Cajas 4|N|S|||rentradas|numcajo4|#,##0||"
               Top             =   1560
               Width           =   1215
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000018&
               Height          =   285
               Index           =   17
               Left            =   2025
               MaxLength       =   5
               TabIndex        =   17
               Tag             =   "Nro.Cajas 5|N|S|||rentradas|numcajo5|#,##0||"
               Top             =   1920
               Width           =   1215
            End
            Begin VB.Line Line3 
               X1              =   135
               X2              =   5985
               Y1              =   2295
               Y2              =   2295
            End
            Begin VB.Line Line2 
               X1              =   135
               X2              =   6030
               Y1              =   405
               Y2              =   405
            End
            Begin VB.Label Label16 
               Caption         =   "Tara"
               Height          =   255
               Left            =   4860
               TabIndex        =   85
               Top             =   180
               Width           =   1185
            End
            Begin VB.Label Label14 
               Caption         =   "Peso Caja"
               Height          =   255
               Left            =   3600
               TabIndex        =   84
               Top             =   180
               Width           =   1185
            End
            Begin VB.Label Label19 
               Caption         =   "x  Peso 1  "
               Height          =   255
               Index           =   0
               Left            =   3465
               TabIndex        =   83
               Top             =   510
               Width           =   840
            End
            Begin VB.Label Label15 
               Caption         =   "Tarifa"
               Height          =   255
               Index           =   0
               Left            =   135
               TabIndex        =   82
               Top             =   510
               Width           =   1830
            End
            Begin VB.Label Label15 
               Caption         =   "Tarifa"
               Height          =   255
               Index           =   1
               Left            =   135
               TabIndex        =   81
               Top             =   870
               Width           =   1830
            End
            Begin VB.Label Label15 
               Caption         =   "Tarifa"
               Height          =   255
               Index           =   2
               Left            =   135
               TabIndex        =   80
               Top             =   1230
               Width           =   1830
            End
            Begin VB.Label Label15 
               Caption         =   "Tarifa"
               Height          =   255
               Index           =   3
               Left            =   135
               TabIndex        =   79
               Top             =   1590
               Width           =   1830
            End
            Begin VB.Label Label15 
               Caption         =   "Tarifa"
               Height          =   255
               Index           =   4
               Left            =   135
               TabIndex        =   78
               Top             =   1950
               Width           =   1830
            End
            Begin VB.Label Label13 
               Caption         =   "Cajas"
               Height          =   255
               Left            =   2025
               TabIndex        =   77
               Top             =   180
               Width           =   1185
            End
            Begin VB.Label Label19 
               Caption         =   "x  Peso 1"
               Height          =   255
               Index           =   1
               Left            =   3465
               TabIndex        =   76
               Top             =   870
               Width           =   705
            End
            Begin VB.Label Label19 
               Caption         =   "x  Peso 1"
               Height          =   255
               Index           =   2
               Left            =   3465
               TabIndex        =   75
               Top             =   1230
               Width           =   705
            End
            Begin VB.Label Label19 
               Caption         =   "x  Peso 1"
               Height          =   255
               Index           =   3
               Left            =   3465
               TabIndex        =   74
               Top             =   1590
               Width           =   705
            End
            Begin VB.Label Label19 
               Caption         =   "x  Peso 1"
               Height          =   255
               Index           =   4
               Left            =   3465
               TabIndex        =   73
               Top             =   1950
               Width           =   705
            End
            Begin VB.Label Label10 
               Caption         =   "= "
               Height          =   255
               Index           =   0
               Left            =   4635
               TabIndex        =   72
               Top             =   510
               Width           =   150
            End
            Begin VB.Label Label10 
               Caption         =   "= "
               Height          =   255
               Index           =   1
               Left            =   4635
               TabIndex        =   71
               Top             =   870
               Width           =   150
            End
            Begin VB.Label Label10 
               Caption         =   "= "
               Height          =   255
               Index           =   2
               Left            =   4635
               TabIndex        =   70
               Top             =   1230
               Width           =   150
            End
            Begin VB.Label Label10 
               Caption         =   "= "
               Height          =   255
               Index           =   3
               Left            =   4635
               TabIndex        =   69
               Top             =   1590
               Width           =   150
            End
            Begin VB.Label Label10 
               Caption         =   "= "
               Height          =   255
               Index           =   4
               Left            =   4635
               TabIndex        =   68
               Top             =   1950
               Width           =   150
            End
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   3
            Left            =   4995
            MaxLength       =   7
            TabIndex        =   19
            Tag             =   "Tara Vehiculo|N|S|0|999999|rentradas|taravehi|###,##0||"
            Top             =   3000
            Width           =   1155
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
            Height          =   285
            Index           =   11
            Left            =   4995
            MaxLength       =   7
            TabIndex        =   24
            Tag             =   "Peso Neto|N|N|0|999999|rentradas|kilosnet|###,##0||"
            Top             =   3690
            Width           =   1155
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
            Height          =   285
            Index           =   21
            Left            =   4995
            MaxLength       =   7
            TabIndex        =   12
            Tag             =   "Peso Bruto|N|N|||rentradas|kilosbru|###,##0||"
            Top             =   270
            Width           =   1155
         End
         Begin VB.Label Label21 
            Caption         =   "Otras Taras"
            Height          =   255
            Left            =   3780
            TabIndex        =   89
            Top             =   3405
            Width           =   1185
         End
         Begin VB.Label Label15 
            Caption         =   "Neto Transportista"
            Height          =   255
            Index           =   6
            Left            =   270
            TabIndex        =   88
            Top             =   3690
            Width           =   1830
         End
         Begin VB.Label Label15 
            Caption         =   "Bonificación"
            Height          =   255
            Index           =   5
            Left            =   270
            TabIndex        =   87
            Top             =   3000
            Width           =   1830
         End
         Begin VB.Label Label8 
            Caption         =   "Tara Vehiculo"
            Height          =   255
            Left            =   3780
            TabIndex        =   61
            Top             =   3045
            Width           =   1185
         End
         Begin VB.Label Label17 
            Caption         =   "Peso Bruto"
            Height          =   255
            Left            =   3735
            TabIndex        =   45
            Top             =   315
            Width           =   1185
         End
         Begin VB.Label Label7 
            Caption         =   "Peso Neto"
            Height          =   255
            Left            =   3780
            TabIndex        =   44
            Top             =   3720
            Width           =   1185
         End
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   0
         Left            =   1530
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   3840
         Width           =   240
      End
      Begin VB.Label Label22 
         Caption         =   "Nº Orden"
         Height          =   255
         Left            =   3870
         TabIndex        =   91
         Top             =   1305
         Width           =   810
      End
      Begin VB.Label Label20 
         Caption         =   "Transportado por"
         Height          =   255
         Left            =   3960
         TabIndex        =   86
         Top             =   3870
         Width           =   1395
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1080
         ToolTipText     =   "Buscar Campo"
         Top             =   1305
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1080
         ToolTipText     =   "Buscar Tarifa"
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label12 
         Caption         =   "Tarifa"
         Height          =   255
         Left            =   240
         TabIndex        =   60
         Top             =   3480
         Width           =   705
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1080
         ToolTipText     =   "Buscar Transportista"
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Transp."
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   3120
         Width           =   705
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1080
         ToolTipText     =   "Buscar Capataz"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label23 
         Caption         =   "Capataz"
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   2760
         Width           =   705
      End
      Begin VB.Label Label11 
         Caption         =   "Recolectado"
         Height          =   255
         Left            =   2100
         TabIndex        =   54
         Top             =   3870
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Nº Campo"
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   1305
         Width           =   810
      End
      Begin VB.Label Label4 
         Caption         =   "Socio"
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   900
         Width           =   690
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1080
         ToolTipText     =   "Buscar Socio"
         Top             =   900
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Entrada"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   3840
         Width           =   1035
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1080
         ToolTipText     =   "Buscar Variedad"
         Top             =   540
         Width           =   240
      End
      Begin VB.Label Label5 
         Caption         =   "Variedad"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Partida"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   40
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label28 
         Caption         =   "Poblacion"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Zona"
         Height          =   255
         Index           =   26
         Left            =   240
         TabIndex        =   38
         Top             =   2400
         Width           =   735
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   25
      Left            =   9600
      TabIndex        =   32
      Top             =   765
      Width           =   1425
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
      Begin VB.Menu mnTararTractor 
         Caption         =   "&Tarar Tractor"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnPaletizacion 
         Caption         =   "Paletización"
         Shortcut        =   ^P
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
Attribute VB_Name = "frmEntBasculaNat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MONICA                   -+-+
' +-+- Menú: Entrada de Bascula        -+-+
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

' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmC1 As frmCal 'calendario fecha
Attribute frmC1.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1

Private WithEvents frmVar As frmComVar 'variedades
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmCap As frmManCapataz 'capataces
Attribute frmCap.VB_VarHelpID = -1
Private WithEvents frmTrans As frmManTranspor 'transportista
Attribute frmTrans.VB_VarHelpID = -1
Private WithEvents frmTar As frmManTarTra 'tarifa de transportista
Attribute frmTar.VB_VarHelpID = -1
Private WithEvents frmCamp As frmManCampos 'campos
Attribute frmCamp.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'mensajes
Attribute frmMens.VB_VarHelpID = -1

' *****************************************************
Dim CodTipoMov As String
Dim v_cadena As String

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
Private NumCajas As Currency
Private NumCajasAnt As Currency
Private NumKilosAnt As Currency

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Public ImpresoraDefecto As String

Dim Lineas As Collection
Dim NF As Integer


Private Sub cmdAceptar_Click()
Dim Mens As String

    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BÚSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                InsertarCabecera
            
'                If InsertarDesdeForm2(Me, 1) Then
                    ' *** canviar o llevar el WHERE, repasar codEmpre ****
                    Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
                    'Data1.RecordSource = "Select * from " & NombreTabla & " where codempre = " & codEmpre & Ordenacion
                    ' ***************************************************************
                    PosicionarData
                    mnPaletizacion_Click
                    
                    TerminaBloquear
'                End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                Text1(4).Text = Text1(10).Text & " " & Format(Text1(22).Text, "hh:mm:ss")
                If ModificaDesdeFormulario2(Me, 1) Then
                    Mens = ""
                    If Not ActualizarChivato(Mens, "U") Then
                        MsgBox "Error actualizando chivato: " & vbCrLf & Mens, vbExclamation
                    End If
                    
                    '[Monica]08/02/2012: Si han modificado variedad socio o campo actualizamos en traza
                    If Data1.Recordset!codvarie <> CLng(Text1(2).Text) Or Data1.Recordset!Codsocio <> CLng(Text1(1).Text) Or Data1.Recordset!CodCampo <> CLng(Text1(5).Text) Or _
                       Data1.Recordset!FechaEnt <> Text1(10).Text Or Data1.Recordset!horaentr <> Text1(4).Text Then
                         Mens = "No se han realizado los cambios en Trazabilidad. " & vbCrLf
                         If Not ActualizarTraza(Text1(0).Text, Text1(2).Text, Text1(1).Text, Text1(5).Text, Text1(10).Text, Text1(4).Text, Mens) Then
                            MsgBox Mens, vbExclamation
                         End If
                    End If
                    
                    TerminaBloquear
                    PosicionarData
                    
                    If HanModificadoCajas Then
                        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
                            CrearPaletizacion
                            TerminaBloquear
                        End If
                    Else
                        If HanModificadoKilos Then
                            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
                                ActualizarPaletizacion
                                TerminaBloquear
                            End If
                        End If
                    End If
                    
                End If
            Else
                ModoLineas = 0
            End If
        
        Case 5 'modifico la tara del tractor
            If DatosOk Then
                Text1(4).Text = Text1(10).Text & " " & Format(Text1(22).Text, "hh:mm:ss")
                If ModificaDesdeFormulario2(Me, 1) Then
                    TerminaBloquear
                    PosicionarData
                    
                    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
                        ActualizarPaletizacion
                        TerminaBloquear
                    End If
                End If
            Else
                ModoLineas = 0
            End If
    End Select
    Screen.MousePointer = vbDefault
    
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub cmdAux_Click(Index As Integer)
'    TerminaBloquear
'    Select Case Index
'        Case 4 'Secciones
'            Set frmSec = New frmManSeccion
'            frmSec.DatosADevolverBusqueda = "0|1|"
'            frmSec.CodigoActual = Text1(1).Text
'            frmSec.Show vbModal
'            Set frmSec = Nothing
'            PonerFoco Text1(1)
'
'        Case 0, 1 'fecha de alta y fecha de baja
'           If ModoLineas = 0 Then Exit Sub
'           Screen.MousePointer = vbHourglass
'
'           Dim esq As Long
'           Dim dalt As Long
'           Dim menu As Long
'           Dim obj As Object
'
'           Set frmC = New frmCal
'
'           esq = cmdAux(Index).Left
'           dalt = cmdAux(Index).Top
'
'           Set obj = cmdAux(Index).Container
'
'           While cmdAux(Index).Parent.Name <> obj.Name
'                esq = esq + obj.Left
'                dalt = dalt + obj.Top
'                Set obj = obj.Container
'           Wend
'
'           menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
'
'           frmC.Left = esq + cmdAux(Index).Parent.Left + 30
'           frmC.Top = dalt + cmdAux(Index).Parent.Top + cmdAux(Index).Height + menu - 40
'
'
'           frmC.NovaData = Now
'           indice = Index + 2
'           Me.cmdAux(0).Tag = Index
'
'           PonerFormatoFecha txtaux1(indice)
'           If txtaux1(indice).Text <> "" Then frmC.NovaData = CDate(txtaux1(indice).Text)
'
'           Screen.MousePointer = vbDefault
'           frmC.Show vbModal
'           Set frmC = Nothing
'           PonerFoco txtaux1(indice)
'
'        Case 2, 3 'cuentas contables de cliente y proveedor
'            If vSeccion Is Nothing Then Exit Sub
'
'            indice = Index + 2
'            Set frmCtas = New frmCtasConta
'            frmCtas.NumDigit = 0
'            frmCtas.DatosADevolverBusqueda = "0|1|"
'            frmCtas.CodigoActual = txtaux1(indice).Text
'            frmCtas.Show vbModal
'            Set frmCtas = Nothing
'            PonerFoco txtaux1(indice)
'
'
'        Case 5 'codigo de iva
'            Set frmTIva = New frmTipIVAConta
'            frmTIva.DeConsulta = True
'            frmTIva.DatosADevolverBusqueda = "0|1|"
'            frmTIva.CodigoActual = txtaux1(6).Text
'            frmTIva.Show vbModal
'            Set frmTIva = Nothing
'            PonerFoco txtaux1(6)
'
'    End Select
'    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
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

Private Sub Combo15_Change(Index As Integer)
    If Combo15(Index).ListIndex <> -1 Then
        Label19(Index).Caption = "x " & PesoEnvase(Combo15(Index).ItemData(Combo15(Index).ListIndex))
        Text1_LostFocus (13 + Index)
    End If
End Sub

Private Sub Combo15_Click(Index As Integer)
    If Combo15(Index).ListIndex <> -1 Then
        Label19(Index).Caption = "x " & PesoEnvase(Combo15(Index).ItemData(Combo15(Index).ListIndex))
        Text1_LostFocus (13 + Index)
    End If
End Sub

Private Sub Combo15_KeyPress(Index As Integer, KeyAscii As Integer)
    If Combo15(Index).ListIndex = -1 Then
        Label19(Index).Caption = "x " & PesoEnvase(Combo15(Index).ItemData(Combo15(Index).ListIndex))
        Text1_LostFocus (13 + Index)
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then PrimeraVez = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim I As Integer
Dim Sql As String
Dim RS As ADODB.Recordset


    PrimeraVez = True
    
    ' ICONETS DE LA BARRA
    btnPrimero = 17 'index del botó "primero"
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
        .Buttons(11).Image = 26    'tarar tractor
        .Buttons(12).Image = 24  'paletizacion
        .Buttons(13).Image = 10  'Imprimir
        .Buttons(14).Image = 11  'Eixir
        'el 13 i el 14 son separadors
        .Buttons(btnPrimero).Image = 6  'Primer
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Següent
        .Buttons(btnPrimero + 3).Image = 9 'Últim
    End With
    
'    ' ******* si n'hi han llínies *******
'    'ICONETS DE LES BARRES ALS TABS DE LLÍNIA
'    For i = 0 To ToolAux.Count - 1
'        With Me.ToolAux(i)
'            .HotImageList = frmPpal.imgListComun_OM16
'            .DisabledImageList = frmPpal.imgListComun_BN16
'            .ImageList = frmPpal.imgListComun16
'            .Buttons(1).Image = 3   'Insertar
'            .Buttons(2).Image = 4   'Modificar
'            .Buttons(3).Image = 5   'Borrar
'        End With
'    Next i
'    ' ***********************************
    
    'cargamos la primera parte de la cadena xml
    v_cadena = "<?xml version=" & """1.0""" & " standalone=" & """yes""" & " ?><DATAPACKET "
    v_cadena = v_cadena & "Version=" & """1.0""" & "><METADATA><FIELDS>"
    v_cadena = v_cadena & "<FIELD attrname=" & """notacamp""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """fechaent""" & " fieldtype=" & """dateTime""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """codprodu""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """codvarie""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """codsocio""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """codcampo""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """kilosbru""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """kilosnet""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """numcajo1""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """numcajo2""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """numcajo3""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """numcajo4""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """numcajo5""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """matricul""" & " fieldtype=" & """string""" & " WIDTH=" & """10""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """codcapat""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """identifi""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """altura""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """zona""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "</FIELDS></METADATA><ROWDATA>"
    
    For I = 0 To imgAyuda.Count - 1
        imgAyuda(I).Picture = frmPpal.ImageListB.ListImages(10).Picture
    Next I

    'cargar IMAGES de busqueda
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    
    
    For I = 0 To 4
        Me.Label15(I).Caption = ""
        Me.Label19(I).Caption = ""
    Next I
    
    
    ' *** si n'hi han tabs, per a que per defecte sempre es pose al 1r***
    Me.SSTab1.Tab = 0
    ' *******************************************************************
    
    LimpiarCampos   'Neteja els camps TextBox
    
'    If vParamAplic.SeTaraTractor Then
'        Text1(3).TabIndex = 56
'        cmdAceptar.TabIndex = 17
'        cmdCancelar.TabIndex = 18
'    Else
'        Text1(3).TabIndex = 17
'        cmdAceptar.TabIndex = 18
'        cmdCancelar.TabIndex = 19
'    End If
    
    
    CodTipoMov = "NOC"

'    ' ******* si n'hi han llínies *******
'    DataGridAux(0).ClearFields
'    DataGridAux(1).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "rentradas"
    Ordenacion = " ORDER BY numnotac desc "
    '************************************************
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    '***** canviar el nom de la PK de la capçalera; repasar codEmpre *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where codcampo=-1"
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
    ' Cargamos los combos para insertar
    Sql = "select codtipen, nomtipen, pesocaja, escaja, numorden, 0 from confenva where not numorden is null "
    Sql = Sql & " union "
    Sql = Sql & " select codtipen, nomtipen, pesocaja, escaja, numorden, 1 from confenva where numorden is null "
    Sql = Sql & " order by 6, 5"
    Sql = Sql & " limit 5 "

    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not RS.EOF
        PosicionarCombo Me.Combo15(I), RS.Fields(0).Value
         
        Label19(I) = "x " & Format(RS.Fields(2).Value, "##0.00")
        I = I + 1
        RS.MoveNext
    Wend
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1 'búsqueda
        ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
        Text1(0).BackColor = vbYellow 'codclien
        ' ****************************************************************************
    End If
End Sub


Private Sub LimpiarCampos()
Dim I As Integer

    On Error Resume Next
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
'    Me.chkAbonos(0).Value = 0
    
    For I = 0 To Combo1.Count - 1
        Combo1(I).ListIndex = -1
'        Me.chkAbonos(I).Value = 0
    Next I
    
    ' *** si n'hi han combos a la capçalera ***
    ' *****************************************

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
Dim I As Integer, NumReg As Byte
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
    CmdAceptar.visible = b
       
    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a més neteja els camps Text1
    
' cambio la siguiente expresion por la de abajo
'    BloquearText1 Me, Modo
    For I = 0 To Text1.Count - 1
        BloquearTxt Text1(I), Not (Modo = 1 Or Modo = 3 Or Modo = 4)
    Next I
    
    BloquearCombo Me, Modo
    
    If Modo = 3 Then
        Combo1(1).ListIndex = 1
        Combo1(2).ListIndex = 0
    End If
    
    If vParamAplic.NroNotaManual Then
        'claveprimaria
        BloquearTxt Text1(0), Not (Modo = 1 Or (Modo = 3 And vParamAplic.NroNotaManual) Or Modo = 4)
    Else
        b = (Modo <> 1)
        'Campos Nº entrada bloqueado y en azul
        BloquearTxt Text1(0), b, True
    End If
    
    'taras desbloqueadas unicamente para buscar
    For I = 18 To 20
        BloquearTxt Text1(I), Not (Modo = 1)
    Next I
    For I = 8 To 9
        BloquearTxt Text1(I), Not (Modo = 1)
    Next I
    
    PonerTarasVisibles

    For I = 22 To 22
        BloquearTxt Text1(I), Not (Modo = 3 Or Modo = 4 Or Modo = 1)
    Next I
    
    BloquearTxt Text1(3), Not (((Modo = 3) And Not vParamAplic.SeTaraTractor) Or Modo = 1 Or Modo = 4 Or Modo = 5)
    BloquearTxt Text1(24), Not (Modo = 3 Or Modo = 4 Or Modo = 1)
    
    ' **** si n'hi han imagens de buscar en la capçalera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    For I = 0 To imgFec.Count - 1
        BloquearImgFec Me, I, Modo
    Next I
'    BloquearImgFec Me, 25, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
    ' ********************************************************
    ' *** si n'hi han llínies i imagens de buscar que no estiguen als grids ******
    'Llínies Departaments
'    b = (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) And (NumTabMto = 0))
    ' ****************************************************************************
            
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    For I = 0 To 4
        Combo15(I).Enabled = Not (Modo = 0 Or Modo = 2)
        
        If Not Combo15(I).Enabled Then
            Combo15(I).BackColor = &H80000018
        Else
            Combo15(I).BackColor = vbWhite
        End If
    Next I
    
    
    PonerLongCampos

'     '-----------------------------
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
Dim I As Byte
    
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
    Toolbar1.Buttons(7).Enabled = b
    Me.mnNuevo.Enabled = b
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) 'And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(9).Enabled = b
    Me.mnEliminar.Enabled = b
    'tara tractor
    Toolbar1.Buttons(11).Enabled = b
    'Paletizacion
    Toolbar1.Buttons(12).Enabled = b And vParamAplic.HayTraza
    'Imprimir
    'Toolbar1.Buttons(12).Enabled = (b Or Modo = 0)
    Toolbar1.Buttons(13).Enabled = b
       
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
'    indice = CByte(Me.cmdAux(0).Tag + 2)
'    txtaux1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmC1_Selec(vFecha As Date)
Dim indice As Byte
    indice = CByte(Me.imgFec(0).Tag)
    Text1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmPar_DatoSeleccionado(CadenaSeleccion As String)
'partidas
Dim Zona As String
Dim Poblacion As String
Dim otroCampo As String
Dim CodPobla As String

    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo de partida
    FormateaCampo Text1(3)
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre de partida
    Text3(3).Text = RecuperaValor(CadenaSeleccion, 5) 'codzona
    Text4(3).Text = RecuperaValor(CadenaSeleccion, 6) 'nomzona
    Text5(3).Text = RecuperaValor(CadenaSeleccion, 4)
    
    
'    Zona = ""
'    Text3(3).Text = ""
'    Text4(3).Text = ""
'    Text4(3).Text = ""
'
'    Zona = RecuperaValor(CadenaSeleccion, 3)
'    Text3(3).Text = Zona
'    otroCampo = "codpobla"
'    If Zona <> "" Then
'        Text4(3).Text = DevuelveDesdeBDNew(cAgro, "rzonas", "nomzonas", "codzona", Zona, "N", otroCampo)
'        If otroCampo <> "" Then
'            CodPobla = otroCampo
'            If CodPobla <> "" Then Text5(3).Text = DevuelveDesdeBDNew(cAgro, "rpueblos", "despobla", "codpobla", CodPobla, "T")
'        End If
'    End If
    
End Sub

Private Sub frmCamp_DatoSeleccionado(CadenaSeleccion As String)
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo de campo
    FormateaCampo Text1(5)
    If EstaCampoDeAlta(Text1(5).Text) Then
        PonerDatosCampo Text1(5).Text
    Else
        MsgBox "El campo no está dado de alta. Reintroduzca.", vbExclamation
        Text1(5).Text = ""
        PonerFoco Text1(5)
    End If
End Sub

Private Sub frmcap_DatoSeleccionado(CadenaSeleccion As String)
    Text1(12).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo de capataz
    FormateaCampo Text1(12)
    Text2(12).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre de capataz
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
    PonerDatosCampo Text1(5)
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    Text1(1).Text = RecuperaValor(CadenaSeleccion, 1) 'codsocio
    FormateaCampo Text1(1)
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2) 'nomsocio
End Sub

Private Sub frmTar_DatoSeleccionado(CadenaSeleccion As String)
    Text1(7).Text = RecuperaValor(CadenaSeleccion, 1) ' codigo tarifa
    FormateaCampo Text1(7)
    Text2(7).Text = RecuperaValor(CadenaSeleccion, 2) ' nombre tarifa
End Sub

Private Sub frmTrans_DatoSeleccionado(CadenaSeleccion As String)
    Text1(6).Text = RecuperaValor(CadenaSeleccion, 1) ' codigo transportista
    FormateaCampo Text1(6)
    Text2(6).Text = RecuperaValor(CadenaSeleccion, 2) ' nombre transportista
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codvariedad
    FormateaCampo Text1(2)
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre variedad
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
       End Select
       
       Me.imgFec(0).Tag = indice
       
       PonerFormatoFecha Text1(indice)
       If Text1(indice).Text <> "" Then frmC1.NovaData = CDate(Text1(indice).Text)
    
       Screen.MousePointer = vbDefault
       frmC1.Show vbModal
       Set frmC1 = Nothing
       PonerFoco Text1(indice)
    
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
Dim I As Integer



    If Data1.Recordset.EOF Then Exit Sub
    
    '[Monica]05/06/2014: para el caso de Natural la impresion de entradas es por una impresora de tickets
    If vParamAplic.Cooperativa = 9 Then
    
        ActivaTicket

        '-- Impresion directa
        ImprimirElTicketDirecto2 Text1(0).Text, CDate(Text1(10).Text), True
        
        DesactivaTicket
        
        Exit Sub
    End If
    
    
    
    Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
    Dim nomDocu As String 'Nombre de Informe rpt de crystal
    Dim ImprimeDirecto As Integer
     
    indRPT = 25 'Ticket de Entrada
     
    If Not PonerParamRPT(indRPT, "", 1, nomDocu, , ImprimeDirecto) Then Exit Sub
    'Nombre fichero .rpt a Imprimir
    ' he añadido estas dos lineas para que llame al rpt correspondiente
    
    If ImprimeDirecto = 0 Then
        frmImprimir.NombreRPT = nomDocu
        
        ActivaTicket
                
        With frmVisReport
            .FormulaSeleccion = "{rentradas.numnotac}=" & Data1.Recordset!numnotac
            .SoloImprimir = True
            .OtrosParametros = ""
            .NumeroParametros = 1
            .MostrarTree = False
            .Informe = App.Path & "\informes\" & nomDocu    ' "ValEntrada.rpt"
            .InfConta = False
            .ConSubInforme = True
            .SubInformeConta = ""
            .Opcion = 0
            .ExportarPDF = False
            .Show vbModal
        End With
        
        DesactivaTicket
    Else
        NroCopias = InputBox("Introduzca el Número de Copias:", "", , 5000, 4000)
    
        If NroCopias = "" Then Exit Sub
        If ComprobarCero(NroCopias) = 0 Then Exit Sub
        
        ' imprimimos
        If EsNumerico(NroCopias) Then
            ' impresion directa por la printer
'             ImprimirEntradaDirectaPrinter Text1(0).Text, CInt(NroCopias)
            ' impresion directa por LPT
            
'[Monica]31/10/2013: antes directo por lpt
'            ImprimirEntradaDirectaLPT Text1(0).Text, CInt(NroCopias)
' ahora para catadau
            frmImprimir.NombreRPT = nomDocu
            
            ActivaTicket
        
            For I = 1 To NroCopias
        
                cadParam = "|pPagina=" & I & "|"
                
                With frmVisReport
                    .FormulaSeleccion = "{rentradas.numnotac}=" & Data1.Recordset!numnotac
                    .SoloImprimir = True
                    .OtrosParametros = cadParam ' ""
                    .NumeroParametros = 1
                    .MostrarTree = False
                    .Informe = App.Path & "\informes\" & nomDocu    ' "ValEntrada.rpt"
                    .InfConta = False
                    .ConSubInforme = True
                    .SubInformeConta = ""
                    .Opcion = 0
                    .ExportarPDF = False
                    .Show vbModal
                End With
            Next I
            DesactivaTicket
        
        End If

    End If
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

Private Sub mnPaletizacion_Click()
Dim Sql As String
Dim RS As ADODB.Recordset
Dim cajas As Currency
Dim Cad As String

    If vParamAplic.HayTraza = False Then Exit Sub
    
    frmEntBascula2.crear = 1
    
    Sql = "select count(*) from trzpalets where numnotac = " & Trim(Data1.Recordset!numnotac)
    If TotalRegistros(Sql) <> 0 Then
        Cad = "La paletización para esta entrada ya está realizada." & vbCrLf
        Cad = Cad & vbCrLf & "            ¿ Desea crearla de nuevo ? "
        If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            frmEntBascula2.crear = 0
        End If
    End If
    
    
    cajas = 0
'    Cajas = DBLet(Data1.Recordset!numcajo1, "N") + _
'            DBLet(Data1.Recordset!numcajo2, "N") + _
'            DBLet(Data1.Recordset!numcajo3, "N") + _
'            DBLet(Data1.Recordset!numcajo4, "N") + _
'            DBLet(Data1.Recordset!numcajo5, "N")
            
    ' ahora las cajas se suman si rparam.escaja es true
    If EsCaja(CStr(DBLet(Data1.Recordset!tipocajo1, "N"))) Then cajas = cajas + DBLet(Data1.Recordset!numcajo1, "N")
    If EsCaja(CStr(DBLet(Data1.Recordset!tipocajo2, "N"))) Then cajas = cajas + DBLet(Data1.Recordset!numcajo2, "N")
    If EsCaja(CStr(DBLet(Data1.Recordset!tipocajo3, "N"))) Then cajas = cajas + DBLet(Data1.Recordset!numcajo3, "N")
    If EsCaja(CStr(DBLet(Data1.Recordset!tipocajo4, "N"))) Then cajas = cajas + DBLet(Data1.Recordset!numcajo4, "N")
    If EsCaja(CStr(DBLet(Data1.Recordset!tipocajo5, "N"))) Then cajas = cajas + DBLet(Data1.Recordset!numcajo5, "N")
    
    frmEntBascula2.NumNota = ImporteSinFormato(Data1.Recordset!numnotac)
    frmEntBascula2.NumCajones = CStr(cajas)
    frmEntBascula2.Numkilos = ImporteSinFormato(Text1(11).Text)
    frmEntBascula2.Codsocio = Text1(1).Text
    frmEntBascula2.CodCampo = Text1(5).Text
    frmEntBascula2.codvarie = Text1(2).Text
    frmEntBascula2.Fecha = Text1(10).Text
    frmEntBascula2.Hora = Text1(22).Text

    
    frmEntBascula2.Show vbModal


End Sub


Private Sub mnTararTractor_Click()

    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonTarar

End Sub



Private Sub mnVerTodos_Click()
    BotonVerTodos
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
        Case 11 ' tarar tractor
            mnTararTractor_Click
        Case 12 ' Paletizacion
            mnPaletizacion_Click
        Case 13 'Imprimir
'            AbrirListado (10)
            mnImprimir_Click
        Case 14    'Eixir
            mnSalir_Click
            
        Case btnPrimero To btnPrimero + 3 'Fleches Desplaçament
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonBuscar()
Dim I As Integer
' ***** Si la clau primaria de la capçalera no es Text1(0), canviar-ho en <=== *****
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
    
    If Text1(22).Text <> "" Then
        Text1(4).Text = Text1(22).Text
        Text1(4).Tag = Replace(Text1(8).Tag, "FH", "FHH")
    End If

    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
    Text1(4).Tag = Replace(Text1(4).Tag, "FHH", "FH")
    
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
    Cad = Cad & "Nota|rentradas.numnotac|N|0000000|10·"
    Cad = Cad & "Socio|rentradas.codsocio|N|000000|9·"
    Cad = Cad & "Nombre|rsocios.nomsocio|T||45·"
    Cad = Cad & "Variedad|variedades.nomvarie|T||40·"
    
    NombreTabla1 = "(rentradas inner join rsocios on rentradas.codsocio = rsocios.codsocio) inner join variedades on rentradas.codvarie = variedades.codvarie "
    
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = NombreTabla1
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        frmB.vDevuelve = "0|" '*** els camps que volen que torne ***
        frmB.vTitulo = "Entradas de Báscula" ' ***** repasa açò: títol de BuscaGrid *****
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
Dim I As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    Cad = ""
    I = 0
    Do
        J = I + 1
        I = InStr(J, DatosADevolverBusqueda, "|")
        If I > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, I - J)
            J = Val(Aux)
            Cad = Cad & Text1(J).Text & "|"
        End If
    Loop Until I = 0
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
Dim Sql As String
Dim RS As ADODB.Recordset
Dim I As Integer
Dim Minimo As Integer

    LimpiarCampos 'Huida els TextBox
    PonerModo 3
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la capçalera *******
'    Text1(0).Text = SugerirCodigoSiguienteStr("rentradas", "numnotac")
'    FormateaCampo Text1(0)
       
    If vParamAplic.NroNotaManual Then
        PonerFoco Text1(0)
    Else
        PonerFoco Text1(2) '*** 1r camp visible que siga PK ***
    End If
    ' ***********************************************************
    Text1(10).Text = Now
    Text1(22).Text = Mid(Now, 12, 8)
    ' *** si n'hi han camps de descripció a la capçalera ***
    PosarDescripcions
    ' ******************************************************

    ' *** si n'hi han tabs, em posicione al 1r ***
    Me.SSTab1.Tab = 0
    ' ********************************************
    
    
    Sql = "select codtipen, nomtipen, pesocaja, escaja, numorden, 0 from confenva where not numorden is null "
    Sql = Sql & " union "
    Sql = Sql & " select codtipen, nomtipen, pesocaja, escaja, numorden, 1 from confenva where numorden is null "
    Sql = Sql & " order by 6, 5"
    Sql = Sql & " limit 5 "

    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not RS.EOF
        PosicionarCombo Me.Combo15(I), RS.Fields(0).Value
         
        Label19(I) = "x " & Format(RS.Fields(2).Value, "##0.00")
        I = I + 1
        RS.MoveNext
    Wend
    
    'tipo de envases
    

End Sub


Private Sub BotonModificar()



'    NumCajasAnt = CCur(ComprobarCero(Text1(13).Text)) + CCur(ComprobarCero(Text1(14).Text)) + _
'               CCur(ComprobarCero(Text1(15).Text)) + CCur(ComprobarCero(Text1(16).Text)) + _
'               CCur(ComprobarCero(Text1(17).Text))

    ' ahora el numero de cajas se suma unicamente si rparam.escaja es true
    NumCajasAnt = 0
    If EsCaja(CStr(DBLet(Data1.Recordset!tipocajo1, "N"))) Then NumCajasAnt = NumCajasAnt + CCur(ComprobarCero(Text1(13).Text))
    If EsCaja(CStr(DBLet(Data1.Recordset!tipocajo2, "N"))) Then NumCajasAnt = NumCajasAnt + CCur(ComprobarCero(Text1(14).Text))
    If EsCaja(CStr(DBLet(Data1.Recordset!tipocajo3, "N"))) Then NumCajasAnt = NumCajasAnt + CCur(ComprobarCero(Text1(15).Text))
    If EsCaja(CStr(DBLet(Data1.Recordset!tipocajo4, "N"))) Then NumCajasAnt = NumCajasAnt + CCur(ComprobarCero(Text1(16).Text))
    If EsCaja(CStr(DBLet(Data1.Recordset!tipocajo5, "N"))) Then NumCajasAnt = NumCajasAnt + CCur(ComprobarCero(Text1(17).Text))
    

    NumKilosAnt = CCur(ComprobarCero(Text1(11).Text))
    PonerModo 4

    ' *** bloquejar els camps visibles de la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    ' *************************************************************************
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(10)
    ' *********************************************************
End Sub

Private Sub BotonTarar()

    PonerModo 5

    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(3)
    ' *********************************************************
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
    Cad = "¿Seguro que desea eliminar la Entrada?"
    Cad = Cad & vbCrLf & "Número: " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
    Cad = Cad & vbCrLf & "Fecha : " & Data1.Recordset.Fields(1)
    Cad = Cad & vbCrLf & "Socio: " & Text2(1).Text
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
Dim I As Integer
Dim CodPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String
Dim Sql As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la capçalera
    
    Text1(22).Text = Mid(Text1(4).Text, 12, 8)
    
    Sql = "select porcbonif from rbonifentradas where codvarie = " & DBSet(Text1(2).Text, "N") & " and fechaent = " & DBSet(Text1(10).Text, "F")
    Text2(8).Text = Format(DevuelveValor(Sql), "#,##0.00")
    If Combo1(0).ListIndex = 1 Then
        Text2(8).Text = Format(0, "#,##0.00")
    End If
    
    PosarDescripcions

    ' ********************************************************************************
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu   'Activar opcions de menú según nivell
                        'de permisos de l'usuari
    
End Sub


Private Sub cmdCancelar_Click()
Dim I As Integer
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
                TerminaBloquear
                PonerModo 2
                PonerCampos
                ' *** primer camp visible de la capçalera ***
                PonerFoco Text1(0)
            
'            Select Case ModoLineas
'                Case 1 'afegir llínia
'                    ModoLineas = 0
'                    ' *** les llínies que tenen datagrid (en o sense tab) ***
'                    If NumTabMto = 0 Or NumTabMto = 1 Or NumTabMto = 2 Or NumTabMto = 4 Then
'                        DataGridAux(NumTabMto).AllowAddNew = False
'                        ' **** repasar si es diu Data1 l'adodc de la capçalera ***
'                        'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
'                        ' ********************************************************
'                        LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
'                        DataGridAux(NumTabMto).Enabled = True
'                        DataGridAux(NumTabMto).SetFocus
'
'                        ' *** si n'hi han camps de descripció dins del grid, els neteje ***
'                        'txtAux2(2).text = ""
'                        ' *****************************************************************
'
'                        ' ***  bloquejar i huidar els camps que estan fora del datagrid ***
'                        Select Case NumTabMto
'                            Case 0 'cuentas bancarias
'                                'BotonModificar
''                                BloquearTxt txtaux(11), True
''                                BloquearTxt txtaux(12), True
'                            Case 1 'secciones
'                                For I = 0 To txtaux1.Count - 1
'                                    txtaux1(I).Text = ""
'                                    BloquearTxt txtaux1(I), True
'                                Next I
'                                txtAux2(1).Text = ""
'                                txtAux2(4).Text = ""
'                                txtAux2(5).Text = ""
'                                BloquearTxt txtAux2(1), True
'                                BloquearTxt txtAux2(4), True
'                                BloquearTxt txtAux2(5), True
'                            Case 2 'telefonos
'                                For I = 0 To txtAux.Count
'                                    BloquearTxt txtAux(I), True
'                                Next I
'                        End Select
'                    ' *** els tabs que no tenen datagrid ***
'                    ElseIf NumTabMto = 3 Then
'                        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
'                        CargaFrame 3, True
'                    End If
'
'                    ' *** si n'hi han tabs ***
'                    SituarTab (NumTabMto)
'                    'SSTab1.Tab = 1
'                    'SSTab2.Tab = NumTabMto
'                    ' ************************
'
'                    If Not AdoAux(NumTabMto).Recordset.EOF Then
'                        AdoAux(NumTabMto).Recordset.MoveFirst
'                    End If
'
'                Case 2 'modificar llínies
'                    ModoLineas = 0
'
'                    ' *** si n'hi han tabs ***
'                    SituarTab (NumTabMto)
'                    'SSTab1.Tab = 1
'                    'SSTab2.Tab = NumTabMto
'                    ' ***********************
'
'                    PonerModo 4
'                    If Not AdoAux(NumTabMto).Recordset.EOF Then
'                        ' *** l'Index de Fields es el que canvie de la PK de llínies ***
'                        V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
'                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
'                        ' ***************************************************************
'                    End If
'
'                    ' ***  bloquejar els camps fora dels grids ***
'
'                    ' ***  bloquejar els camps fora dels grids ***
'                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
'            End Select
'
'            If NumTabMto = 1 Then
'                If Not vSeccion Is Nothing Then
'                    vSeccion.CerrarConta
'                    Set vSeccion = Nothing
'                End If
'            End If
'
'            PosicionarData
'
'            ' *** si n'hi han llínies en grids i camps fora d'estos ***
'            If Not AdoAux(NumTabMto).Recordset.EOF Then
'                DataGridAux_RowColChange NumTabMto, 1, 1
'            Else
'                LimpiarCamposFrame NumTabMto
'            End If
'            ' *********************************************************
    End Select
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Sql As String
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOk = False
    
    Text1(4).Text = Format(Text1(10).Text, "dd/mm/yyyy") & " " & Format(Text1(22).Text, "HH:MM:SS")
    
    b = CompForm2(Me, 1)
    If Not b Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
        'comprobar si existe ya el cod. del campo clave primaria
        If ExisteCP(Text1(0)) Then b = False
    End If
    
    'miramos si hay otros campos con la misma ubicacion
    If b And (Modo = 3 Or Modo = 4) Then
        If b Then
            If Not EstaSocioDeAlta(Text1(1).Text) Then
            ' comprobamos que el socio no este dado de baja
                Sql = "El socio introducido está dado de baja. Reintroduzca. " & vbCrLf & vbCrLf
                MsgBox Sql, vbExclamation
                b = False
                PonerFoco Text1(1)
            End If
        End If
        
        If b Then
            ' comprobamos que el campo no esté dado de baja
            If Not EstaCampoDeAlta(Text1(5).Text) Then
                Sql = "El campo introducido está dado de baja. Reintroduzca. " & vbCrLf & vbCrLf
                MsgBox Sql, vbExclamation
                b = False
                PonerFoco Text1(5)
            End If
        End If
        
        If b Then
            ' comprobamos que el campo es de socio variedad
            If Not EsCampoSocioVariedad(Text1(5).Text, Text1(1).Text, Text1(2).Text) Then
                Sql = "El campo introducido no es del socio variedad. Reintroduzca. " & vbCrLf & vbCrLf
                MsgBox Sql, vbExclamation
                b = False
                PonerFoco Text1(5)
            End If
        End If
        
        If b Then
            ' si el nro de nota es manual comprobamos que no exita en ningun sitio
            If vParamAplic.NroNotaManual And Modo = 3 Then
                If ExisteNota(Text1(0).Text) Then
                    MsgBox "Nro de Nota ya existe. Reintroduzca.", vbExclamation
                    b = False
                    PonerFoco Text1(0)
                End If
            End If
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
    Cad = "(numnotac=" & Text1(0).Text & ")"
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
Dim Mens As String

    On Error GoTo FinEliminar

    conn.BeginTrans
    ' ***** canviar el nom de la PK de la capçalera, repasar codEmpre *******
    vWhere = " WHERE numnotac=" & Data1.Recordset!numnotac
        ' ***********************************************************************
        
    Mens = "Actualizar chivato"
    b = ActualizarChivato(Mens, "Z")
        
        
    ' ***** elimina les llínies ****
    conn.Execute "DELETE FROM trzpalets where numnotac = " & Trim(CStr(Data1.Recordset!numnotac))

    conn.Execute "Delete from " & NombreTabla & vWhere
       
FinEliminar:
    If Err.Number <> 0 Or Not b Then
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
                Else
                    If EstaSocioDeAlta(Text1(Index)) Then
                        PonerCamposSocioVariedad
                    Else
                        MsgBox "El socio está dado de baja. Reintroduzca.", vbExclamation
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    End If
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        Case 2 'VARIEDAD
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "variedades", "nomvarie")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Variedad: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmVar = New frmComVar
                        frmVar.DatosADevolverBusqueda = "0|1|"
'                        frmVar.NuevoCodigo = Text1(Index).Text
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
                    If (Modo = 3 Or Modo = 4) And EsVariedadGrupo6(Text1(Index).Text) Then
                        MsgBox "Esta variedad es del Grupo de Bodega. Revise.", vbExclamation
                        PonerFoco Text1(Index)
                    Else
                        PonerCamposSocioVariedad
                    End If
                End If
            Else
                Text2(Index).Text = ""
            End If
        
                
        Case 5 'codigo de campo
            If Modo = 1 Then Exit Sub
            If PonerFormatoEntero(Text1(Index)) Then
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
                        If Not EsCampoSocioVariedad(Text1(Index).Text, Text1(1).Text, Text1(2).Text) Then
                            MsgBox "El campo no es del Socio Variedad. Reintroduzca.", vbExclamation
                            PonerFoco Text1(Index)
                        Else
                            PonerDatosCampo (Text1(Index))
                            If Modo = 3 Then
                                Combo1(1).ListIndex = DevuelveValor("select recolect from rcampos where codcampo = " & DBSet(Text1(5).Text, "N"))
                            End If
                        End If
                    End If
                End If
            End If
        
                
        Case 6 'transportistas
            If Text1(Index).Text <> "" Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "rtransporte", "nomtrans")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Transportista: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmTrans = New frmManTranspor
                        frmTrans.DatosADevolverBusqueda = "0|1|"
                        frmTrans.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmTrans.Show vbModal
                        Set frmTrans = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                Else
                    If Modo = 3 Then ' solo si estamos insertando metemos la tara del vehiculo
                        Sql = "select taravehi from rtransporte where codtrans = " & DBSet(Text1(Index), "T")
                        Text1(3).Text = DevuelveValor(Sql)
                        PonerFormatoEntero Text1(3)
                    End If
                End If
            Else
                Text2(Index).Text = ""
            End If
                
        Case 7 'tarifa de transporte
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rtarifatra", "nomtarif")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Tarifa de Transporte: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmTar = New frmManTarTra
                        frmTar.DatosADevolverBusqueda = "0|1|"
                        frmTar.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmTar.Show vbModal
                        Set frmTar = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
                
                
                
        Case 12 'capataz
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rcapataz", "nomcapat")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Capataz: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCap = New frmManCapataz
                        frmCap.DatosADevolverBusqueda = "0|1|"
                        frmCap.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmCap.Show vbModal
                        Set frmCap = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 13, 14, 15, 16, 17, 21 'pesos
            If Modo = 1 Then Exit Sub
            PonerFormatoEntero Text1(Index)
            If Text1(Index).Text <> "" Then CalcularTaras
        
        Case 3 ' TARA DE TRACTOR
            If Modo = 1 Then Exit Sub
            If Text1(Index).Text <> "" Then PonerFormatoEntero Text1(Index)
            If vParamAplic.SeTaraTractor Then
                PonerModo 4
                CalcularTaras
                PonerModo 5
                cmdAceptar_Click
            Else
                CalcularTaras
            End If
            
        Case 24 ' OTRAS TARAS
            If Modo = 1 Then Exit Sub
            PonerFormatoEntero Text1(Index)
            If Text1(Index).Text <> "" Then CalcularTaras
        
        Case 22 'formato hora
            If Modo = 1 Then Exit Sub
            PonerFormatoHora Text1(Index)
        
            
        Case 10 'Fecha no comprobaremos que esté dentro de campaña
                    'Fecha de alta y fecha de baja
            '[Monica]28/08/2013: antes no comprobamos que la fecha esté en la campaña ahora sí
            PonerFormatoFecha Text1(Index), True
            
    End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 12: KEYBusqueda KeyAscii, 0 'capataz
                Case 6: KEYBusqueda KeyAscii, 3 'transportista
                Case 7: KEYBusqueda KeyAscii, 4 'tarifa
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

    Text2(1).Text = PonerNombreDeCod(Text1(1), "rsocios", "nomsocio", "codsocio", "N")
    Text2(12).Text = PonerNombreDeCod(Text1(12), "rcapataz", "nomcapat", "codcapat", "N")
    Text2(2).Text = PonerNombreDeCod(Text1(2), "variedades", "nomvarie", "codvarie", "N")
    Text2(6).Text = PonerNombreDeCod(Text1(6), "rtransporte", "nomtrans", "codtrans", "T")
    Text2(7).Text = PonerNombreDeCod(Text1(7), "rtarifatra", "nomtarif", "codtarif", "N")
    
    PonerDatosCampo Text1(5).Text
    
'    If Text1(5).Text <> "" Then
'        Text2(5).Text = Round2(ImporteSinFormato(Text1(5).Text) / cFaneca, 4)
'        PonerFormatoDecimal Text2(5), 7
'    End If
'
'    If Text1(6).Text <> "" Then
'        Text2(6).Text = Round2(ImporteSinFormato(Text1(6).Text) / cFaneca, 4)
'        PonerFormatoDecimal Text2(6), 7
'    End If
'
'    If Text1(7).Text <> "" Then
'        Text2(7).Text = Round2(ImporteSinFormato(Text1(7).Text) / cFaneca, 4)
'        PonerFormatoDecimal Text2(7), 7
'    End If
    
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
'Dim sql As String
'Dim vWhere As String
'Dim eliminar As Boolean
'
'    On Error GoTo Error2
'
'    ModoLineas = 3 'Posem Modo Eliminar Llínia
'
'    If Modo = 4 Then 'Modificar Capçalera
'        cmdAceptar_Click
'        If ModoLineas = 0 Then Exit Sub
'    End If
'
'    NumTabMto = Index
'    PonerModo 5, Index
'
'    If AdoAux(Index).Recordset.EOF Then Exit Sub
'    If Not SepuedeBorrar(Index) Then Exit Sub
'    NumTabMto = Index
'    eliminar = False
'
'    vWhere = ObtenerWhereCab(True)
'
'    ' ***** independentment de si tenen datagrid o no,
'    ' canviar els noms, els formats i el DELETE *****
'    Select Case Index
'        Case 0 'telefonos
'            sql = "¿Seguro que desea eliminar el telefono?"
'            sql = sql & vbCrLf & "Teléfono: " & AdoAux(Index).Recordset!idtelefono & " - " & AdoAux(Index).Recordset!imei
'            If MsgBox(sql, vbQuestion + vbYesNo) = vbYes Then
'                eliminar = True
'                sql = "DELETE FROM rsocios_telefonos"
'                sql = sql & vWhere & " AND idtelefono= " & DBLet(AdoAux(Index).Recordset!idtelefono, "T")
'            End If
'        Case 1 'secciones
'            sql = "¿Seguro que desea eliminar la sección?"
'            sql = sql & vbCrLf & "Sección: " & AdoAux(Index).Recordset!codsecci & " - " & AdoAux(Index).Recordset!nomsecci
'            If MsgBox(sql, vbQuestion + vbYesNo) = vbYes Then
'                eliminar = True
'                sql = "DELETE FROM rsocios_seccion"
'                sql = sql & vWhere & " AND codsecci= " & DBLet(AdoAux(Index).Recordset!codsecci, "N")
'            End If
'
'    End Select
'
'    If eliminar Then
'        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
'        TerminaBloquear
'        Conn.Execute sql
'        ' *** si n'hi han tabs sense datagrid, posar l'If ***
'        If Index <> 3 Then _
'            CargaGrid Index, True
'        ' ***************************************************
'        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
''            PonerCampos
'
'        End If
'        ' *** si n'hi han tabs sense datagrid ***
'        If Index = 3 Then CargaFrame 3, True
'        ' ***************************************
'        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
'        ' *** si n'hi han tabs ***
'        SituarTab (NumTabMto)
'        ' ************************
'    End If
'
'    ModoLineas = 0
'    PosicionarData
'
'    Exit Sub
'Error2:
'    Screen.MousePointer = vbDefault
'    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub


Private Sub BotonAnyadirLinea(Index As Integer)
'Dim NumF As String
'Dim vWhere As String, vTabla As String
'Dim anc As Single
'Dim I As Integer
'
'    ModoLineas = 1 'Posem Modo Afegir Llínia
'
'    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
'        cmdAceptar_Click
'        If ModoLineas = 0 Then Exit Sub
'    End If
'
'    NumTabMto = Index
'    PonerModo 5, Index
'
'    ' *** bloquejar la clau primaria de la capçalera ***
'    BloquearTxt Text1(0), True
'    ' **************************************************
'
'    ' *** posar el nom del les distintes taules de llínies ***
'    Select Case Index
'        Case 0: vTabla = "rsocios_telefonos"
'        Case 1: vTabla = "rsocios_seccion"
'    End Select
'    ' ********************************************************
'
'    vWhere = ObtenerWhereCab(False)
'
'    Select Case Index
'        Case 0 ' *** pose els index dels tabs de llínies que tenen datagrid ***
'            ' *** canviar la clau primaria de les llínies,
'            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
'            If Index <> 4 Then ' *** els index als que no volem sugerir-li un codi ***
'                NumF = SugerirCodigoSiguienteStr(vTabla, "idtelefono", vWhere)
'            Else
'                NumF = ""
'            End If
'            ' ***************************************************************
'
'            AnyadirLinea DataGridAux(Index), AdoAux(Index)
'
'            anc = DataGridAux(Index).Top
'            If DataGridAux(Index).Row < 0 Then
'                anc = anc + 210
'            Else
'                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
'            End If
'
'            LLamaLineas Index, ModoLineas, anc
'
'            Select Case Index
'                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
'                Case 0 'cuentas
'                    For I = 0 To txtAux.Count - 1
'                        txtAux(I).Text = ""
'                    Next I
'                    txtAux(0).Text = Text1(0).Text 'codsocio
'                    txtAux(1).Text = NumF 'idtelefono
'                    PonerFoco txtAux(1)
'
'            End Select
'
'         Case 1   'secciones
'            ' *** canviar la clau primaria de les llínies,
'            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
'            If Index <> 4 Then ' *** els index als que no volem sugerir-li un codi ***
'                NumF = SugerirCodigoSiguienteStr(vTabla, "codsecci", vWhere)
'            Else
'                NumF = ""
'            End If
'            ' ***************************************************************
'
'            AnyadirLinea DataGridAux(Index), AdoAux(Index)
'
'            anc = DataGridAux(Index).Top
'            If DataGridAux(Index).Row < 0 Then
'                anc = anc + 210
'            Else
'                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
'            End If
'
'            LLamaLineas Index, ModoLineas, anc
'
'            Select Case Index
'                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
'                Case 1 'secciones
'                    For I = 0 To txtaux1.Count - 1
'                        txtaux1(I).Text = ""
'                    Next I
'                    txtaux1(0).Text = Text1(0).Text 'codsocio
'                    txtaux1(1).Text = NumF 'codseccion
'                    txtAux2(1).Text = ""
'                    txtAux2(4).Text = ""
'                    txtAux2(5).Text = ""
'                    txtAux2(0).Text = ""
'                    PonerFoco txtaux1(1)
'
'            End Select
'
'
''        ' *** si n'hi han llínies sense datagrid ***
''        Case 3
''            LimpiarCamposLin "FrameAux3"
''            txtaux(42).Text = text1(0).Text 'codclien
''            txtaux(43).Text = vSesion.Empresa
''            Me.cmbAux(28).ListIndex = 0
''            Me.cmbAux(29).ListIndex = 1
''            PonerFoco txtaux(25)
''        ' ******************************************
'    End Select
End Sub


Private Sub BotonModificarLinea(Index As Integer)
'    Dim anc As Single
'    Dim I As Integer
'    Dim J As Integer
'
'    If AdoAux(Index).Recordset.EOF Then Exit Sub
'    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub
'
'    ModoLineas = 2 'Modificar llínia
'
'    If Modo = 4 Then 'Modificar Capçalera
'        cmdAceptar_Click
'        If ModoLineas = 0 Then Exit Sub
'    End If
'
'    NumTabMto = Index
'    PonerModo 5, Index
'    ' *** bloqueje la clau primaria de la capçalera ***
'    BloquearTxt Text1(0), True
'    ' *********************************
'
'    Select Case Index
'        Case 0, 1 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
'            If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
'                I = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
'                DataGridAux(Index).Scroll 0, I
'                DataGridAux(Index).Refresh
'            End If
'
'            anc = DataGridAux(Index).Top
'            If DataGridAux(Index).Row < 0 Then
'                anc = anc + 210
'            Else
'                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
'            End If
'
'    End Select
'
'    Select Case Index
'        ' *** valor per defecte al modificar dels camps del grid ***
'        Case 0 'telefonos
'            For I = 0 To 16
'                txtAux(I).Text = DataGridAux(Index).Columns(I).Text
'            Next I
'
'        Case 1 'secciones
'            For I = 0 To 1
'                txtaux1(I).Text = DataGridAux(Index).Columns(I).Text
'            Next I
'            txtAux2(1).Text = DataGridAux(Index).Columns(2).Text
'            For I = 3 To 7
'                txtaux1(I - 1).Text = DataGridAux(Index).Columns(I).Text
'            Next I
'    End Select
'
'    LLamaLineas Index, ModoLineas, anc
'
'    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
'    Select Case Index
'        Case 0 'telefonos
'            PonerFoco txtAux(2)
'        Case 1 'secciones
'            PonerFoco txtaux1(2)
'            If txtaux1(1).Text <> "" Then
'                Set vSeccion = New CSeccion
'                If vSeccion.LeerDatos(txtaux1(1)) Then
'                    If vSeccion.AbrirConta Then
'                        If txtaux1(4).Text <> "" Then
'                            txtAux2(4).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", txtaux1(4).Text, "T")
'                        End If
'                        If txtaux1(5).Text <> "" Then
'                            txtAux2(5).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", txtaux1(5).Text, "T")
'                        End If
'                        If txtaux1(6).Text <> "" Then
'                            txtAux2(0).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", txtaux1(6).Text, "N")
'                        End If
'                    End If
'                End If
'            End If
'    End Select
'    ' ***************************************************************************************
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

'    ' *** si n'hi han tabs sense datagrid posar el If ***
'    If Index <> 3 Then DeseleccionaGrid DataGridAux(Index)
'    ' ***************************************************
'
'    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
'    Select Case Index
'        Case 0 'telefonos
'            For jj = 1 To 4
'                txtAux(jj).visible = b
'                txtAux(jj).Top = alto
'            Next jj
'        Case 1 'secciones
'            For jj = 1 To txtaux1.Count - 1
'                txtaux1(jj).visible = b
'                txtaux1(jj).Top = alto
'            Next jj
'            txtAux2(1).visible = b
'            txtAux2(1).Top = alto
'
'            For jj = 0 To cmdAux.Count - 1
'                cmdAux(jj).visible = b
'                cmdAux(jj).Top = txtaux1(3).Top
'                cmdAux(jj).Height = txtaux1(3).Height
'            Next jj
'    End Select
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
'Dim cadMen As String
'Dim Nuevo As Boolean
'
'    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
'
'    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
'    'mostrar mensajes ni hacer nada
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
'
'
'    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
'    Select Case Index
'        Case 2 'NIF
'            txtAux(Index).Text = UCase(txtAux(Index).Text)
'            ValidarNIF txtAux(Index).Text
'
'        Case 5 'NOMBRE
'            If txtAux(Index).Text <> "" Then txtAux(Index).Text = UCase(txtAux(Index).Text)
'
'        Case 12, 13 'ENTIDAD Y SUCURSAL BANCARIA
'            PonerFormatoEntero txtAux(Index)
'
'        Case 16
'            CmdAceptar.SetFocus
'    End Select
'
'    ' ******************************************************************************
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
'   If Not txtAux(Index).MultiLine Then ConseguirFocoLin txtAux(Index)
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    If Not txtAux(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
''    If Not txtAux(Index).MultiLine Then
'        If KeyAscii = teclaBuscar Then
'            If Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) Then
'                Select Case Index
'                    Case 4: KEYBusqueda KeyAscii, 7 'pais
'                    Case 10: KEYBusqueda KeyAscii, 3 'mercado
'                    Case 11: KEYBusqueda KeyAscii, 4 'cadena
'                End Select
'            End If
'        Else
'            KEYpress KeyAscii
'        End If
''    End If
End Sub


Private Function DatosOkLlin(nomframe As String) As Boolean
Dim RS As ADODB.Recordset
Dim Sql As String
Dim b As Boolean
Dim cant As Integer
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
        Case 0 'capataz
            Set frmCap = New frmManCapataz
            frmCap.DatosADevolverBusqueda = "0|1|"
            frmCap.CodigoActual = Text1(12).Text
            frmCap.Show vbModal
            Set frmCap = Nothing
            PonerFoco Text1(12)
        
       Case 1 'Socios
            Set frmSoc = New frmManSocios
'            frmSoc.DeConsulta = True
            frmSoc.DatosADevolverBusqueda = "0|1|"
'            frmSoc.CodigoActual = Text1(1).Text
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco Text1(1)
    
       Case 2 'Variedades
            Set frmVar = New frmComVar
'            frmVar.DeConsulta = True
            frmVar.DatosADevolverBusqueda = "0|1|"
'            frmVar.CodigoActual = Text1(2).Text
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco Text1(2)
    
       Case 3 'Transportista
            Set frmTrans = New frmManTranspor
            frmTrans.DeConsulta = True
            frmTrans.DatosADevolverBusqueda = "0|1|"
            frmTrans.CodigoActual = Text1(6).Text
            frmTrans.Show vbModal
            Set frmTrans = Nothing
            PonerFoco Text1(6)
            
       Case 4 ' Tarifa
            Set frmTar = New frmManTarTra
            frmTar.DeConsulta = True
            frmTar.DatosADevolverBusqueda = "0|1|"
            frmTar.CodigoActual = Text1(7).Text
            frmTar.Show vbModal
            Set frmTar = Nothing
            PonerFoco Text1(7)
       
       Case 5 ' Campos
'            Set frmCamp = New frmManCampos
''            frmCamp.DeConsulta = True
'            frmCamp.DatosADevolverBusqueda = "0|"
''            frmCamp.CodigoActual = Text1(5).Text
'            frmCamp.Show vbModal
'            Set frmCamp = Nothing
            PonerCamposSocioVariedad
            PonerFoco Text1(5)
    
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub


' *********************************************************************************
Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
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
        SSTab1.Tab = 1
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


'Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
'Dim b As Boolean
'Dim I As Byte
'Dim tots As String
'
'    On Error GoTo ECarga
'
'    tots = MontaSQLCarga(Index, enlaza)
'
'    'b = DataGridAux(Index).Enabled
'    'DataGridAux(Index).Enabled = False
'
''    AdoAux(Index).ConnectionString = Conn
''    AdoAux(Index).RecordSource = MontaSQLCarga(Index, enlaza)
''    AdoAux(Index).CursorType = adOpenDynamic
''    AdoAux(Index).LockType = adLockPessimistic
''    DataGridAux(Index).ScrollBars = dbgNone
''    AdoAux(Index).Refresh
''    Set DataGridAux(Index).DataSource = AdoAux(Index)
'
''    DataGridAux(Index).AllowRowSizing = False
''    DataGridAux(Index).RowHeight = 290
''    If PrimeraVez Then
''        DataGridAux(Index).ClearFields
''        DataGridAux(Index).ReBind
''        DataGridAux(Index).Refresh
''    End If
''
''    For i = 0 To DataGridAux(Index).Columns.Count - 1
''        DataGridAux(Index).Columns(i).AllowSizing = False
''    Next i
'
'    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
'
'
'    'DataGridAux(Index).Enabled = b
''    PrimeraVez = False
'
'    Select Case Index
'        Case 0 'telefonos
'            'si es visible|control|tipo campo|nombre campo|ancho control|
'            tots = "N||||0|;S|txtAux(1)|T|Telefono|900|;" 'codsocio,idtelefono
'            tots = tots & "S|txtAux(2)|T|NIF|1200|;"
'            tots = tots & "S|txtAux(3)|T|IMEI|3050|;"
'            tots = tots & "S|txtAux(4)|T|C.P|700|;"
'            tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
'            arregla tots, DataGridAux(Index), Me
'
'            DataGridAux(Index).Columns(2).Alignment = dbgLeft
'            DataGridAux(Index).Columns(3).Alignment = dbgLeft
'            DataGridAux(Index).Columns(4).Alignment = dbgLeft
'
'            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
''            BloquearTxt txtAux(14), Not b
''            BloquearTxt txtAux(15), Not b
'
''            If (enlaza = True) And (Not AdoAux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
''                For i = 5 To 16
''                    txtAux(i).Text = DataGridAux(Index).Columns(i).Text
''                Next i
''            Else
''                For i = 0 To 16
''                    txtAux(i).Text = ""
''                Next i
''            End If
''
'        Case 1 'secciones
'            'si es visible|control|tipo campo|nombre campo|ancho control|
'            tots = "N||||0|;S|txtaux1(1)|T|Cód.|800|;S|cmdAux(4)|B|||;" 'codsocio,codsecci
'            tots = tots & "S|txtAux2(1)|T|Nombre|3700|;"
'            tots = tots & "S|txtaux1(2)|T|F.Alta|1200|;S|cmdAux(0)|B|||;"
'            tots = tots & "S|txtaux1(3)|T|F.Baja|1200|;S|cmdAux(1)|B|||;"
'            tots = tots & "S|txtaux1(4)|T|Cta.Cliente|1200|;S|cmdAux(2)|B|||;"
'            tots = tots & "S|txtaux1(5)|T|Cta.Prov.|1200|;S|cmdAux(3)|B|||;"
'            tots = tots & "S|txtaux1(6)|T|Iva|800|;S|cmdAux(5)|B|||;"
'            arregla tots, DataGridAux(Index), Me
'
'            DataGridAux(Index).Columns(2).Alignment = dbgLeft
'            DataGridAux(Index).Columns(5).Alignment = dbgLeft
'            DataGridAux(Index).Columns(6).Alignment = dbgLeft
'
'            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
''            BloquearTxt txtAux(14), Not b
''            BloquearTxt txtAux(15), Not b
'
'            If (enlaza = True) And (Not AdoAux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
''                txtAux2(4).Text = PonerNombreCuenta(txtaux1(4), Modo)
''                txtAux2(5).Text = PonerNombreCuenta(txtaux1(5), Modo)
''                txtAux2(0).Text = PonerNombreDeCod(txtaux1(6), "tiposiva", "nombriva", "codigiva", "N", cConta)
'            Else
'                For I = 0 To 6
'                    txtaux1(I).Text = ""
'                Next I
'                txtAux2(0).Text = ""
'                txtAux2(1).Text = ""
'                txtAux2(4).Text = ""
'                txtAux2(5).Text = ""
'            End If
'    End Select
'    DataGridAux(Index).ScrollBars = dbgAutomatic
'
'    ' **** si n'hi han llínies en grids i camps fora d'estos ****
'    If Not AdoAux(Index).Recordset.EOF Then
'        DataGridAux_RowColChange Index, 1, 1
'    Else
'        LimpiarCamposFrame Index
'    End If
'    ' **********************************************************
'
'ECarga:
'    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
'End Sub
'

Private Sub InsertarLinea()
''Inserta registre en les taules de Llínies
'Dim nomframe As String
'Dim b As Boolean
'
'    On Error Resume Next
'
'    ' *** posa els noms del frames, tant si son de grid com si no ***
'    Select Case NumTabMto
'        Case 0: nomframe = "FrameAux0" 'telefonos
'        Case 1: nomframe = "FrameAux1" 'secciones
'    End Select
'    ' ***************************************************************
'
'    If DatosOkLlin(nomframe) Then
'        TerminaBloquear
'        If InsertarDesdeForm2(Me, 2, nomframe) Then
'            ' *** si n'hi ha que fer alguna cosa abas d'insertar
'            ' *************************************************
'            b = BLOQUEADesdeFormulario2(Me, Data1, 1)
'
'            '++monica: en caso de estar insertando seccion y que no existan las
'            'cuentas contables hacemos esto para que las inserte en contabilidad.
'            If NumTabMto = 1 Then
'               txtAux2(4).Text = PonerNombreCuenta(txtaux1(4), 3, Text1(0))
'               txtAux2(5).Text = PonerNombreCuenta(txtaux1(5), 3, Text1(0))
'            End If
'
'            Select Case NumTabMto
'                Case 0, 1 ' *** els index de les llinies en grid (en o sense tab) ***
'                    CargaGrid NumTabMto, True
'                    If b Then BotonAnyadirLinea NumTabMto
''                Case 3 ' *** els index dels tabs que NO tenen grid ***
''                    CargaFrame 3, True
''                    If b Then BotonModificar
''                    ModoLineas = 0
''                LLamaLineas NumTabMto, 0
'            End Select
'
'            SituarTab (NumTabMto)
'        End If
'    End If
End Sub


Private Sub ModificarLinea()
''Modifica registre en les taules de Llínies
'Dim nomframe As String
'Dim V As Integer
'Dim Cad As String
'    On Error Resume Next
'
'    ' *** posa els noms del frames, tant si son de grid com si no ***
'    Select Case NumTabMto
'        Case 0: nomframe = "FrameAux0" 'telefonos
'        Case 1: nomframe = "FrameAux1" 'secciones
'    End Select
'    ' **************************************************************
'
'    If DatosOkLlin(nomframe) Then
'        TerminaBloquear
'        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
'            ' *** si cal que fer alguna cosa abas d'insertar ***
'            If NumTabMto = 0 Then
'            End If
'            ' ******************************************************
'            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
'            ModoLineas = 0
'
'            If NumTabMto <> 3 Then
'                V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
'                CargaGrid NumTabMto, True
'            End If
'
'            ' *** si n'hi han tabs ***
'            SituarTab (NumTabMto)
'
'            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
'            If NumTabMto <> 3 Then
'                DataGridAux(NumTabMto).SetFocus
'                AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
'            End If
'            ' ***********************************************************
'
'            LLamaLineas NumTabMto, 0
'
'        End If
'    End If
'
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
Dim I As Integer, k As Integer
Dim Sql As String
Dim RS As ADODB.Recordset

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For I = 0 To Combo1.Count - 1
        Combo1(I).Clear
    Next I
    
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
    
    'tipo de recoleccion
    Combo1(1).AddItem "Cooperativa"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Socio"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
  
    'transportado por
    Combo1(2).AddItem "Cooperativa"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 0
    Combo1(2).AddItem "Socio"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 1

    For I = 0 To Combo15.Count - 1
        Combo15(I).Clear
    Next I
    
    'tipo de envases
    Sql = "select codtipen, nomtipen, pesocaja, escaja, numorden, 0 from confenva where not numorden is null "
    Sql = Sql & " union "
    Sql = Sql & " select codtipen, nomtipen, pesocaja, escaja, numorden, 1 from confenva where numorden is null "
    Sql = Sql & " order by 6, 5"

    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 1
    While Not RS.EOF
'        Sql = Replace(Rs.Fields(1).Value, "Factura", "Fac.")
        For k = 0 To 4
            Combo15(k).AddItem RS.Fields(1).Value 'campo del codigo
            Combo15(k).ItemData(Combo15(k).NewIndex) = RS.Fields(0).Value
        Next k
        
        I = I + 1
        RS.MoveNext
    Wend


End Sub


Private Sub txtAux1_LostFocus(Index As Integer)
'Dim cadMen As String
'Dim Nuevo As Boolean
'
'    If Not PerderFocoGnral(txtaux1(Index), Modo) Then Exit Sub
'
'    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
'    'mostrar mensajes ni hacer nada
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
'
'
'    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
'    Select Case Index
'        Case 1 ' seccion
'                If PonerFormatoEntero(txtaux1(Index)) Then
'                    Set vSeccion = New CSeccion
'                    If vSeccion.LeerDatos(txtaux1(Index)) Then
'                        txtAux2(Index).Text = vSeccion.Nombre
'                        b = vSeccion.AbrirConta
'                    Else
'                        Set vSeccion = Nothing
'                        cadMen = "No existe la Sección: " & txtaux1(Index).Text & vbCrLf
'                        cadMen = cadMen & "¿Desea crearla?" & vbCrLf
'                        If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
'                            Set frmSec = New frmManSeccion
'                            frmSec.DatosADevolverBusqueda = "0|1|"
'                            frmSec.NuevoCodigo = txtaux1(Index).Text
'                            txtaux1(Index).Text = ""
'                            TerminaBloquear
'                            frmSec.Show vbModal
'                            Set frmSec = Nothing
'                            If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
'                        Else
'                            txtaux1(Index).Text = ""
'                        End If
'                    End If
'                Else
'                    txtaux1(Index).Text = ""
'                End If
'
''                If PonerFormatoEntero(txtaux1(Index)) Then
''                txtAux2(Index).Text = PonerNombreDeCod(txtaux1(Index), "seccion", "nomsecci")
''                If txtAux2(Index).Text = "" Then
''                    cadMen = "No existe la Sección: " & txtaux1(Index).Text & vbCrLf
''                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
''                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
''                        Set frmSec = New frmManSeccion
''                        frmSec.DatosADevolverBusqueda = "0|1|"
''                        frmSec.NuevoCodigo = Text1(Index).Text
''                        txtaux1(Index).Text = ""
''                        TerminaBloquear
''                        frmSec.Show vbModal
''                        Set frmSec = Nothing
''                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
''                    Else
''                        txtaux1(Index).Text = ""
''                    End If
''                    PonerFoco txtaux1(Index)
''                End If
''            Else
''                txtAux2(Index).Text = ""
''            End If
'
'        Case 2, 3 'fecha de alta y de baja
'            PonerFormatoFecha txtaux1(Index)
'
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
'
'    End Select
'
'    ' ******************************************************************************
End Sub

'Private Sub txtAux1_GotFocus(Index As Integer)
'   If Not txtaux1(Index).MultiLine Then ConseguirFocoLin txtaux1(Index)
'End Sub
'
'Private Sub txtAux1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    If Not txtaux1(Index).MultiLine Then KEYdown KeyCode
'End Sub
'
'Private Sub txtAux1_KeyPress(Index As Integer, KeyAscii As Integer)
'    If Not txtaux1(Index).MultiLine Then
'        If KeyAscii = teclaBuscar Then
'            If Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) Then
'                Select Case Index
'                    Case 4: KEYBusqueda KeyAscii, 7 'pais
'                    Case 10: KEYBusqueda KeyAscii, 3 'mercado
'                    Case 11: KEYBusqueda KeyAscii, 4 'cadena
'                End Select
'            End If
'        Else
'            KEYpress KeyAscii
'        End If
'    End If
'End Sub



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
        .Show vbModal
    End With
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub

Private Sub CalcularTaras()
Dim Tara1 As Currency
Dim Tara2 As Currency
Dim Tara3 As Currency
Dim Tara4 As Currency
Dim Tara5 As Currency
Dim Tara11 As Currency
Dim Tara12 As Currency
Dim Tara13 As Currency
Dim Tara14 As Currency
Dim Tara15 As Currency
Dim PesoBruto As Currency
Dim PesoNeto As Currency
Dim PesoTrans As Currency
Dim TaraVehi As Currency
Dim OtrasTaras As Currency
Dim PesoCaja As Currency

    Tara1 = 0
    Tara2 = 0
    Tara3 = 0
    Tara4 = 0
    Tara5 = 0
    
    Tara11 = 0
    Tara12 = 0
    Tara13 = 0
    Tara14 = 0
    Tara15 = 0
    
    Text1(18).Text = ""
    Text1(19).Text = ""
    Text1(20).Text = ""
    Text1(8).Text = ""
    Text1(9).Text = ""
    
    'tara 1
    If Text1(13).Text <> "" Then
        PesoCaja = PesoEnvase(Combo15(0).ItemData(Combo15(0).ListIndex))
        Tara1 = Round2(CCur(ImporteSinFormato(Text1(13).Text)) * PesoCaja, 0)
        Tara11 = Round2(CCur(ImporteSinFormato(Text1(13).Text)) * PesoCaja, 0)
        Text1(18).Text = Tara1
        PonerFormatoEntero Text1(18)
    End If
    'tara 2
    If Text1(14).Text <> "" Then
        PesoCaja = PesoEnvase(Combo15(1).ItemData(Combo15(1).ListIndex))
        Tara2 = Round2(CCur(ImporteSinFormato(Text1(14).Text)) * PesoCaja, 0)
        Tara12 = Round2(CCur(ImporteSinFormato(Text1(14).Text)) * PesoCaja, 0)
        Text1(19).Text = Tara2
        PonerFormatoEntero Text1(19)
    End If
    'tara 3
    If Text1(15).Text <> "" Then
        PesoCaja = PesoEnvase(Combo15(2).ItemData(Combo15(2).ListIndex))
        Tara3 = Round2(CCur(ImporteSinFormato(Text1(15).Text)) * PesoCaja, 0)
        Tara13 = Round2(CCur(ImporteSinFormato(Text1(15).Text)) * PesoCaja, 0)
        Text1(20).Text = Tara3
        PonerFormatoEntero Text1(20)
    End If
    'tara 4
    If Text1(16).Text <> "" Then
        PesoCaja = PesoEnvase(Combo15(3).ItemData(Combo15(3).ListIndex))
        Tara4 = Round2(CCur(ImporteSinFormato(Text1(16).Text)) * PesoCaja, 0)
        Tara14 = Round2(CCur(ImporteSinFormato(Text1(16).Text)) * PesoCaja, 0)
        Text1(8).Text = Tara4
        PonerFormatoEntero Text1(8)
    End If
    'tara 5
    If Text1(17).Text <> "" Then
        PesoCaja = PesoEnvase(Combo15(4).ItemData(Combo15(4).ListIndex))
        Tara5 = Round2(CCur(ImporteSinFormato(Text1(17).Text)) * PesoCaja, 0)
        Tara15 = Round2(CCur(ImporteSinFormato(Text1(17).Text)) * PesoCaja, 0)
        Text1(9).Text = Tara5
        PonerFormatoEntero Text1(9)
    End If

    'peso neto
    PesoBruto = 0
    TaraVehi = 0
    OtrasTaras = 0
    If Text1(21).Text <> "" Then PesoBruto = CCur(Text1(21).Text)
    If Text1(3).Text <> "" Then TaraVehi = CCur(Text1(3).Text)
    If Text1(24).Text <> "" Then OtrasTaras = CCur(Text1(24).Text)
    
    PesoNeto = PesoBruto - Tara1 - Tara2 - Tara3 - Tara4 - Tara5 - TaraVehi - OtrasTaras
    PesoTrans = PesoBruto - Tara11 - Tara12 - Tara13 - Tara14 - Tara15 - TaraVehi - OtrasTaras
    Text1(11).Text = CStr(PesoNeto)
    Text1(23).Text = CStr(PesoTrans)
    PonerFormatoEntero Text1(11)
End Sub

Private Function PesoEnvase(CodEnvase As Integer) As Currency
Dim Sql As String

    Sql = "select pesocaja from confenva where codtipen = " & DBSet(CodEnvase, "N")
    
    PesoEnvase = DevuelveValor(Sql)

End Function


Private Sub PonerCamposSocioVariedad()
Dim Cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim RS As ADODB.Recordset


    If Text1(1).Text = "" Or Text1(2).Text = "" Then Exit Sub
    
    If Not (Modo = 3 Or Modo = 4) Then Exit Sub

    Cad = "rcampos.codsocio = " & DBSet(Text1(1).Text, "N") & " and rcampos.fecbajas is null"
    Cad = Cad & " and rcampos.codvarie = " & DBSet(Text1(2).Text, "N")
     
    Cad1 = "select count(*) from rcampos where " & Cad
     
    NumRegis = TotalRegistros(Cad1)
    
    If NumRegis = 0 Then Exit Sub
    If NumRegis = 1 Then
        Cad1 = "select codcampo from rcampos where " & Cad
        Set RS = New ADODB.Recordset
        RS.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            Text1(5).Text = DBLet(RS.Fields(0).Value)
            PonerDatosCampo Text1(5).Text
        End If
    Else
        Set frmMens = New frmMensajes
        frmMens.cadWhere = " and " & Cad '"rcampos.codsocio = " & NumCod & " and rcampos.fecbajas is null"
        frmMens.campo = Text1(5).Text
        frmMens.OpcionMensaje = 6
        frmMens.Show vbModal
        Set frmMens = Nothing
    End If
    
End Sub

Private Sub PonerDatosCampo(campo As String)
Dim Cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim RS As ADODB.Recordset


    If campo = "" Then Exit Sub
    
'    If Not (Modo = 3 Or Modo = 4) Then Exit Sub

    Cad = "rcampos.codcampo = " & DBSet(campo, "N") & " and rcampos.fecbajas is null"
     
    Cad1 = "select rcampos.codparti, rpartida.nomparti, rpartida.codzonas, rzonas.nomzonas, "
    Cad1 = Cad1 & " rpueblos.despobla, rcampos.nrocampo from rcampos, rpartida, rzonas, rpueblos "
    Cad1 = Cad1 & " where " & Cad
    Cad1 = Cad1 & " and rcampos.codparti = rpartida.codparti "
    Cad1 = Cad1 & " and rpartida.codzonas = rzonas.codzonas "
    Cad1 = Cad1 & " and rpartida.codpobla = rpueblos.codpobla "
     
    Set RS = New ADODB.Recordset
    RS.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
'    Text1(5).Text = ""
    Text2(0).Text = ""
    Text2(4).Text = ""
    Text2(3).Text = ""
    Text3(3).Text = ""
    Text4(3).Text = ""
    Text5(3).Text = ""
    If Not RS.EOF Then
        Text1(5).Text = campo
        PonerFormatoEntero Text1(5)
        Text2(4).Text = DBLet(RS.Fields(0).Value, "N") ' codigo de partida
        If Text2(4).Text <> "" Then Text2(4).Text = Format(Text2(4).Text, "0000")
        Text2(3).Text = DBLet(RS.Fields(1).Value, "T") ' nombre de partida
        Text3(3).Text = DBLet(RS.Fields(2).Value, "N") ' codigo de zona
        If Text3(3).Text <> "" Then Text3(3).Text = Format(Text3(3).Text, "0000")
        Text4(3).Text = DBLet(RS.Fields(3).Value, "T") ' nombre de zona
        Text5(3).Text = DBLet(RS.Fields(4).Value, "T") ' descripcion de poblacion
        Text2(0).Text = DBLet(RS.Fields(5).Value, "N") ' nro de campo
    End If
    
    Set RS = Nothing
    
End Sub

Private Sub PonerTarasVisibles()
    'tara1
    Text1(13).Enabled = (vParamAplic.TipoCaja1 <> "")
    Text1(13).visible = (vParamAplic.TipoCaja1 <> "")
    Text1(18).Enabled = (vParamAplic.TipoCaja1 <> "")
    Text1(18).visible = (vParamAplic.TipoCaja1 <> "")

    'tara2
    Text1(14).Enabled = (vParamAplic.TipoCaja2 <> "")
    Text1(14).visible = (vParamAplic.TipoCaja2 <> "")
    Text1(19).Enabled = (vParamAplic.TipoCaja2 <> "")
    Text1(19).visible = (vParamAplic.TipoCaja2 <> "")
    
    'tara3
    Text1(15).Enabled = (vParamAplic.TipoCaja3 <> "")
    Text1(15).visible = (vParamAplic.TipoCaja3 <> "")
    Text1(20).Enabled = (vParamAplic.TipoCaja3 <> "")
    Text1(20).visible = (vParamAplic.TipoCaja3 <> "")
    
    'tara4
    Text1(16).Enabled = (vParamAplic.TipoCaja4 <> "")
    Text1(16).visible = (vParamAplic.TipoCaja4 <> "")
    Text1(8).Enabled = (vParamAplic.TipoCaja4 <> "")
    Text1(8).visible = (vParamAplic.TipoCaja4 <> "")
    
    'tara5
    Text1(17).Enabled = (vParamAplic.TipoCaja5 <> "")
    Text1(17).visible = (vParamAplic.TipoCaja5 <> "")
    Text1(9).Enabled = (vParamAplic.TipoCaja5 <> "")
    Text1(9).visible = (vParamAplic.TipoCaja5 <> "")
End Sub


Private Function HanModificadoCajas() As Boolean
Dim I As Integer
    HanModificadoCajas = False
    
    NumCajas = 0
'    For I = 13 To 17
'        If Text1(I).Text <> "" Then
'            NumCajas = NumCajas + CCur(ComprobarCero(Text1(I).Text))
'        End If
'    Next I
    If Text1(13).Text <> "" And EsCaja(CStr(DBLet(Data1.Recordset!tipocajo1, "N"))) Then NumCajas = NumCajas + CCur(ComprobarCero(Text1(13).Text))
    If Text1(14).Text <> "" And EsCaja(CStr(DBLet(Data1.Recordset!tipocajo2, "N"))) Then NumCajas = NumCajas + CCur(ComprobarCero(Text1(14).Text))
    If Text1(15).Text <> "" And EsCaja(CStr(DBLet(Data1.Recordset!tipocajo3, "N"))) Then NumCajas = NumCajas + CCur(ComprobarCero(Text1(15).Text))
    If Text1(16).Text <> "" And EsCaja(CStr(DBLet(Data1.Recordset!tipocajo4, "N"))) Then NumCajas = NumCajas + CCur(ComprobarCero(Text1(16).Text))
    If Text1(17).Text <> "" And EsCaja(CStr(DBLet(Data1.Recordset!tipocajo5, "N"))) Then NumCajas = NumCajas + CCur(ComprobarCero(Text1(17).Text))

    HanModificadoCajas = (NumCajas <> NumCajasAnt)

End Function

Private Function HanModificadoKilos() As Boolean
Dim I As Integer
    
    HanModificadoKilos = (DBLet(Text1(11).Text, "N") <> NumKilosAnt)

End Function



Private Sub CrearPaletizacion()
Dim Sql As String

    Sql = "delete from trzpalets where numnotac = " & Trim(Data1.Recordset!numnotac)
    conn.Execute Sql
    
    mnPaletizacion_Click

End Sub


Private Sub ActualizarPaletizacion()
Dim RS As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String
Dim KilosTotal As Currency
Dim KilosNeto As Currency
Dim KilosLinea As Currency
Dim Numlineas As Currency
Dim IdPalet As Currency

    If vParamAplic.HayTraza = False Then Exit Sub
    
    Sql = "select numcajones, numkilos, idpalet from trzpalets where numnotac = " & Trim(Data1.Recordset!numnotac)
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenDynamic, adLockReadOnly, adCmdText
    
    If Not RS.EOF Then
        RS.MoveFirst
        
        KilosNeto = DBLet(Data1.Recordset!KilosNet, "N")
'        NumCajas = DBLet(Data1.Recordset!numcajo1, "N") + _
'                   DBLet(Data1.Recordset!numcajo2, "N") + _
'                   DBLet(Data1.Recordset!numcajo3, "N") + _
'                   DBLet(Data1.Recordset!numcajo4, "N") + _
'                   DBLet(Data1.Recordset!numcajo5, "N")

        NumCajas = 0
        If EsCaja(CStr(DBLet(Data1.Recordset!tipocajo1, "N"))) Then NumCajas = NumCajas + DBLet(Data1.Recordset!numcajo1, "N")
        If EsCaja(CStr(DBLet(Data1.Recordset!tipocajo2, "N"))) Then NumCajas = NumCajas + DBLet(Data1.Recordset!numcajo2, "N")
        If EsCaja(CStr(DBLet(Data1.Recordset!tipocajo3, "N"))) Then NumCajas = NumCajas + DBLet(Data1.Recordset!numcajo3, "N")
        If EsCaja(CStr(DBLet(Data1.Recordset!tipocajo4, "N"))) Then NumCajas = NumCajas + DBLet(Data1.Recordset!numcajo4, "N")
        If EsCaja(CStr(DBLet(Data1.Recordset!tipocajo5, "N"))) Then NumCajas = NumCajas + DBLet(Data1.Recordset!numcajo5, "N")
        
        If NumCajas = 0 Then 'vamos por palots y debemos ver cuantos registros=palots tenemos
            Sql1 = "select count(*) from trzpalets where numnotac = " & Trim(Data1.Recordset!numnotac)
            
            Numlineas = TotalRegistros(Sql1)
        End If
        
        KilosTotal = 0
        While Not RS.EOF
            If NumCajas <> 0 Then ' estamos por palet
                KilosLinea = (KilosNeto * DBLet(RS.Fields(0).Value, "N")) \ NumCajas
            Else ' estamos por palot
                KilosLinea = KilosNeto \ Numlineas
            End If
            
            Sql1 = "update trzpalets set numkilos = " & DBSet(KilosLinea, "N")
            Sql1 = Sql1 & " where idpalet = " & DBSet(RS.Fields(2).Value, "N")
            
            conn.Execute Sql1
            
            KilosTotal = KilosTotal + KilosLinea
        
            IdPalet = DBLet(RS.Fields(2).Value, "N")
            
            RS.MoveNext
        Wend
        
        If KilosTotal <> KilosNeto Then ' en el ultimo registro metemos el restante
            Sql1 = "update trzpalets set numkilos = numkilos + " & DBSet(KilosNeto - KilosTotal, "N")
            Sql1 = Sql1 & " where idpalet = " & DBSet(IdPalet, "N")
            
            conn.Execute Sql1
        End If
    End If
    
    Set RS = Nothing
        
End Sub

'***************************************
Private Sub ActivaTicket()
    ImpresoraDefecto = Printer.DeviceName
    XPDefaultPrinter vParamAplic.ImpresoraEntradas
End Sub

Private Sub DesactivaTicket()
    XPDefaultPrinter ImpresoraDefecto
End Sub


'---------------- Procesos para cambio de impresora por defecto ------------------
Private Sub XPDefaultPrinter(PrinterName As String)
    Dim Buffer As String
    Dim DeviceName As String
    Dim DriverName As String
    Dim PrinterPort As String
    Dim r As Long
    ' Get the printer information for the currently selected
    ' printer in the list. The information is taken from the
    ' WIN.INI file.
    Buffer = Space(1024)
    r = GetProfileString("PrinterPorts", PrinterName, "", _
        Buffer, Len(Buffer))

    ' Parse the driver name and port name out of the buffer
    GetDriverAndPort Buffer, DriverName, PrinterPort

       If DriverName <> "" And PrinterPort <> "" Then
           SetDefaultPrinter PrinterName, DriverName, PrinterPort
       End If
End Sub

Private Sub GetDriverAndPort(ByVal Buffer As String, DriverName As _
    String, PrinterPort As String)

    Dim iDriver As Integer
    Dim iPort As Integer
    DriverName = ""
    PrinterPort = ""

    ' The driver name is first in the string terminated by a comma
    iDriver = InStr(Buffer, ",")
    If iDriver > 0 Then

         ' Strip out the driver name
        DriverName = Left(Buffer, iDriver - 1)

        ' The port name is the second entry after the driver name
        ' separated by commas.
        iPort = InStr(iDriver + 1, Buffer, ",")

        If iPort > 0 Then
            ' Strip out the port name
            PrinterPort = Mid(Buffer, iDriver + 1, _
            iPort - iDriver - 1)
        End If
    End If
End Sub

Private Sub SetDefaultPrinter(ByVal PrinterName As String, _
    ByVal DriverName As String, ByVal PrinterPort As String)
    Dim DeviceLine As String
    Dim r As Long
    Dim L As Long
    DeviceLine = PrinterName & "," & DriverName & "," & PrinterPort
    ' Store the new printer information in the [WINDOWS] section of
    ' the WIN.INI file for the DEVICE= item
    r = WriteProfileString("windows", "Device", DeviceLine)
    ' Cause all applications to reload the INI file:
    L = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")
End Sub
'------------------ Fin de los procesos relacionados con el cambio de impresora ----



Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String
Dim MenError As String

    If Not vParamAplic.NroNotaManual Then

        Set vTipoMov = New CTiposMov
        If vTipoMov.Leer(CodTipoMov) Then
            Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
            Sql = CadenaInsertarDesdeForm(Me)
            If InsertarOferta(Sql, vTipoMov) Then
                CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
            End If
        End If
        Text1(0).Text = Format(Text1(0).Text, "0000000")
        
    Else ' el nro de nota es manual
        If InsertarDesdeForm2(Me, 1) Then
            MenError = "Actualizar chivato: "
            If Not ActualizarChivato(MenError, "I") Then
                MsgBox "Error Actualizando chivato" & vbCrLf & MenError
            End If
            CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
            PonerCadenaBusqueda
            PonerModo 2
        End If
    
    End If
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
        If ExisteNota(Text1(0).Text) Then
            devuelve = Text1(0).Text
        Else
            devuelve = ""
        End If
    
'        devuelve = DevuelveDesdeBDNew(cAgro, NombreTabla, "numnotac", "numnotac", Text1(0).Text, "N")
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
    MenError = "Error al insertar en la tabla (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    MenError = "Error al actualizar el contador."
    vTipoMov.IncrementarContador (CodTipoMov)
    
    MenError = "Actualizar chivato: "
    bol = ActualizarChivato(MenError, "I")
    
    
    
EInsertarOferta:
    If Err.Number <> 0 Or Not bol Then
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

Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim Sql As String

    On Error Resume Next
    
    Sql = " numnotac= " & Text1(0).Text
    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function


Private Sub ImprimirEntradaDirectaPrinter(NumNota As String, Copias As Integer)
    Dim NomImpre As String
  '  Dim FechaT As Date
    
    Dim RS As ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    Dim Rs4 As ADODB.Recordset
    Dim Sql As String
    Dim Lin As String ' línea de impresión
    Dim I As Integer
    Dim N As Integer
    Dim ImporteIva As Currency
    Dim EnEfectivo As Boolean
    
    Dim Veces As Integer
    Dim Partida As String
    Dim Situacion As String
    Dim Clase As String
    Dim Tara As Currency
    Dim cajas As Currency
    
On Error GoTo EImpTickD

    ActivaTicket
    
    Printer.Font = "Courier New"
    Printer.FontSize = 10
                        
'            Lin = "1234567890123456789012345678901234567890132456789012345678901234567890123456789012345678901234567890"
'            Printer.Print Lin
'            Printer.FontBold = True
'            Printer.Print Lin
'            Printer.FontBold = False
'            Printer.FontUnderline = True
'            Printer.Print Lin
'            Printer.FontUnderline = False
'            Printer.Print Lin
'            Printer.FontItalic = True
'            Printer.Print Lin
'            Printer.FontItalic = False
            
    
    '-- Obtenemos cabeceras y pies en un recordset (rs1)
    
    Sql = "select rentradas.*, rsocios.nomsocio, variedades.nomvarie from rentradas, rsocios, variedades "
    Sql = Sql & " where numnotac = " & DBSet(NumNota, "N")
    Sql = Sql & " and rentradas.codsocio = rsocios.codsocio "
    Sql = Sql & " and rentradas.codvarie = variedades.codvarie "
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly
    If Not RS.EOF Then
                '-- Impresión de la cabecera
'                Lin = "         1         2         3         4"
'                Printer.Print Lin
'                Lin = "1234567890123456789012345678901234567890"
'                Printer.Print Lin
    '    Lin = vEmpresa.nomempre
    
        Veces = Round2(CInt(Copias) / 2, 0)
    
        For I = 1 To Veces
    
            Printer.FontBold = True
            ' LINEA 1
            Lin = RellenaABlancos(vParam.NombreEmpresa, True, 43) & _
                  Space(2) & _
                  RellenaABlancos(vParam.NombreEmpresa, True, 43)
            Printer.Print Lin
            ' LINEA 2
            Lin = LineaCentrada("SECCION HORTOFRUTICOLA") & _
                  Space(2) & _
                  LineaCentrada("SECCION HORTOFRUTICOLA")
            Printer.Print Lin
            Printer.FontBold = False
            
            ' LINEA 3
            Lin = ""
            Printer.Print Lin
            
            ' LINEA 4
            Lin = "Fecha   : " & Format(RS!FechaEnt, "dd/mm/yyyy") & "  N." & Format(NumNota, "00000000") & " Hora:" & Format(RS!horaentr, "hh:mm") & _
                   Space(2) & _
                  "Fecha   : " & Format(RS!FechaEnt, "dd/mm/yyyy") & "  N." & Format(NumNota, "00000000") & " Hora:" & Format(RS!horaentr, "hh:mm")
'                  1234567890                         1234567890      1234                     56789012      345678                         90123
            Printer.Print Lin
            
            ' LINEA 5
            If I = 1 Or I = 2 Then
                Lin = "Socio   : " & RS!nomsocio
            Else
                Lin = ""
            End If
            Printer.Print Lin
            
            ' LINEA 6
            Partida = DevuelveValor("select nomparti from rcampos, rpartida where rcampos.codparti = rpartida.codparti and rcampos.codcampo = " & DBSet(RS!CodCampo, "N"))
            
            Lin = RellenaABlancos("Huerto  : " & Format(RS!CodCampo, "00000000") & "-" & Mid(Partida, 1, 24), True, 43) & _
                   Space(2) & _
                  RellenaABlancos("Huerto  : " & Format(RS!CodCampo, "00000000") & "-" & Mid(Partida, 1, 24), True, 43)
            Printer.Print Lin
'                  1234567890                         12345678      9    012345678901234567890123

            
            ' LINEA 7
            Situacion = ""
            Situacion = DevuelveValor("select nomsitua from rsituacioncampo, rcampos where rcampos.codsitua = rsituacioncampo.codsitua and rcampos.codsitua <> 0 and rcampos.codcampo = " & DBSet(RS!CodCampo, "N"))
            
            Lin = RellenaABlancos("Variedad: " & Format(RS!codvarie, "0000") & " " & DBLet(RS!nomvarie, "T") & " " & Situacion, True, 43) & _
                  Space(2) & _
                  RellenaABlancos("Variedad: " & Format(RS!codvarie, "0000") & " " & DBLet(RS!nomvarie, "T") & " " & Situacion, True, 43)
            Printer.Print Lin

            ' LINEA 8
            Clase = ""
            Clase = DevuelveValor("select nomclase from clases, variedades where variedades.codvarie = " & DBSet(RS!codvarie, "N") & " and variedades.codclase = clases.codclase ")
            
            Lin = RellenaABlancos("Grupo   : " & Clase, True, 43) & Space(2) & RellenaABlancos("Grupo   : " & Clase, True, 43) & Space(2)
            Printer.Print Lin
            
            ' LINEA 9
            Lin = ""
            Printer.Print Lin
            
            ' LINEA 10
'            Cajas = DBLet(Rs!numcajo1, "N") + DBLet(Rs!numcajo2, "N") + DBLet(Rs!numcajo3, "N") + DBLet(Rs!numcajo4, "N") + DBLet(Rs!numcajo5, "N")
            cajas = 0
            If EsCaja(CStr(DBLet(RS!tipocajo1, "N"))) Then cajas = cajas + DBLet(RS!numcajo1, "N")
            If EsCaja(CStr(DBLet(RS!tipocajo2, "N"))) Then cajas = cajas + DBLet(RS!numcajo2, "N")
            If EsCaja(CStr(DBLet(RS!tipocajo3, "N"))) Then cajas = cajas + DBLet(RS!numcajo3, "N")
            If EsCaja(CStr(DBLet(RS!tipocajo4, "N"))) Then cajas = cajas + DBLet(RS!numcajo4, "N")
            If EsCaja(CStr(DBLet(RS!tipocajo5, "N"))) Then cajas = cajas + DBLet(RS!numcajo5, "N")

            Tara = DBLet(RS!taracaja1, "N") + DBLet(RS!taracaja2, "N") + DBLet(RS!taracaja3, "N") + DBLet(RS!taracaja4, "N") + DBLet(RS!taracaja5, "N") + DBLet(RS!TaraVehi, "N")
            
            
            Lin = RellenaABlancos("Nro.Cajas : " & RellenaABlancos(Format(cajas, "###,##0"), False, 6) & "    " & "Total Tara: " & RellenaABlancos(Format(Tara, "###,##0"), False, 6), True, 43) & _
                  Space(2) & _
                  RellenaABlancos("Nro.Cajas : " & RellenaABlancos(Format(cajas, "###,##0"), False, 6) & "    " & "Total Tara: " & RellenaABlancos(Format(Tara, "###,##0"), False, 6), True, 43)

            Printer.Print Lin
            
            ' LINEA 11
            Lin = RellenaABlancos("Peso Bruto: " & RellenaABlancos(Format(RS!KilosBru, "###,##0"), False, 6) & "    " & "Peso Neto : " & RellenaABlancos(Format(RS!KilosNet, "###,##0"), False, 6), True, 43) & _
                  Space(2) & _
                  RellenaABlancos("Peso Bruto: " & RellenaABlancos(Format(RS!KilosBru, "###,##0"), False, 6) & "    " & "Peso Neto : " & RellenaABlancos(Format(RS!KilosNet, "###,##0"), False, 6), True, 43)
            Printer.Print Lin
            
            
            Lin = ""
            Printer.Print Lin
'            Printer.Print Lin
                
        Next I
                
    End If
    
    Printer.NewPage
    Printer.EndDoc
    
    DesactivaTicket
    
    Exit Sub
EImpTickD:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Imprimir ticket."
End Sub


Private Function LineaCentrada(Lin As String) As String
    Dim queda As Integer
    Dim Parte As Integer
    queda = 43 - Len(Lin)
    Parte = queda / 2
    If Parte Then
        LineaCentrada = String(Parte, " ") & Lin & String(queda - Parte, " ")
    Else
        LineaCentrada = Lin
    End If
End Function



Private Sub ImprimirEntradaDirectaLPT(NumNota As String, Copias As Integer)
    Dim NomImpre As String
  '  Dim FechaT As Date
    
    Dim RS As ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    Dim Rs4 As ADODB.Recordset
    Dim Sql As String
    Dim Lin As String ' línea de impresión
    Dim I As Integer
    Dim N As Integer
    Dim ImporteIva As Currency
    Dim EnEfectivo As Boolean
    
    Dim Veces As Integer
    Dim Partida As String
    Dim Situacion As String
    Dim Clase As String
    Dim Tara As Currency
    Dim cajas As Currency
    
    
On Error GoTo EImpTickD

    Set Lineas = New Collection
    
    If CargarLineas(NumNota, Copias) Then
        NF = FreeFile
'        Open "c:\t1.txt" For Output As #NF
        
        Open "LPT1" For Output As #NF
            
        For I = 1 To Lineas.Count
            Print #NF, Lineas(I)
        Next I
        
        Close (NF)
    End If
    
    Set Lineas = Nothing
    Exit Sub

EImpTickD:
    Set Lineas = Nothing
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "Imprimir Entrada."
End Sub


Private Function CargarLineas(NumNota As String, Copias As Integer) As Boolean
Dim Sql As String
Dim RS As ADODB.Recordset
Dim Lin As String ' línea de impresión
Dim I As Integer
Dim N As Integer
Dim ImporteIva As Currency
Dim EnEfectivo As Boolean
    
Dim Veces As Integer
Dim Partida As String
Dim Situacion As String
Dim Clase As String
Dim Tara As Currency
Dim cajas As Currency
Dim GGN As String
    
    On Error GoTo eCargarLineas
    
    CargarLineas = False
    
    
    Sql = "select rentradas.*, rsocios.nomsocio, variedades.nomvarie from rentradas, rsocios, variedades "
    Sql = Sql & " where numnotac = " & DBSet(NumNota, "N")
    Sql = Sql & " and rentradas.codsocio = rsocios.codsocio "
    Sql = Sql & " and rentradas.codvarie = variedades.codvarie "
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly
    If Not RS.EOF Then
    
        Veces = Round2(CInt(Copias) / 2, 0)
    
        For I = 1 To Veces
            ' LINEA 1
            Lin = RellenaABlancos(vParam.NombreEmpresa, True, 43) & _
                  Space(2) & _
                  RellenaABlancos(vParam.NombreEmpresa, True, 43)
            Lineas.Add Lin
            
            ' LINEA 2
            Lin = LineaCentrada("SECCION HORTOFRUTICOLA") & _
                  Space(2) & _
                  LineaCentrada("SECCION HORTOFRUTICOLA")
            Lineas.Add Lin
            
            
            ' LINEA 3
            Lin = ""
            Lineas.Add Lin
            
            ' LINEA 4
            Lin = "Fecha   : " & Format(RS!FechaEnt, "dd/mm/yyyy") & "  N." & Format(NumNota, "00000000") & " Hora:" & Format(RS!horaentr, "hh:mm") & _
                   Space(2) & _
                  "Fecha   : " & Format(RS!FechaEnt, "dd/mm/yyyy") & "  N." & Format(NumNota, "00000000") & " Hora:" & Format(RS!horaentr, "hh:mm")
'                  1234567890                         1234567890      1234                     56789012      345678                         90123
            Lineas.Add Lin
            
            ' LINEA 5
            If I = 1 Or I = 2 Then
                Lin = "Socio   : " & RS!nomsocio
            Else
                Lin = ""
            End If
            Lineas.Add Lin
            
            ' LINEA 6
            Partida = DevuelveValor("select nomparti from rcampos, rpartida where rcampos.codparti = rpartida.codparti and rcampos.codcampo = " & DBSet(RS!CodCampo, "N"))
            
            '[Monica]27/04/2012: Añadimos el GGN si lo tiene
            GGN = CStr(DevuelveValor("select numeroggn from rcampos, rglobalgap where rcampos.codcampo = " & DBSet(RS!CodCampo, "N") & " and rcampos.codigoggap = rglobalgap.codigo "))
            If GGN <> "0" Then
                Lin = RellenaABlancos("Huerto  : " & Format(RS!CodCampo, "00000000") & "-" & RellenaABlancos(RellenaABlancos(Mid(Partida, 1, 23 - Len(GGN)), True, 23 - Len(GGN)) & " " & GGN, True, 24), True, 43) & _
                       Space(2) & _
                      RellenaABlancos("Huerto  : " & Format(RS!CodCampo, "00000000") & "-" & RellenaABlancos(RellenaABlancos(Mid(Partida, 1, 23 - Len(GGN)), True, 23 - Len(GGN)) & " " & GGN, True, 24), True, 43)
    '                  1234567890                         12345678      9    012345678901234567890123
                Lineas.Add Lin
            Else
                Lin = RellenaABlancos("Huerto  : " & Format(RS!CodCampo, "00000000") & "-" & Mid(Partida, 1, 24), True, 43) & _
                       Space(2) & _
                      RellenaABlancos("Huerto  : " & Format(RS!CodCampo, "00000000") & "-" & Mid(Partida, 1, 24), True, 43)
    '                  1234567890                         12345678      9    012345678901234567890123
                Lineas.Add Lin
            End If
            
            ' LINEA 7
            Situacion = ""
            Situacion = DevuelveValor("select nomsitua from rsituacioncampo, rcampos where rcampos.codsitua = rsituacioncampo.codsitua and rcampos.codsitua <> 0 and rcampos.codcampo = " & DBSet(RS!CodCampo, "N"))
            
            Lin = RellenaABlancos("Variedad: " & Format(RS!codvarie, "0000") & " " & DBLet(RS!nomvarie, "T") & " " & Situacion, True, 43) & _
                  Space(2) & _
                  RellenaABlancos("Variedad: " & Format(RS!codvarie, "0000") & " " & DBLet(RS!nomvarie, "T") & " " & Situacion, True, 43)
            Lineas.Add Lin

            ' LINEA 8
            Clase = ""
            Clase = DevuelveValor("select nomclase from clases, variedades where variedades.codvarie = " & DBSet(RS!codvarie, "N") & " and variedades.codclase = clases.codclase ")
            
            Lin = RellenaABlancos("Grupo   : " & Clase, True, 43) & Space(2) & RellenaABlancos("Grupo   : " & Clase, True, 43) & Space(2)
            Lineas.Add Lin
            
            ' LINEA 9
            Lin = ""
            Lineas.Add Lin
            
            ' LINEA 10
            'Cajas = DBLet(Rs!numcajo1, "N") + DBLet(Rs!numcajo2, "N") + DBLet(Rs!numcajo3, "N") + DBLet(Rs!numcajo4, "N") + DBLet(Rs!numcajo5, "N")
            cajas = 0
            If EsCaja(CStr(DBLet(RS!tipocajo1, "N"))) Then cajas = cajas + DBLet(RS!numcajo1, "N")
            If EsCaja(CStr(DBLet(RS!tipocajo2, "N"))) Then cajas = cajas + DBLet(RS!numcajo2, "N")
            If EsCaja(CStr(DBLet(RS!tipocajo3, "N"))) Then cajas = cajas + DBLet(RS!numcajo3, "N")
            If EsCaja(CStr(DBLet(RS!tipocajo4, "N"))) Then cajas = cajas + DBLet(RS!numcajo4, "N")
            If EsCaja(CStr(DBLet(RS!tipocajo5, "N"))) Then cajas = cajas + DBLet(RS!numcajo5, "N")
            
            Tara = DBLet(RS!taracaja1, "N") + DBLet(RS!taracaja2, "N") + DBLet(RS!taracaja3, "N") + DBLet(RS!taracaja4, "N") + DBLet(RS!taracaja5, "N") + DBLet(RS!TaraVehi, "N")
            
            Lin = RellenaABlancos("Nro.Cajas : " & RellenaABlancos(Format(cajas, "###,##0"), False, 6) & "    " & "Total Tara: " & RellenaABlancos(Format(Tara, "###,##0"), False, 6), True, 43) & _
                  Space(2) & _
                  RellenaABlancos("Nro.Cajas : " & RellenaABlancos(Format(cajas, "###,##0"), False, 6) & "    " & "Total Tara: " & RellenaABlancos(Format(Tara, "###,##0"), False, 6), True, 43)
            Lineas.Add Lin

            
            ' LINEA 11
            Lin = RellenaABlancos("Peso Bruto: " & RellenaABlancos(Format(RS!KilosBru, "###,##0"), False, 6) & "    " & "Peso Neto : " & RellenaABlancos(Format(RS!KilosNet, "###,##0"), False, 6), True, 43) & _
                  Space(2) & _
                  RellenaABlancos("Peso Bruto: " & RellenaABlancos(Format(RS!KilosBru, "###,##0"), False, 6) & "    " & "Peso Neto : " & RellenaABlancos(Format(RS!KilosNet, "###,##0"), False, 6), True, 43)
            Lineas.Add Lin
            
            
            Lin = ""
            Lineas.Add Lin
'            Lineas.Add Lin
                
        Next I
    End If

    CargarLineas = True
    Exit Function
    
eCargarLineas:
    MuestraError Err.Number, "Cargando las lineas de impresión:", Err.Description
End Function




Private Function ActualizarChivato(Mens As String, Operacion As String) As Boolean
Dim Sql As String
Dim RS As ADODB.Recordset
Dim Sql2 As String
Dim RS1 As ADODB.Recordset
Dim Cadena As String
Dim Producto As String
Dim NumF As String

    On Error GoTo eActualizarChivato

    ActualizarChivato = False
    
    Sql = "select codvarie, numcajo1, numnotac, codsocio, codcampo, codcapat, codtarif, "
    Sql = Sql & "kilosbru, kilosnet, tipoentr, fechaent, codtrans, nropesada "
    Sql = Sql & "from rentradas"
    Sql = Sql & " where numnotac = " & DBSet(Text1(0).Text, "N")
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    If Not RS.EOF Then
        Producto = DevuelveValor("select codprodu from variedades where codvarie = " & DBSet(RS!codvarie, "N"))
        
        Cadena = v_cadena & "<ROW notacamp=" & """" & Format(DBLet(RS!numnotac, "N"), "######0") & """"
        Cadena = Cadena & " fechaent=" & """" & Format(RS!FechaEnt, "yyyymmdd") & """"
        Cadena = Cadena & " codprodu=" & """" & Format(DBLet(Producto, "N"), "#####0") & """"
        Cadena = Cadena & " codvarie=" & """" & Format(DBLet(RS!codvarie, "N"), "#####0") & """"
        Cadena = Cadena & " codsocio=" & """" & Format(DBLet(RS!Codsocio, "N"), "#####0") & """"
        Cadena = Cadena & " codcampo=" & """" & Format(DBLet(RS!CodCampo, "N"), "#######0") & """"
        Cadena = Cadena & " kilosbru=" & """" & Format(DBLet(RS!KilosBru, "N"), "###0") & """"
        Cadena = Cadena & " kilosnet=" & """" & Format(DBLet(RS!KilosNet, "N"), "###0") & """"
        Cadena = Cadena & " numcajo1=" & """" & Format(DBLet(RS!numcajo1, "N"), "##0") & """"
        Cadena = Cadena & " numcajo2=" & """" & Format(0, "##0") & """"
        Cadena = Cadena & " numcajo3=" & """" & Format(0, "##0") & """"
        Cadena = Cadena & " numcajo4=" & """" & Format(0, "##0") & """"
        Cadena = Cadena & " numcajo5=" & """" & Format(0, "##0") & """"
        Cadena = Cadena & " matricul=" & """" & DBLet(RS!codTrans, "T") & """"
        Cadena = Cadena & " codcapat=" & """" & Format(DBLet(RS!codcapat, "N"), "###0") & """"
        Cadena = Cadena & " identifi=" & """" & Format(0, "#####0") & """"
        Cadena = Cadena & " altura=" & """" & Format(vParamAplic.CajasporPalet, "##0") & """"
        Cadena = Cadena & " zona=" & """" & Format(0, "#########0") & """"
        Cadena = Cadena & " /></ROWDATA></DATAPACKET>"
    
            
        NumF = DevuelveValor("select max(numorden) + 1 from chivato")
        
        
        Sql = "insert into chivato (numorden, basedato, nomtabla, operacio, fechadia, separado,"
        Sql = Sql & "claveant, clavenue, nombmemo, nombmem1, nombmem2, horaproc, nombmem3, nombmem4) values ("
        Sql = Sql & DBSet(NumF, "N") & ","
        Sql = Sql & "'agro',"
        Sql = Sql & "'sentba',"
        
        Select Case Operacion
            Case "I" ' insertada
                Sql = Sql & "'I',"
            Case "U" ' actualizada
                Sql = Sql & "'U',"
            Case "Z" ' borrada
                Sql = Sql & "'D',"
        End Select
        
        Sql = Sql & DBSet(Now, "F") & ","
        Sql = Sql & DBSet("&", "T") & ","
        Sql = Sql & DBSet(RS!numnotac, "N") & ","
        Sql = Sql & DBSet(RS!numnotac, "N") & ","
        Sql = Sql & DBSet(Cadena, "T") & ","
        Sql = Sql & ValorNulo & ","
        Sql = Sql & ValorNulo & ","
        Sql = Sql & "'" & Format(Now, "hh:mm:ss") & "',"
        Sql = Sql & ValorNulo & ","
        Sql = Sql & ValorNulo & ")"
        
        conn.Execute Sql
            
    End If
    
    Set RS = Nothing
    
    ActualizarChivato = True
    Exit Function
    
eActualizarChivato:
    Mens = Mens & Err.Description
End Function



