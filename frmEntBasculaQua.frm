VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEntBasculaQua 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entrada en b�scula"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   15405
   Icon            =   "frmEntBasculaQua.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   15405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Cuenta Principal|N|N|0|1|cltebanc|ctaprpal|0||"
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   180
      TabIndex        =   127
      Top             =   90
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   128
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
      Left            =   3840
      TabIndex        =   125
      Top             =   90
      Width           =   1305
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   210
         TabIndex        =   126
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
               Object.ToolTipText     =   "Tara Tractor"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Paletizaci�n"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   5220
      TabIndex        =   123
      Top             =   90
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   124
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
      Index           =   0
      Left            =   12450
      TabIndex        =   122
      Top             =   240
      Width           =   1605
   End
   Begin VB.Frame Frame2 
      Height          =   825
      Index           =   0
      Left            =   180
      TabIndex        =   44
      Top             =   840
      Width           =   15075
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
         Left            =   6660
         MaxLength       =   10
         TabIndex        =   2
         Top             =   270
         Width           =   1200
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   6660
         MaxLength       =   20
         TabIndex        =   64
         Tag             =   "Hora|FH|N|||rentradas|horaentr|yyyy-mm-dd hh:mm:ss||"
         Top             =   360
         Width           =   1200
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
         Left            =   4380
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Entrada|F|N|||rentradas|fechaent|dd/mm/yyyy||"
         Top             =   270
         Width           =   1230
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
         Left            =   1410
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Numero de Nota|N|S|1|9999999|rentradas|numnotac|0000000|S|"
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label Label26 
         Caption         =   "Hora"
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
         Left            =   6075
         TabIndex        =   65
         Top             =   300
         Width           =   570
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   4095
         Picture         =   "frmEntBasculaQua.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   300
         Width           =   240
      End
      Begin VB.Label Label18 
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
         Height          =   255
         Left            =   3315
         TabIndex        =   63
         Top             =   300
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "N� Nota"
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
         TabIndex        =   45
         Top             =   300
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   210
      TabIndex        =   41
      Top             =   6825
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
         TabIndex        =   43
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
      Left            =   14205
      TabIndex        =   42
      Top             =   6930
      Width           =   1095
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
      Left            =   12945
      TabIndex        =   40
      Top             =   6930
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   750
      Top             =   6585
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
      Left            =   14205
      TabIndex        =   46
      Top             =   6930
      Visible         =   0   'False
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4980
      Left            =   210
      TabIndex        =   48
      Top             =   1725
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   8784
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   6
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos entrada"
      TabPicture(0)   =   "frmEntBasculaQua.frx":0097
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
      Tab(0).Control(20)=   "Label37"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label33"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "FrameDatosDtoAdministracion"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text5(3)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text4(3)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text3(3)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text1(2)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text2(2)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Combo1(0)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text1(1)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text2(1)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text1(5)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text2(3)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text2(4)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Combo1(1)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text2(12)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text1(12)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text2(6)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text1(6)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Text2(7)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Text1(7)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Combo1(2)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Text2(0)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Text2(5)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Combo1(3)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).ControlCount=   45
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
         Left            =   5010
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Tag             =   "Ausencia Plagas|N|N|0|1|rentradas|ausenciaplagas||N|"
         Top             =   4440
         Width           =   1650
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
         Index           =   5
         Left            =   3390
         MaxLength       =   20
         TabIndex        =   120
         Top             =   1320
         Width           =   1515
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         Left            =   5670
         MaxLength       =   4
         TabIndex        =   97
         Top             =   1320
         Width           =   1035
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
         Left            =   3420
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Tag             =   "Transportado por|N|N|0|1|rentradas|transportadopor||N|"
         Top             =   4440
         Width           =   1560
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
         Index           =   7
         Left            =   1380
         MaxLength       =   3
         TabIndex        =   8
         Tag             =   "C�digo Tarifa|N|S|0|999|rentradas|codtarif|000||"
         Top             =   3750
         Width           =   795
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
         Index           =   7
         Left            =   2205
         TabIndex        =   74
         Top             =   3750
         Width           =   4500
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
         Left            =   1380
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "C�digo Transporte|T1|S|||rentradas|codtrans|||"
         Top             =   3330
         Width           =   1245
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
         Index           =   6
         Left            =   2640
         TabIndex        =   72
         Top             =   3330
         Width           =   4080
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
         Index           =   12
         Left            =   1380
         MaxLength       =   4
         TabIndex        =   6
         Tag             =   "C�digo Capataz|N|S|0|9999|rentradas|codcapat|0000||"
         Top             =   2940
         Width           =   795
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
         Index           =   12
         Left            =   2205
         TabIndex        =   70
         Top             =   2940
         Width           =   4500
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
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Tag             =   "Recolectado|N|N|0|1|rentradas|recolect||N|"
         Top             =   4440
         Width           =   1500
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         Left            =   1380
         MaxLength       =   4
         TabIndex        =   68
         Top             =   1740
         Width           =   885
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
         Left            =   2310
         MaxLength       =   30
         TabIndex        =   67
         Top             =   1740
         Width           =   4395
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
         Left            =   1380
         MaxLength       =   8
         TabIndex        =   5
         Tag             =   "C�digo Campo|N|N|1|99999999|rentradas|codcampo|00000000|N|"
         Top             =   1320
         Width           =   1095
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
         Index           =   1
         Left            =   2310
         MaxLength       =   40
         TabIndex        =   61
         Top             =   930
         Width           =   4395
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
         Left            =   1380
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "C�digo Socio|N|N|1|999999|rentradas|codsocio|000000|N|"
         Top             =   930
         Width           =   900
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
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Tag             =   "Tipo Entrada|N|N|0|3|rentradas|tipoentr||N|"
         Top             =   4440
         Width           =   1710
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
         Left            =   2310
         MaxLength       =   30
         TabIndex        =   57
         Top             =   520
         Width           =   4395
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
         Left            =   1380
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Variedad|N|N|1|9999|rentradas|codvarie|0000||"
         Top             =   520
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
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
         Left            =   1380
         MaxLength       =   6
         TabIndex        =   52
         Top             =   2535
         Width           =   885
      End
      Begin VB.TextBox Text4 
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
         Left            =   2295
         MaxLength       =   30
         TabIndex        =   51
         Top             =   2535
         Width           =   4425
      End
      Begin VB.TextBox Text5 
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
         Left            =   1380
         MaxLength       =   30
         TabIndex        =   50
         Top             =   2130
         Width           =   5340
      End
      Begin VB.Frame FrameDatosDtoAdministracion 
         Caption         =   "Pesos y Taras"
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
         Height          =   4365
         Left            =   6840
         TabIndex        =   49
         Top             =   405
         Width           =   8100
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
            Index           =   35
            Left            =   6660
            MaxLength       =   13
            TabIndex        =   15
            Tag             =   "Zona|N|S|||rentradas|zona|#,###,###,##0||"
            Top             =   315
            Width           =   1245
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
            Index           =   23
            Left            =   5070
            MaxLength       =   3
            TabIndex        =   14
            Tag             =   "Altura|N|S|||rentradas|altura|##0||"
            Top             =   315
            Width           =   825
         End
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Caption         =   "Salidas"
            Height          =   3435
            Left            =   5070
            TabIndex        =   99
            Top             =   765
            Width           =   2985
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
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
               Index           =   3
               Left            =   1965
               MaxLength       =   7
               TabIndex        =   37
               Tag             =   "Tara Vehiculo|N|S|0|999999|rentradas|taravehisa|###,##0||"
               Top             =   2640
               Width           =   885
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
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
               Index           =   24
               Left            =   1965
               MaxLength       =   7
               TabIndex        =   39
               Tag             =   "Otras Taras|N|S|0|999999|rentradas|otrastarasa|###,##0||"
               Top             =   3060
               Width           =   885
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
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
               Index           =   29
               Left            =   90
               MaxLength       =   5
               TabIndex        =   31
               Tag             =   "Nro.Cajas 5|N|S|||rentradas|numcajosa5|#,##0||"
               Top             =   2130
               Width           =   765
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
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
               Index           =   28
               Left            =   90
               MaxLength       =   5
               TabIndex        =   30
               Tag             =   "Nro.Cajas 4|N|S|||rentradas|numcajosa4|#,##0||"
               Top             =   1740
               Width           =   765
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
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
               Index           =   27
               Left            =   90
               MaxLength       =   5
               TabIndex        =   29
               Tag             =   "Nro.Cajas 3|N|S|||rentradas|numcajosa3|#,##0||"
               Top             =   1350
               Width           =   765
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
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
               Index           =   26
               Left            =   90
               MaxLength       =   5
               TabIndex        =   28
               Tag             =   "Nro.Cajas 2|N|S|||rentradas|numcajosa2|#,##0||"
               Top             =   960
               Width           =   765
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
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
               Index           =   34
               Left            =   1950
               MaxLength       =   7
               TabIndex        =   36
               Tag             =   "Tara 5|N|S|0|999999|rentradas|taracajasa5|###,##0||"
               Top             =   2130
               Width           =   885
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
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
               Index           =   33
               Left            =   1950
               MaxLength       =   7
               TabIndex        =   35
               Tag             =   "Tara 4|N|S|0|999999|rentradas|taracajasa4|###,##0||"
               Top             =   1740
               Width           =   885
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
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
               Index           =   32
               Left            =   1950
               MaxLength       =   7
               TabIndex        =   34
               Tag             =   "Tara 3|N|S|0|999999|rentradas|taracajasa3|###,##0||"
               Top             =   1350
               Width           =   885
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
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
               Index           =   31
               Left            =   1950
               MaxLength       =   7
               TabIndex        =   33
               Tag             =   "Tara 2|N|S|0|999999|rentradas|taracajasa2|###,##0||"
               Top             =   960
               Width           =   885
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
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
               Index           =   30
               Left            =   1950
               MaxLength       =   7
               TabIndex        =   32
               Tag             =   "Tara 1|N|S|0|999999|rentradas|taracajasa1|###,##0||"
               Top             =   570
               Width           =   885
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
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
               Index           =   25
               Left            =   90
               MaxLength       =   5
               TabIndex        =   27
               Tag             =   "Nro.Cajas 1|N|S|||rentradas|numcajosa1|#,##0||"
               Top             =   570
               Width           =   765
            End
            Begin VB.Label Label10 
               Caption         =   "= "
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
               Index           =   9
               Left            =   1770
               TabIndex        =   116
               Top             =   2175
               Width           =   150
            End
            Begin VB.Label Label10 
               Caption         =   "= "
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
               Left            =   1770
               TabIndex        =   115
               Top             =   1785
               Width           =   150
            End
            Begin VB.Label Label10 
               Caption         =   "= "
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
               Left            =   1770
               TabIndex        =   114
               Top             =   1395
               Width           =   150
            End
            Begin VB.Label Label10 
               Caption         =   "= "
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
               Index           =   6
               Left            =   1770
               TabIndex        =   113
               Top             =   1005
               Width           =   150
            End
            Begin VB.Label Label10 
               Caption         =   "= "
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
               Index           =   5
               Left            =   1770
               TabIndex        =   112
               Top             =   615
               Width           =   150
            End
            Begin VB.Label Label8 
               Caption         =   "Tara Veh�culo"
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
               Left            =   900
               TabIndex        =   111
               Top             =   2655
               Width           =   1155
            End
            Begin VB.Label Label21 
               Caption         =   "Otras Taras"
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
               Left            =   900
               TabIndex        =   110
               Top             =   3075
               Width           =   1155
            End
            Begin VB.Label Label30 
               Caption         =   "Salidas"
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
               Height          =   255
               Left            =   90
               TabIndex        =   109
               Top             =   -30
               Width           =   1185
            End
            Begin VB.Label Label19 
               Caption         =   "x  Peso 1"
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
               Index           =   9
               Left            =   990
               TabIndex        =   107
               Top             =   2160
               Width           =   675
            End
            Begin VB.Label Label19 
               Caption         =   "x  Peso 1"
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
               Left            =   990
               TabIndex        =   106
               Top             =   1770
               Width           =   675
            End
            Begin VB.Label Label19 
               Caption         =   "x  Peso 1"
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
               Left            =   990
               TabIndex        =   105
               Top             =   1380
               Width           =   675
            End
            Begin VB.Label Label19 
               Caption         =   "x  Peso 1"
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
               Index           =   6
               Left            =   990
               TabIndex        =   104
               Top             =   990
               Width           =   675
            End
            Begin VB.Label Label19 
               Caption         =   "x  Peso 1  "
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
               Index           =   5
               Left            =   990
               TabIndex        =   103
               Top             =   600
               Width           =   675
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00972E0B&
               X1              =   90
               X2              =   2850
               Y1              =   240
               Y2              =   240
            End
            Begin VB.Label Label27 
               Caption         =   "Cajas"
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
               Left            =   90
               TabIndex        =   102
               Top             =   300
               Width           =   705
            End
            Begin VB.Label Label25 
               Caption         =   "Peso Caja"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   990
               TabIndex        =   101
               Top             =   300
               Width           =   945
            End
            Begin VB.Label Label24 
               Caption         =   "Tara"
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
               Left            =   1950
               TabIndex        =   100
               Top             =   300
               Width           =   765
            End
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
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
            Index           =   8
            Left            =   2160
            MaxLength       =   5
            TabIndex        =   38
            Top             =   3420
            Width           =   1335
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Caption         =   "Entradas"
            Height          =   2670
            Left            =   135
            TabIndex        =   76
            Top             =   735
            Width           =   4905
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
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
               Index           =   13
               Left            =   2025
               MaxLength       =   5
               TabIndex        =   16
               Tag             =   "Nro.Cajas 1|N|S|||rentradas|numcajo1|#,##0||"
               Top             =   600
               Width           =   765
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
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
               Index           =   18
               Left            =   3885
               MaxLength       =   7
               TabIndex        =   22
               Tag             =   "Tara 1|N|S|0|999999|rentradas|taracaja1|###,##0||"
               Top             =   600
               Width           =   885
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
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
               Index           =   19
               Left            =   3885
               MaxLength       =   7
               TabIndex        =   23
               Tag             =   "Tara 2|N|S|0|999999|rentradas|taracaja2|###,##0||"
               Top             =   990
               Width           =   885
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
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
               Index           =   20
               Left            =   3885
               MaxLength       =   7
               TabIndex        =   24
               Tag             =   "Tara 3|N|S|0|999999|rentradas|taracaja3|###,##0||"
               Top             =   1380
               Width           =   885
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
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
               Index           =   8
               Left            =   3885
               MaxLength       =   7
               TabIndex        =   25
               Tag             =   "Tara 4|N|S|0|999999|rentradas|taracaja4|###,##0||"
               Top             =   1770
               Width           =   885
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
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
               Index           =   9
               Left            =   3885
               MaxLength       =   7
               TabIndex        =   26
               Tag             =   "Tara 5|N|S|0|999999|rentradas|taracaja5|###,##0||"
               Top             =   2160
               Width           =   885
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
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
               Index           =   14
               Left            =   2025
               MaxLength       =   5
               TabIndex        =   17
               Tag             =   "Nro.Cajas 2|N|S|||rentradas|numcajo2|#,##0||"
               Top             =   990
               Width           =   765
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
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
               Index           =   15
               Left            =   2025
               MaxLength       =   5
               TabIndex        =   18
               Tag             =   "Nro.Cajas 3|N|S|||rentradas|numcajo3|#,##0||"
               Top             =   1380
               Width           =   765
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
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
               Index           =   16
               Left            =   2025
               MaxLength       =   5
               TabIndex        =   19
               Tag             =   "Nro.Cajas 4|N|S|||rentradas|numcajo4|#,##0||"
               Top             =   1770
               Width           =   765
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
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
               Index           =   17
               Left            =   2025
               MaxLength       =   5
               TabIndex        =   20
               Tag             =   "Nro.Cajas 5|N|S|||rentradas|numcajo5|#,##0||"
               Top             =   2160
               Width           =   765
            End
            Begin VB.Label Label29 
               Caption         =   "Entradas"
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
               Height          =   255
               Left            =   150
               TabIndex        =   108
               Top             =   0
               Width           =   1185
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00972E0B&
               X1              =   150
               X2              =   4740
               Y1              =   270
               Y2              =   270
            End
            Begin VB.Label Label16 
               Caption         =   "Tara"
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
               Left            =   3870
               TabIndex        =   94
               Top             =   330
               Width           =   975
            End
            Begin VB.Label Label14 
               Caption         =   "Peso Caja"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   2880
               TabIndex        =   93
               Top             =   330
               Width           =   945
            End
            Begin VB.Label Label19 
               Caption         =   "x  Peso 1  "
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
               Left            =   2925
               TabIndex        =   92
               Top             =   630
               Width           =   675
            End
            Begin VB.Label Label15 
               Caption         =   "Tarifa"
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
               Left            =   135
               TabIndex        =   91
               Top             =   630
               Width           =   1830
            End
            Begin VB.Label Label15 
               Caption         =   "Tarifa"
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
               Left            =   135
               TabIndex        =   90
               Top             =   1020
               Width           =   1830
            End
            Begin VB.Label Label15 
               Caption         =   "Tarifa"
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
               Left            =   135
               TabIndex        =   89
               Top             =   1410
               Width           =   1830
            End
            Begin VB.Label Label15 
               Caption         =   "Tarifa"
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
               Left            =   135
               TabIndex        =   88
               Top             =   1800
               Width           =   1830
            End
            Begin VB.Label Label15 
               Caption         =   "Tarifa"
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
               Left            =   135
               TabIndex        =   87
               Top             =   2190
               Width           =   1830
            End
            Begin VB.Label Label13 
               Caption         =   "Cajas"
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
               Left            =   2025
               TabIndex        =   86
               Top             =   330
               Width           =   1185
            End
            Begin VB.Label Label19 
               Caption         =   "x  Peso 1"
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
               Left            =   2925
               TabIndex        =   85
               Top             =   1020
               Width           =   675
            End
            Begin VB.Label Label19 
               Caption         =   "x  Peso 1"
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
               Left            =   2925
               TabIndex        =   84
               Top             =   1410
               Width           =   675
            End
            Begin VB.Label Label19 
               Caption         =   "x  Peso 1"
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
               Left            =   2925
               TabIndex        =   83
               Top             =   1800
               Width           =   675
            End
            Begin VB.Label Label19 
               Caption         =   "x  Peso 1"
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
               Left            =   2925
               TabIndex        =   82
               Top             =   2190
               Width           =   675
            End
            Begin VB.Label Label10 
               Caption         =   "= "
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
               Left            =   3675
               TabIndex        =   81
               Top             =   630
               Width           =   150
            End
            Begin VB.Label Label10 
               Caption         =   "= "
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
               Left            =   3675
               TabIndex        =   80
               Top             =   1020
               Width           =   150
            End
            Begin VB.Label Label10 
               Caption         =   "= "
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
               Left            =   3675
               TabIndex        =   79
               Top             =   1410
               Width           =   150
            End
            Begin VB.Label Label10 
               Caption         =   "= "
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
               Left            =   3675
               TabIndex        =   78
               Top             =   1800
               Width           =   150
            End
            Begin VB.Label Label10 
               Caption         =   "= "
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
               Left            =   3675
               TabIndex        =   77
               Top             =   2190
               Width           =   150
            End
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
            Index           =   11
            Left            =   2145
            MaxLength       =   7
            TabIndex        =   21
            Tag             =   "Peso Neto|N|N|0|999999|rentradas|kilosnet|###,##0||"
            Top             =   3840
            Width           =   1335
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
            Index           =   21
            Left            =   2160
            MaxLength       =   7
            TabIndex        =   13
            Tag             =   "Peso Bruto|N|N|||rentradas|kilosbru|###,##0||"
            Top             =   330
            Width           =   1365
         End
         Begin VB.Label Label32 
            Caption         =   "Zona"
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
            Left            =   6030
            TabIndex        =   118
            Top             =   345
            Width           =   585
         End
         Begin VB.Label Label31 
            Caption         =   "Altura"
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
            Left            =   4020
            TabIndex        =   117
            Top             =   345
            Width           =   855
         End
         Begin VB.Label Label15 
            Caption         =   "Mermas"
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
            Index           =   5
            Left            =   300
            TabIndex        =   96
            Top             =   3450
            Width           =   1830
         End
         Begin VB.Label Label17 
            Caption         =   "Peso Bruto"
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
            Left            =   255
            TabIndex        =   60
            Top             =   375
            Width           =   1185
         End
         Begin VB.Label Label7 
            Caption         =   "Peso Neto"
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
            Left            =   300
            TabIndex        =   59
            Top             =   3840
            Width           =   1155
         End
      End
      Begin VB.Label Label33 
         Caption         =   "Ausencia Plagas"
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
         Left            =   5040
         TabIndex        =   121
         Top             =   4140
         Width           =   1710
      End
      Begin VB.Label Label37 
         Caption         =   "Fic.Cult."
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
         Left            =   2550
         TabIndex        =   119
         Top             =   1365
         Width           =   915
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   0
         Left            =   1590
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   4170
         Width           =   240
      End
      Begin VB.Label Label22 
         Caption         =   "Orden"
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
         Left            =   4980
         TabIndex        =   98
         Top             =   1365
         Width           =   600
      End
      Begin VB.Label Label20 
         Caption         =   "Transportado por"
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
         TabIndex        =   95
         Top             =   4140
         Width           =   1515
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1080
         ToolTipText     =   "Buscar Campo"
         Top             =   1365
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1080
         ToolTipText     =   "Buscar Tarifa"
         Top             =   3780
         Width           =   240
      End
      Begin VB.Label Label12 
         Caption         =   "Tarifa"
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
         Left            =   150
         TabIndex        =   75
         Top             =   3750
         Width           =   825
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1080
         ToolTipText     =   "Buscar Transportista"
         Top             =   3360
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Transp."
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
         Left            =   150
         TabIndex        =   73
         Top             =   3330
         Width           =   825
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1080
         ToolTipText     =   "Buscar Capataz"
         Top             =   2940
         Width           =   240
      End
      Begin VB.Label Label23 
         Caption         =   "Capataz"
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
         Left            =   150
         TabIndex        =   71
         Top             =   2940
         Width           =   825
      End
      Begin VB.Label Label11 
         Caption         =   "Recolectado"
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
         Left            =   1890
         TabIndex        =   69
         Top             =   4140
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "N�Campo"
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
         Left            =   150
         TabIndex        =   66
         Top             =   1365
         Width           =   1080
      End
      Begin VB.Label Label4 
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
         Left            =   150
         TabIndex        =   62
         Top             =   930
         Width           =   810
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1080
         ToolTipText     =   "Buscar Socio"
         Top             =   930
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Entrada"
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
         Left            =   150
         TabIndex        =   58
         Top             =   4140
         Width           =   1485
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
         Left            =   150
         TabIndex        =   56
         Top             =   540
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Partida"
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
         TabIndex        =   55
         Top             =   1770
         Width           =   855
      End
      Begin VB.Label Label28 
         Caption         =   "Poblacion"
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
         Left            =   150
         TabIndex        =   54
         Top             =   2160
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Zona"
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
         Left            =   150
         TabIndex        =   53
         Top             =   2550
         Width           =   855
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   25
      Left            =   9600
      TabIndex        =   47
      Top             =   1245
      Width           =   1425
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   14700
      TabIndex        =   129
      Top             =   240
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
         Caption         =   "&Tarar Salida"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnPendiente 
         Caption         =   "&Pendiente Tara Salida"
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
Attribute VB_Name = "frmEntBasculaQua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MONICA                   -+-+
' +-+- Men�: Entrada de Bascula        -+-+
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

Private Const IdPrograma = 4006
'Dim T1 As Single

Public DatosADevolverBusqueda As String    'Tindr� el n� de text que vol que torne, empipat
Public Event DatoSeleccionado(CadenaSeleccion As String)

' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmEntPrev As frmBasico2
Attribute frmEntPrev.VB_VarHelpID = -1

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
Private WithEvents frmMens2 As frmMensajes 'mensajes
Attribute frmMens2.VB_VarHelpID = -1

' *****************************************************
Dim CodTipoMov As String
Dim v_cadena As String

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
Dim cadB As String

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
Private cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Public ImpresoraDefecto As String

Dim Lineas As Collection
Dim NF As Integer
Dim VarieAnt As String


Private Sub cmdAceptar_Click()
Dim Mens As String

    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'B�SQUEDA
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
                    Mens = ""
                    If Not ActualizarChivato(Mens, "U") Then
                        MsgBox "Error actualizando chivato: " & vbCrLf & Mens, vbExclamation
                    End If
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
'           menu = Me.Height - Me.ScaleHeight 'ac� tinc el heigth del men� i de la toolbar
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

' *** si n'hi han combos a la cap�alera ***
Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbLightBlue
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Combo1(Index).BackColor = vbLightBlue Then Combo1(Index).BackColor = vbWhite
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
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
Dim i As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    
    ' ICONETS DE LA BARRA
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
        .Buttons(8).Image = 10  'Imprimir
    End With
    
    With Me.Toolbar2
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26  'tarar tractor
        .Buttons(2).Image = 24  'paletizacion
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
    
'    ' ******* si n'hi han ll�nies *******
'    'ICONETS DE LES BARRES ALS TABS DE LL�NIA
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
    
    For i = 0 To imgAyuda.Count - 1
        imgAyuda(i).Picture = frmPpal.ImageListB.ListImages(10).Picture
    Next i
    
    'cargar IMAGES de busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    For i = 0 To 4
        Me.Label15(i).Caption = ""
        Me.Label19(i).Caption = ""
    Next i
    
    For i = 5 To 9
        Me.Label19(i).Caption = ""
    Next i
    
    ' cargamos los labels de parametros
    If vParamAplic.TipoCaja1 <> "" Then
        Me.Label15(0).Caption = vParamAplic.TipoCaja1
        Me.Label19(0).Caption = "x  " & vParamAplic.PesoCaja1
        Me.Label19(5).Caption = "x  " & vParamAplic.PesoCaja1
    End If
    If vParamAplic.TipoCaja2 <> "" Then
        Me.Label15(1).Caption = vParamAplic.TipoCaja2
        Me.Label19(1).Caption = "x  " & vParamAplic.PesoCaja2
        Me.Label19(6).Caption = "x  " & vParamAplic.PesoCaja2
    End If
    If vParamAplic.TipoCaja3 <> "" Then
        Me.Label15(2).Caption = vParamAplic.TipoCaja3
        Me.Label19(2).Caption = "x  " & vParamAplic.PesoCaja3
        Me.Label19(7).Caption = "x  " & vParamAplic.PesoCaja3
    End If
    If vParamAplic.TipoCaja4 <> "" Then
        Me.Label15(3).Caption = vParamAplic.TipoCaja4
        Me.Label19(3).Caption = "x  " & vParamAplic.PesoCaja4
        Me.Label19(8).Caption = "x  " & vParamAplic.PesoCaja4
    End If
    If vParamAplic.TipoCaja5 <> "" Then
        Me.Label15(4).Caption = vParamAplic.TipoCaja5
        Me.Label19(4).Caption = "x  " & vParamAplic.PesoCaja5
        Me.Label19(9).Caption = "x  " & vParamAplic.PesoCaja5
    End If
    
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

'    ' ******* si n'hi han ll�nies *******
'    DataGridAux(0).ClearFields
'    DataGridAux(1).ClearFields
    
    '*** canviar el nom de la taula i l'ordenaci� de la cap�alera ***
    NombreTabla = "rentradas"
    Ordenacion = " ORDER BY numnotac asc "
    '************************************************
    
    'Mirem com est� guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    '***** canviar el nom de la PK de la cap�alera; repasar codEmpre *************
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
         
    ' *** si n'hi han combos (cap�alera o ll�nies) ***
    CargaCombo
    ' ************************************************
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1 'b�squeda
        ' *** posar de groc els camps visibles de la clau primaria de la cap�alera ***
        Text1(0).BackColor = vbLightBlue 'codclien
        ' ****************************************************************************
    End If
End Sub


Private Sub LimpiarCampos()
Dim i As Integer

    On Error Resume Next
    
    limpiar Me   'M�tode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
'    Me.chkAbonos(0).Value = 0
    
    For i = 0 To Combo1.Count - 1
        Combo1(i).ListIndex = -1
'        Me.chkAbonos(I).Value = 0
    Next i
    
    ' *** si n'hi han combos a la cap�alera ***
    ' *****************************************

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
    
    If Modo = 5 Then Me.lblIndicador.Caption = "Tarar Salida"
    
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
    cmdCancelar.visible = b
    cmdAceptar.visible = b
       
    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a m�s neteja els camps Text1
    
' cambio la siguiente expresion por la de abajo
'   BloquearText1 Me, Modo
    For i = 0 To Text1.Count - 1
        BloquearTxt Text1(i), Not (Modo = 1 Or Modo = 3 Or Modo = 4)
    Next i
    
    BloquearCombo Me, Modo
    
    If Modo = 3 Then
        Combo1(1).ListIndex = 1
        Combo1(2).ListIndex = 0
        Combo1(3).ListIndex = 1
    End If
    
    If vParamAplic.NroNotaManual Then
        'claveprimaria
        BloquearTxt Text1(0), Not (Modo = 1 Or (Modo = 3 And vParamAplic.NroNotaManual) Or Modo = 4)
    Else
        b = (Modo <> 1)
        'Campos N� entrada bloqueado y en azul
        BloquearTxt Text1(0), b, True
    End If
    
    
    PonerTarasVisibles

    'taras desbloqueadas unicamente para buscar
    ' taras entrada
    For i = 18 To 20
        BloquearTxt Text1(i), Not (Modo = 1)
        Text1(i).Enabled = (Modo = 1)
    Next i
    For i = 8 To 9
        BloquearTxt Text1(i), Not (Modo = 1)
        Text1(i).Enabled = (Modo = 1)
    Next i
    ' taras salida
    For i = 30 To 34
        BloquearTxt Text1(i), Not (Modo = 1)
        Text1(i).Enabled = (Modo = 1)
    Next i
    ' peso neto
    BloquearTxt Text1(11), Not (Modo = 1)
    Text1(11).Enabled = (Modo = 1)

    For i = 22 To 22
        BloquearTxt Text1(i), Not (Modo = 3 Or Modo = 4 Or Modo = 1)
    Next i
    
    BloquearTxt Text1(3), Not (Modo = 1 Or Modo = 5)
    BloquearTxt Text1(24), Not (Modo = 1 Or Modo = 5)
    For i = 25 To 29
        BloquearTxt Text1(i), Not (Modo = 1 Or Modo = 5)
    Next i
    For i = 13 To 17
        BloquearTxt Text1(i), (Modo = 5) Or (Modo = 2) Or (Modo = 0)
    Next i
    
    Frame4.Enabled = (Modo = 1 Or Modo = 5)
    Frame3.Enabled = (Modo = 1 Or Modo = 3 Or Modo = 4)
    
    
    ' **** si n'hi han imagens de buscar en la cap�alera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    For i = 0 To imgFec.Count - 1
        BloquearImgFec Me, i, Modo
    Next i
'    BloquearImgFec Me, 25, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
    ' ********************************************************
    ' *** si n'hi han ll�nies i imagens de buscar que no estiguen als grids ******
    'Ll�nies Departaments
'    b = (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) And (NumTabMto = 0))
    ' ****************************************************************************
            
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    PonerLongCampos

'     '-----------------------------
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
    Toolbar1.Buttons(1).Enabled = b
    Me.mnNuevo.Enabled = b
    
    'Paletizacion
    Toolbar2.Buttons(2).Enabled = b
    Me.mnPendiente.Enabled = b
    
    
    
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) 'And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b
    'tara tractor
    Toolbar2.Buttons(1).Enabled = b
    'Imprimir
    'Toolbar1.Buttons(12).Enabled = (b Or Modo = 0)
    Toolbar1.Buttons(8).Enabled = b
       
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
        '   Com la clau principal es �nica, en posar el sql apuntant
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
Dim OtroCampo As String
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
        MsgBox "El campo no est� dado de alta. Reintroduzca.", vbExclamation
        Text1(5).Text = ""
        PonerFoco Text1(5)
    End If
End Sub

Private Sub frmCap_DatoSeleccionado(CadenaSeleccion As String)
    Text1(12).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo de capataz
    FormateaCampo Text1(12)
    Text2(12).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre de capataz
End Sub

Private Sub frmEntPrev_DatoSeleccionado(CadenaSeleccion As String)
Dim cadB As String
    
    If CadenaSeleccion <> "" Then
        cadB = "numnotac = " & DBSet(RecuperaValor(CadenaSeleccion, 1), "N")
        
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
    PonerDatosCampo Text1(5)
End Sub

Private Sub frmMens2_DatoSeleccionado(CadenaSeleccion As String)
    cadB = " numnotac = " & RecuperaValor(CadenaSeleccion, 1)
    If chkVistaPrevia(0) = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If


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
    MsgBox vCadena, vbInformation, "Descripci�n de Ayuda"
    
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


    If Data1.Recordset.EOF Then Exit Sub
    
    Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
    Dim nomDocu As String 'Nombre de Informe rpt de crystal
    Dim ImprimeDirecto As Integer
     
    indRPT = 25 'Ticket de Entrada
     
    If Not PonerParamRPT(indRPT, "", 1, nomDocu, , ImprimeDirecto) Then Exit Sub
    'Nombre fichero .rpt a Imprimir
    ' he a�adido estas dos lineas para que llame al rpt correspondiente
    
    If ImprimeDirecto = 0 Then
        frmImprimir.NombreRPT = nomDocu
        
        ActivaTicket
        
        With frmVisReport
            .FormulaSeleccion = "{rentradas.numnotac}=" & Data1.Recordset!NumNotac
            .SoloImprimir = True
            .OtrosParametros = ""
            .NumeroParametros = 1
            .MostrarTree = False
            .Informe = App.Path & "\informes\" & nomDocu    ' "ValEntrada.rpt"
            .InfConta = False
            .ConSubInforme = False
            .SubInformeConta = ""
            .Opcion = 0
            .ExportarPDF = False
            .Show vbModal
        End With
        
        DesactivaTicket
    Else
    
' ahora cojo la impresion directa de david de quatretonda
            ImprimirDirectoAlb "rentradas.numnotac = " & Text1(0).Text

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


Private Sub mnPendiente_Click()
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cajas As Currency
Dim Cad As String

    
    Sql = "select count(*) from rentradas where taracajasa1 is null and taracajasa2 is null and taracajasa3 is null and taracajasa4 is null and taracajasa5 is null and taravehisa is null and otrastarasa is null "
    If TotalRegistros(Sql) = 0 Then
        Cad = "No hay entradas pendientes de tarar salida." & vbCrLf
        MsgBox Cad, vbExclamation
    Else
        Sql = "select numnotac from rentradas where taracajasa1 is null and taracajasa2 is null and taracajasa3 is null and taracajasa4 is null and taracajasa5 is null and taravehisa is null and otrastarasa is null "
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        Cad = ""
        
        While Not Rs.EOF
            Cad = Cad & DBSet(Rs!NumNotac, "N") & ","
            
            Rs.MoveNext
        Wend
        
        Set Rs = Nothing
        
        If Cad <> "" Then
            Cad = Mid(Cad, 1, Len(Cad) - 1)
            Cad = " numnotac in (" & Cad & ")"
            
            Set frmMens2 = New frmMensajes
            
            frmMens2.cadWhere = Cad
            frmMens2.OpcionMensaje = 28
            frmMens2.Show vbModal
            
            Set frmMens2 = Nothing
            
            If Data1.Recordset.EOF Then Exit Sub
        
            mnTararTractor_Click
            
        End If
        
    End If
    
End Sub






Private Sub mnPaletizacion_Click()
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cajas As Currency
Dim Cad As String

    If vParamAplic.HayTraza = False Then Exit Sub
    
    frmEntBascula2.crear = 1
    
    Sql = "select count(*) from trzpalets where numnotac = " & Trim(Data1.Recordset!NumNotac)
    If TotalRegistros(Sql) <> 0 Then
        Cad = "La paletizaci�n para esta entrada ya est� realizada." & vbCrLf
        Cad = Cad & vbCrLf & "            � Desea crearla de nuevo ? "
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
    If vParamAplic.EsCaja1 Then cajas = cajas + DBLet(Data1.Recordset!numcajo1, "N")
    If vParamAplic.EsCaja2 Then cajas = cajas + DBLet(Data1.Recordset!numcajo2, "N")
    If vParamAplic.EsCaja3 Then cajas = cajas + DBLet(Data1.Recordset!numcajo3, "N")
    If vParamAplic.EsCaja4 Then cajas = cajas + DBLet(Data1.Recordset!numcajo4, "N")
    If vParamAplic.EsCaja5 Then cajas = cajas + DBLet(Data1.Recordset!numcajo5, "N")
    
    frmEntBascula2.NumNota = ImporteSinFormato(Data1.Recordset!NumNotac)
    frmEntBascula2.NumCajones = CStr(cajas)
    frmEntBascula2.NumKilos = ImporteSinFormato(Text1(11).Text)
    frmEntBascula2.Codsocio = Text1(1).Text
    frmEntBascula2.codCampo = Text1(5).Text
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
        Case 1  'Nou
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
        Case 5  'B�scar
           mnBuscar_Click
        Case 6  'Tots
            mnVerTodos_Click
        Case 8 'Imprimir
'            AbrirListado (10)
            mnImprimir_Click
    End Select
End Sub

Private Sub BotonBuscar()
Dim i As Integer
' ***** Si la clau primaria de la cap�alera no es Text1(0), canviar-ho en <=== *****
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonerFoco Text1(0) ' <===
        Text1(0).BackColor = vbLightBlue ' <===
        ' *** si n'hi han combos a la cap�alera ***
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
    
    If Text1(22).Text <> "" Then
        Text1(4).Text = Text1(22).Text
        Text1(4).Tag = Replace(Text1(8).Tag, "FH", "FHH")
    End If

    cadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
    Text1(4).Tag = Replace(Text1(4).Tag, "FHH", "FH")
    
    If chkVistaPrevia(0) = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        ' *** foco al 1r camp visible de la cap�alera que siga clau primaria ***
        PonerFoco Text1(0)
        ' **********************************************************************
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)


    Set frmEntPrev = New frmBasico2
    
    AyudaEntradaBascula frmEntPrev
    
    Set frmEntPrev = Nothing


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
    cadB = ""
    
    PonerModo 0
    
    If chkVistaPrevia(0).Value = 1 Then
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
    ' *** si n'hi han camps de descripci� a la cap�alera ***
    PosarDescripcions
    ' ******************************************************
    
    Text1(35).Text = "55.000.017"
    
    ' *** si n'hi han tabs, em posicione al 1r ***
    Me.SSTab1.Tab = 0
    ' ********************************************
End Sub


Private Sub BotonModificar()



'    NumCajasAnt = CCur(ComprobarCero(Text1(13).Text)) + CCur(ComprobarCero(Text1(14).Text)) + _
'               CCur(ComprobarCero(Text1(15).Text)) + CCur(ComprobarCero(Text1(16).Text)) + _
'               CCur(ComprobarCero(Text1(17).Text))

    ' ahora el numero de cajas se suma unicamente si rparam.escaja es true
    NumCajasAnt = 0
    If vParamAplic.EsCaja1 Then NumCajasAnt = NumCajasAnt + CCur(ComprobarCero(Text1(13).Text))
    If vParamAplic.EsCaja2 Then NumCajasAnt = NumCajasAnt + CCur(ComprobarCero(Text1(14).Text))
    If vParamAplic.EsCaja3 Then NumCajasAnt = NumCajasAnt + CCur(ComprobarCero(Text1(15).Text))
    If vParamAplic.EsCaja4 Then NumCajasAnt = NumCajasAnt + CCur(ComprobarCero(Text1(16).Text))
    If vParamAplic.EsCaja5 Then NumCajasAnt = NumCajasAnt + CCur(ComprobarCero(Text1(17).Text))
    

    NumKilosAnt = CCur(ComprobarCero(Text1(11).Text))
    PonerModo 4

    ' *** bloquejar els camps visibles de la clau primaria de la cap�alera ***
    BloquearTxt Text1(0), True
    ' *************************************************************************
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(10)
    ' *********************************************************
End Sub

Private Sub BotonTarar()

    PonerModo 5

    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(25)
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
    Cad = "�Seguro que desea eliminar la Entrada?"
    Cad = Cad & vbCrLf & "N�mero: " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
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
Dim Sql As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la cap�alera
    
    Text1(22).Text = Mid(Text1(4).Text, 12, 8)
    
    
    CalculoCampoMermas
    
    PosarDescripcions

    ' ********************************************************************************
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu   'Activar opcions de men� seg�n nivell
                        'de permisos de l'usuari
    
End Sub

Private Sub CalculoCampoMermas()
Dim Sql As String
Dim PorcMerma As Currency
Dim KilosNet As Long
Dim TotMerma As Long
Dim Bruto As Long

    Sql = "select porcmerm from variedades where codvarie = " & DBSet(Text1(2).Text, "N")
    PorcMerma = DevuelveValor(Sql)
    
    Bruto = Text1(21).Text
    
    If Modo = 2 Then  ' se trata unicamente de visualizar segun los datos introducidos
        TotMerma = Bruto
        If SalidaTarada Then
            If Text1(18).Text <> "" Then TotMerma = TotMerma - CCur(Text1(18).Text) ' menos taras caja de salida
            If Text1(19).Text <> "" Then TotMerma = TotMerma - CCur(Text1(19).Text)
            If Text1(20).Text <> "" Then TotMerma = TotMerma - CCur(Text1(20).Text)
            If Text1(8).Text <> "" Then TotMerma = TotMerma - CCur(Text1(8).Text)
            If Text1(9).Text <> "" Then TotMerma = TotMerma - CCur(Text1(9).Text)
            
            If Text1(30).Text <> "" Then TotMerma = TotMerma + CCur(Text1(30).Text) ' mas taras caja de entrada
            If Text1(31).Text <> "" Then TotMerma = TotMerma + CCur(Text1(31).Text)
            If Text1(32).Text <> "" Then TotMerma = TotMerma + CCur(Text1(32).Text)
            If Text1(33).Text <> "" Then TotMerma = TotMerma + CCur(Text1(33).Text)
            If Text1(34).Text <> "" Then TotMerma = TotMerma + CCur(Text1(34).Text)
            
            If Text1(3).Text <> "" Then TotMerma = TotMerma - CCur(Text1(3).Text) 'menos taravehi salida
            If Text1(24).Text <> "" Then TotMerma = TotMerma - CCur(Text1(24).Text) 'menos otras taras salida
        
        Else
            If Text1(18).Text <> "" Then TotMerma = TotMerma - CCur(Text1(18).Text)
            If Text1(19).Text <> "" Then TotMerma = TotMerma - CCur(Text1(19).Text)
            If Text1(20).Text <> "" Then TotMerma = TotMerma - CCur(Text1(20).Text)
            If Text1(8).Text <> "" Then TotMerma = TotMerma - CCur(Text1(8).Text)
            If Text1(9).Text <> "" Then TotMerma = TotMerma - CCur(Text1(9).Text)
        End If
        TotMerma = TotMerma - CCur(Text1(11).Text) 'menos kilosnet
    
    Else
        KilosNet = Bruto
        
        If SalidaTarada Then
            If Text1(18).Text <> "" Then KilosNet = KilosNet - CCur(Text1(18).Text) ' menos taras caja de salida
            If Text1(19).Text <> "" Then KilosNet = KilosNet - CCur(Text1(19).Text)
            If Text1(20).Text <> "" Then KilosNet = KilosNet - CCur(Text1(20).Text)
            If Text1(8).Text <> "" Then KilosNet = KilosNet - CCur(Text1(8).Text)
            If Text1(9).Text <> "" Then KilosNet = KilosNet - CCur(Text1(9).Text)
            
            If Text1(30).Text <> "" Then KilosNet = KilosNet + CCur(Text1(30).Text) ' mas taras caja de entrada
            If Text1(31).Text <> "" Then KilosNet = KilosNet + CCur(Text1(31).Text)
            If Text1(32).Text <> "" Then KilosNet = KilosNet + CCur(Text1(32).Text)
            If Text1(33).Text <> "" Then KilosNet = KilosNet + CCur(Text1(33).Text)
            If Text1(34).Text <> "" Then KilosNet = KilosNet + CCur(Text1(34).Text)
            
            If Text1(3).Text <> "" Then KilosNet = KilosNet - CCur(Text1(3).Text) 'menos taravehi salida
            If Text1(24).Text <> "" Then KilosNet = KilosNet - CCur(Text1(24).Text) 'menos otras taras salida
        
        Else
            If Text1(18).Text <> "" Then KilosNet = KilosNet - CCur(Text1(18).Text) ' menos taras caja de salida
            If Text1(19).Text <> "" Then KilosNet = KilosNet - CCur(Text1(19).Text)
            If Text1(20).Text <> "" Then KilosNet = KilosNet - CCur(Text1(20).Text)
            If Text1(8).Text <> "" Then KilosNet = KilosNet - CCur(Text1(8).Text)
            If Text1(9).Text <> "" Then KilosNet = KilosNet - CCur(Text1(9).Text)
        
        End If
        
        TotMerma = Round2(KilosNet * PorcMerma * 0.01, 0)
        KilosNet = KilosNet - TotMerma
        Text1(11).Text = Format(KilosNet, "###,##0")
    End If
    
    Text2(8).Text = Format(TotMerma, "###,##0")
    
End Sub

Private Function SalidaTarada() As Boolean

    SalidaTarada = (Text1(30).Text <> "" Or Text1(31).Text <> "" Or Text1(32).Text <> "" Or Text1(33).Text <> "" Or _
                    Text1(34).Text <> "" Or Text1(3).Text <> "" Or Text1(24).Text <> "")

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
                TerminaBloquear
                PonerModo 2
                PonerCampos
                ' *** primer camp visible de la cap�alera ***
                PonerFoco Text1(0)
            
'            Select Case ModoLineas
'                Case 1 'afegir ll�nia
'                    ModoLineas = 0
'                    ' *** les ll�nies que tenen datagrid (en o sense tab) ***
'                    If NumTabMto = 0 Or NumTabMto = 1 Or NumTabMto = 2 Or NumTabMto = 4 Then
'                        DataGridAux(NumTabMto).AllowAddNew = False
'                        ' **** repasar si es diu Data1 l'adodc de la cap�alera ***
'                        'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
'                        ' ********************************************************
'                        LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
'                        DataGridAux(NumTabMto).Enabled = True
'                        DataGridAux(NumTabMto).SetFocus
'
'                        ' *** si n'hi han camps de descripci� dins del grid, els neteje ***
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
'                Case 2 'modificar ll�nies
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
'                        ' *** l'Index de Fields es el que canvie de la PK de ll�nies ***
'                        V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el n� de llinia
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
'            ' *** si n'hi han ll�nies en grids i camps fora d'estos ***
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
                Sql = "El socio introducido est� dado de baja. Reintroduzca. " & vbCrLf & vbCrLf
                MsgBox Sql, vbExclamation
                b = False
                PonerFoco Text1(1)
            End If
        End If
        
        If b Then
            ' comprobamos que el campo no est� dado de baja
            If Not EstaCampoDeAlta(Text1(5).Text) Then
                Sql = "El campo introducido est� dado de baja. Reintroduzca. " & vbCrLf & vbCrLf
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

    ' *** canviar-ho per tota la PK de la cap�alera, no llevar els () ***
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
Dim vTipoMov As CTiposMov

    On Error GoTo FinEliminar

    conn.BeginTrans
    ' ***** canviar el nom de la PK de la cap�alera, repasar codEmpre *******
    vWhere = " WHERE numnotac=" & Data1.Recordset!NumNotac
        ' ***********************************************************************
        
    Mens = "Actualizar chivato"
    b = ActualizarChivato(Mens, "Z")
        
        
    ' ***** elimina les ll�nies ****
    conn.Execute "DELETE FROM trzpalets where numnotac = " & Trim(CStr(Data1.Recordset!NumNotac))

    conn.Execute "Delete from " & NombreTabla & vWhere
       
FinEliminar:
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, "Eliminar"
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    
        If Not vParamAplic.NroNotaManual Then
            'Decrementar contador si borramos el ultima factura
            Set vTipoMov = New CTiposMov
            vTipoMov.DevolverContador CodTipoMov, Val(Data1.Recordset!NumNotac)
            Set vTipoMov = Nothing
        End If
    
    End If
End Function


Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
'    If Index = 11 Or Index = 18 Or Index = 19 Or Index = 20 Or Index = 8 Or Index = 9 Or _
'       Index = 30 Or Index = 31 Or Index = 32 Or Index = 33 Or Index = 34 Then
'       If Modo <> 1 Then Exit Sub
'    End If
    If Index = 2 Then VarieAnt = Text1(2).Text
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
                    cadMen = cadMen & "�Desea crearlo?" & vbCrLf
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
                        MsgBox "El socio est� dado de baja. Reintroduzca.", vbExclamation
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
                    cadMen = cadMen & "�Desea crearla?" & vbCrLf
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
                    If VarieAnt <> Text1(Index).Text Then
                        If (Modo = 3 Or Modo = 4) And EsVariedadGrupo6(Text1(Index).Text) Then
                            MsgBox "Esta variedad es del Grupo de Bodega. Revise.", vbExclamation
                            PonerFoco Text1(Index)
                        Else
                            '[Monica]26/09/2011: cuando es de almazara solo es dalt o terra
                            If (Modo = 1 Or Modo = 3 Or Modo = 4) And EsVariedadGrupo5(Text1(Index).Text) Then
                                'es una variedad del grupo de almazara
                                CargaComboAlm
                                Me.Combo1(0).ListIndex = 0
                                If Modo = 1 Then Me.Combo1(0).ListIndex = -1
                                imgAyuda(0).visible = False
                                imgAyuda(0).Enabled = False
                            Else
                                CargaComboNormal
                                Me.Combo1(0).ListIndex = 0
                                imgAyuda(0).visible = True
                                imgAyuda(0).Enabled = True
                            End If
                            PonerCamposSocioVariedad
                       End If
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
                    cadMen = cadMen & "�Desea crearlo?" & vbCrLf
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
                        MsgBox "El campo no est� dado de alta. Reintroduzca.", vbExclamation
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    Else
                        '[Monica]13/08/2018: si es campo de tratamiento no se permiten entradas
                        If EsCampoDeTratamiento(Text1(Index).Text) Then
                            MsgBox "El campo es de tratamiento. Reintroduzca.", vbExclamation
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
            End If
        
                
        Case 6 'transportistas
            If Text1(Index).Text <> "" Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "rtransporte", "nomtrans")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Transportista: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearlo?" & vbCrLf
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
'                    If Modo = 3 Then ' solo si estamos insertando metemos la tara del vehiculo
'                        Sql = "select taravehi from rtransporte where codtrans = " & DBSet(Text1(Index), "T")
'                        Text1(3).Text = DevuelveValor(Sql)
'                        PonerFormatoEntero Text1(3)
'                    End If
                End If
            Else
                Text2(Index).Text = ""
            End If
                
        Case 7 'tarifa de transporte
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rtarifatra", "nomtarif")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Tarifa de Transporte: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearla?" & vbCrLf
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
                    cadMen = cadMen & "�Desea crearlo?" & vbCrLf
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
            
        Case 13, 14, 15, 16, 17, 21, 25, 26, 27, 28, 29 'pesos
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
        
            
        Case 10 'Fecha no comprobaremos que est� dentro de campa�a
                    'Fecha de alta y fecha de baja
            '[Monica]28/08/2013: comprobamos que la fecha est� en la campa�a
            PonerFormatoFecha Text1(Index), True
            
    End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 10: KEYFecha KeyAscii, 0 'fecha entrada
                Case 2: KEYBusqueda KeyAscii, 2 'variedad
                Case 1: KEYBusqueda KeyAscii, 1 'socio
                Case 5: KEYBusqueda KeyAscii, 5 'campo
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
'    ModoLineas = 3 'Posem Modo Eliminar Ll�nia
'
'    If Modo = 4 Then 'Modificar Cap�alera
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
'            sql = "�Seguro que desea eliminar el telefono?"
'            sql = sql & vbCrLf & "Tel�fono: " & AdoAux(Index).Recordset!idtelefono & " - " & AdoAux(Index).Recordset!imei
'            If MsgBox(sql, vbQuestion + vbYesNo) = vbYes Then
'                eliminar = True
'                sql = "DELETE FROM rsocios_telefonos"
'                sql = sql & vWhere & " AND idtelefono= " & DBLet(AdoAux(Index).Recordset!idtelefono, "T")
'            End If
'        Case 1 'secciones
'            sql = "�Seguro que desea eliminar la secci�n?"
'            sql = sql & vbCrLf & "Secci�n: " & AdoAux(Index).Recordset!codsecci & " - " & AdoAux(Index).Recordset!nomsecci
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
'    ModoLineas = 1 'Posem Modo Afegir Ll�nia
'
'    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Cap�alera
'        cmdAceptar_Click
'        If ModoLineas = 0 Then Exit Sub
'    End If
'
'    NumTabMto = Index
'    PonerModo 5, Index
'
'    ' *** bloquejar la clau primaria de la cap�alera ***
'    BloquearTxt Text1(0), True
'    ' **************************************************
'
'    ' *** posar el nom del les distintes taules de ll�nies ***
'    Select Case Index
'        Case 0: vTabla = "rsocios_telefonos"
'        Case 1: vTabla = "rsocios_seccion"
'    End Select
'    ' ********************************************************
'
'    vWhere = ObtenerWhereCab(False)
'
'    Select Case Index
'        Case 0 ' *** pose els index dels tabs de ll�nies que tenen datagrid ***
'            ' *** canviar la clau primaria de les ll�nies,
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
'            ' *** canviar la clau primaria de les ll�nies,
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
''        ' *** si n'hi han ll�nies sense datagrid ***
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
'    ModoLineas = 2 'Modificar ll�nia
'
'    If Modo = 4 Then 'Modificar Cap�alera
'        cmdAceptar_Click
'        If ModoLineas = 0 Then Exit Sub
'    End If
'
'    NumTabMto = Index
'    PonerModo 5, Index
'    ' *** bloqueje la clau primaria de la cap�alera ***
'    BloquearTxt Text1(0), True
'    ' *********************************
'
'    Select Case Index
'        Case 0, 1 ' *** pose els index de ll�nies que tenen datagrid (en o sense tab) ***
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

'    ' *** si n'hi han tabs sense datagrid posar el If ***
'    If Index <> 3 Then DeseleccionaGrid DataGridAux(Index)
'    ' ***************************************************
'
'    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Ll�nies
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
'    ' ******* configurar el LostFocus dels camps de ll�nies (dins i fora grid) ********
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
'
'    ' *** si cal fer atres comprovacions a les ll�nies (en o sense tab) ***
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
''                Mens = "Debe seleccionar que esta cuenta est� activa si desea que sea la principal"
''            End If
'
''            'No puede haber m�s de una cuenta principal
''            If cant > 0 And (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 1) Then
''                Mens = "No puede haber m�s de una cuenta principal."
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


Private Function ActualisaCtaprpal(ByRef NumLinea As Integer)
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

' *** si n'hi han formularis de buscar codi a les ll�nies ***
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
'        NetejaFrameAux "FrameAux3" 'neteja nom�s lo que te TAG
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
'            tots = "N||||0|;S|txtaux1(1)|T|C�d.|800|;S|cmdAux(4)|B|||;" 'codsocio,codsecci
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
'    ' **** si n'hi han ll�nies en grids i camps fora d'estos ****
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
''Inserta registre en les taules de Ll�nies
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
''Modifica registre en les taules de Ll�nies
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
'                V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el n� de llinia
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
    ' *** canviar-ho per la clau primaria de la cap�alera ***
    vWhere = vWhere & " codsocio=" & Val(Text1(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function

' *** neteja els camps dels tabs de grid que
'estan fora d'este, i els camps de descripci� ***
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

    'Ausencia Plagas
    Combo1(3).AddItem "No"
    Combo1(3).ItemData(Combo1(3).NewIndex) = 0
    Combo1(3).AddItem "S�"
    Combo1(3).ItemData(Combo1(3).NewIndex) = 1

End Sub

Private Sub CargaComboAlm()
    
    Combo1(0).Clear
    
    'tipo de entrada
    Combo1(0).AddItem "Dalt"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
'    Combo1(0).AddItem "V.Campo"
'    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
'    Combo1(0).AddItem "P.Integrado"
'    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
'    Combo1(0).AddItem "Ind.Directo"
'    Combo1(0).ItemData(Combo1(0).NewIndex) = 3
    Combo1(0).AddItem "Terra"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 4
'    Combo1(0).AddItem "Venta Directo"
'    Combo1(0).ItemData(Combo1(0).NewIndex) = 5

End Sub


Private Sub CargaComboNormal()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 0 To 0
        Combo1(i).Clear
    Next i
    
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
'    ' ******* configurar el LostFocus dels camps de ll�nies (dins i fora grid) ********
'    Select Case Index
'        Case 1 ' seccion
'                If PonerFormatoEntero(txtaux1(Index)) Then
'                    Set vSeccion = New CSeccion
'                    If vSeccion.LeerDatos(txtaux1(Index)) Then
'                        txtAux2(Index).Text = vSeccion.Nombre
'                        b = vSeccion.AbrirConta
'                    Else
'                        Set vSeccion = Nothing
'                        cadMen = "No existe la Secci�n: " & txtaux1(Index).Text & vbCrLf
'                        cadMen = cadMen & "�Desea crearla?" & vbCrLf
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
''                    cadMen = "No existe la Secci�n: " & txtaux1(Index).Text & vbCrLf
''                    cadMen = cadMen & "�Desea crearla?" & vbCrLf
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
    cadselect = ""
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

Dim Tara1s As Currency
Dim Tara2s As Currency
Dim Tara3s As Currency
Dim Tara4s As Currency
Dim Tara5s As Currency
    
    
    
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
    
    Tara1s = 0
    Tara2s = 0
    Tara3s = 0
    Tara4s = 0
    Tara5s = 0
    
    
    Text1(18).Text = ""
    Text1(19).Text = ""
    Text1(20).Text = ""
    Text1(8).Text = ""
    Text1(9).Text = ""
    
    'tara 1
    If Text1(13).Text <> "" Then
        Tara1 = Round2(CCur(ImporteSinFormato(Text1(13).Text)) * vParamAplic.PesoCaja1, 0)
        Tara11 = Round2(CCur(ImporteSinFormato(Text1(13).Text)) * vParamAplic.PesoCaja11, 0)
        Text1(18).Text = Tara1
        PonerFormatoEntero Text1(18)
    End If
    'tara 2
    If Text1(14).Text <> "" Then
        Tara2 = Round2(CCur(ImporteSinFormato(Text1(14).Text)) * vParamAplic.PesoCaja2, 0)
        Tara12 = Round2(CCur(ImporteSinFormato(Text1(14).Text)) * vParamAplic.PesoCaja12, 0)
        Text1(19).Text = Tara2
        PonerFormatoEntero Text1(19)
    End If
    'tara 3
    If Text1(15).Text <> "" Then
        Tara3 = Round2(CCur(ImporteSinFormato(Text1(15).Text)) * vParamAplic.PesoCaja3, 0)
        Tara13 = Round2(CCur(ImporteSinFormato(Text1(15).Text)) * vParamAplic.PesoCaja13, 0)
        Text1(20).Text = Tara3
        PonerFormatoEntero Text1(20)
    End If
    'tara 4
    If Text1(16).Text <> "" Then
        Tara4 = Round2(CCur(ImporteSinFormato(Text1(16).Text)) * vParamAplic.PesoCaja4, 0)
        Tara14 = Round2(CCur(ImporteSinFormato(Text1(16).Text)) * vParamAplic.PesoCaja14, 0)
        Text1(8).Text = Tara4
        PonerFormatoEntero Text1(8)
    End If
    'tara 5
    If Text1(17).Text <> "" Then
        Tara5 = Round2(CCur(ImporteSinFormato(Text1(17).Text)) * vParamAplic.PesoCaja5, 0)
        Tara15 = Round2(CCur(ImporteSinFormato(Text1(17).Text)) * vParamAplic.PesoCaja15, 0)
        Text1(9).Text = Tara5
        PonerFormatoEntero Text1(9)
    End If

    ' taras de salida
    
    
    Text1(30).Text = ""
    Text1(31).Text = ""
    Text1(32).Text = ""
    Text1(33).Text = ""
    Text1(34).Text = ""
    
    'tara 1
    If Text1(25).Text <> "" Then
        Tara1s = Round2(CCur(ImporteSinFormato(Text1(25).Text)) * vParamAplic.PesoCaja1, 0)
        Tara11 = Round2(CCur(ImporteSinFormato(Text1(25).Text)) * vParamAplic.PesoCaja11, 0)
        Text1(30).Text = Tara1s
        PonerFormatoEntero Text1(30)
    End If
    'tara 2
    If Text1(26).Text <> "" Then
        Tara2s = Round2(CCur(ImporteSinFormato(Text1(26).Text)) * vParamAplic.PesoCaja2, 0)
        Tara12 = Round2(CCur(ImporteSinFormato(Text1(26).Text)) * vParamAplic.PesoCaja12, 0)
        Text1(31).Text = Tara2s
        PonerFormatoEntero Text1(31)
    End If
    'tara 3
    If Text1(27).Text <> "" Then
        Tara3s = Round2(CCur(ImporteSinFormato(Text1(27).Text)) * vParamAplic.PesoCaja3, 0)
        Tara13 = Round2(CCur(ImporteSinFormato(Text1(27).Text)) * vParamAplic.PesoCaja13, 0)
        Text1(32).Text = Tara3s
        PonerFormatoEntero Text1(32)
    End If
    'tara 4
    If Text1(28).Text <> "" Then
        Tara4s = Round2(CCur(ImporteSinFormato(Text1(28).Text)) * vParamAplic.PesoCaja4, 0)
        Tara14 = Round2(CCur(ImporteSinFormato(Text1(28).Text)) * vParamAplic.PesoCaja14, 0)
        Text1(33).Text = Tara4s
        PonerFormatoEntero Text1(33)
    End If
    'tara 5
    If Text1(29).Text <> "" Then
        Tara5s = Round2(CCur(ImporteSinFormato(Text1(29).Text)) * vParamAplic.PesoCaja5, 0)
        Tara15 = Round2(CCur(ImporteSinFormato(Text1(29).Text)) * vParamAplic.PesoCaja15, 0)
        Text1(34).Text = Tara5s
        PonerFormatoEntero Text1(34)
    End If

    
    'peso neto
    PesoBruto = 0
    TaraVehi = 0
    OtrasTaras = 0
    If Text1(21).Text <> "" Then PesoBruto = CCur(Text1(21).Text)
'    If Text1(3).Text <> "" Then TaraVehi = CCur(Text1(3).Text)
'    If Text1(24).Text <> "" Then OtrasTaras = CCur(Text1(24).Text)
    
    PesoNeto = PesoBruto - Tara1 - Tara2 - Tara3 - Tara4 - Tara5 - TaraVehi - OtrasTaras
'    PesoTrans = PesoBruto - Tara11 - Tara12 - Tara13 - Tara14 - Tara15 - TaraVehi - OtrasTaras
    Text1(11).Text = CStr(PesoNeto)
'    Text1(23).Text = CStr(PesoTrans)
    PonerFormatoEntero Text1(11)
    
    CalculoCampoMermas
    
    
End Sub

Private Sub PonerCamposSocioVariedad()
Dim Cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim Rs As ADODB.Recordset

    If Not (Modo = 3 Or Modo = 4) Then Exit Sub

    If Text1(1).Text = "" Or Text1(2).Text = "" Then Exit Sub
    

    Cad = "rcampos.codsocio = " & DBSet(Text1(1).Text, "N") & " and rcampos.fecbajas is null"
    '[Monica]13/08/2018: no se permiten entradas a los campos de tratamientos
    Cad = Cad & " and rcampos.tipocampo <> 3 "
    Cad = Cad & " and rcampos.codvarie = " & DBSet(Text1(2).Text, "N")
     
    Cad1 = "select count(*) from rcampos where " & Cad
     
    NumRegis = TotalRegistros(Cad1)
    
'    If NumRegis = 0 Then Exit Sub
'    If NumRegis = 1 Then
        Cad1 = "select codcampo from rcampos where " & Cad
        Set Rs = New ADODB.Recordset
        Rs.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            Text1(5).Text = DBLet(Rs.Fields(0).Value)
            PonerDatosCampo Text1(5).Text
        End If
'    Else
        Set frmMens = New frmMensajes
        frmMens.cadWhere = " and " & Cad '"rcampos.codsocio = " & NumCod & " and rcampos.fecbajas is null"
        frmMens.campo = Text1(5).Text
        frmMens.OpcionMensaje = 6
        '[Monica]13/08/2018: no se permiten entradas de trataniento
        frmMens.vCampos = " and rcampos.codsocio = " & DBSet(Text1(1).Text, "N") & " and rcampos.fecbajas is null and rcampos.tipocampo <> 3 "
        frmMens.Show vbModal
        Set frmMens = Nothing
'    End If
    
End Sub

Private Sub PonerDatosCampo(campo As String)
Dim Cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim Rs As ADODB.Recordset


    If campo = "" Then Exit Sub
    
'    If Not (Modo = 3 Or Modo = 4) Then Exit Sub

    Cad = "rcampos.codcampo = " & DBSet(campo, "N") & " and rcampos.fecbajas is null"
    '[Monica]13/08/2018: no se permiten entradas de trataniento
    Cad = Cad & " and rcampos.tipocampo <> 3"
     
    '[Monica]25/09/2013: a�adimos la ficha de cultivo
    Cad1 = "select rcampos.codparti, rpartida.nomparti, rpartida.codzonas, rzonas.nomzonas, "
    Cad1 = Cad1 & " rpueblos.despobla, rcampos.nrocampo, rfichculti.descripcion from rcampos, rpartida, rzonas, rpueblos, rfichculti "
    Cad1 = Cad1 & " where " & Cad
    Cad1 = Cad1 & " and rcampos.codparti = rpartida.codparti "
    Cad1 = Cad1 & " and rpartida.codzonas = rzonas.codzonas "
    Cad1 = Cad1 & " and rpartida.codpobla = rpueblos.codpobla "
    Cad1 = Cad1 & " and rcampos.entregafichaculti = rfichculti.codtipo "
     
    Set Rs = New ADODB.Recordset
    Rs.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
'    Text1(5).Text = ""
    Text2(0).Text = ""
    Text2(4).Text = ""
    Text2(3).Text = ""
    Text3(3).Text = ""
    Text4(3).Text = ""
    Text5(3).Text = ""
    '[Monica]25/09/2013: a�adida la ficha de cultivo
    Text2(5).Text = ""
    If Not Rs.EOF Then
        Text1(5).Text = campo
        PonerFormatoEntero Text1(5)
        Text2(4).Text = DBLet(Rs.Fields(0).Value, "N") ' codigo de partida
        If Text2(4).Text <> "" Then Text2(4).Text = Format(Text2(4).Text, "0000")
        Text2(3).Text = DBLet(Rs.Fields(1).Value, "T") ' nombre de partida
        Text3(3).Text = DBLet(Rs.Fields(2).Value, "N") ' codigo de zona
        If Text3(3).Text <> "" Then Text3(3).Text = Format(Text3(3).Text, "0000")
        Text4(3).Text = DBLet(Rs.Fields(3).Value, "T") ' nombre de zona
        Text5(3).Text = DBLet(Rs.Fields(4).Value, "T") ' descripcion de poblacion
        Text2(0).Text = DBLet(Rs.Fields(5).Value, "N") ' nro de campo
        
        '[Monica]25/09/2013: a�adida la ficha de cultivo
        Text2(5).Text = DBLet(Rs.Fields(6).Value, "T") ' ficha de cultivo
    End If
    
    Set Rs = Nothing
    
End Sub

Private Sub PonerTarasVisibles()
    'tara1
    Text1(13).Enabled = (vParamAplic.TipoCaja1 <> "")
    Text1(13).visible = (vParamAplic.TipoCaja1 <> "")
    Text1(18).Enabled = (vParamAplic.TipoCaja1 <> "")
    Text1(18).visible = (vParamAplic.TipoCaja1 <> "")
    Text1(25).Enabled = (vParamAplic.TipoCaja1 <> "")
    Text1(25).visible = (vParamAplic.TipoCaja1 <> "")
    Text1(30).Enabled = (vParamAplic.TipoCaja1 <> "")
    Text1(30).visible = (vParamAplic.TipoCaja1 <> "")

    'tara2
    Text1(14).Enabled = (vParamAplic.TipoCaja2 <> "")
    Text1(14).visible = (vParamAplic.TipoCaja2 <> "")
    Text1(19).Enabled = (vParamAplic.TipoCaja2 <> "")
    Text1(19).visible = (vParamAplic.TipoCaja2 <> "")
    Text1(26).Enabled = (vParamAplic.TipoCaja2 <> "")
    Text1(26).visible = (vParamAplic.TipoCaja2 <> "")
    Text1(31).Enabled = (vParamAplic.TipoCaja2 <> "")
    Text1(31).visible = (vParamAplic.TipoCaja2 <> "")
    
    'tara3
    Text1(15).Enabled = (vParamAplic.TipoCaja3 <> "")
    Text1(15).visible = (vParamAplic.TipoCaja3 <> "")
    Text1(20).Enabled = (vParamAplic.TipoCaja3 <> "")
    Text1(20).visible = (vParamAplic.TipoCaja3 <> "")
    Text1(27).Enabled = (vParamAplic.TipoCaja3 <> "")
    Text1(27).visible = (vParamAplic.TipoCaja3 <> "")
    Text1(32).Enabled = (vParamAplic.TipoCaja3 <> "")
    Text1(32).visible = (vParamAplic.TipoCaja3 <> "")
    
    'tara4
    Text1(16).Enabled = (vParamAplic.TipoCaja4 <> "")
    Text1(16).visible = (vParamAplic.TipoCaja4 <> "")
    Text1(8).Enabled = (vParamAplic.TipoCaja4 <> "")
    Text1(8).visible = (vParamAplic.TipoCaja4 <> "")
    Text1(28).Enabled = (vParamAplic.TipoCaja4 <> "")
    Text1(28).visible = (vParamAplic.TipoCaja4 <> "")
    Text1(33).Enabled = (vParamAplic.TipoCaja4 <> "")
    Text1(33).visible = (vParamAplic.TipoCaja4 <> "")
    
    'tara5
    Text1(17).Enabled = (vParamAplic.TipoCaja5 <> "")
    Text1(17).visible = (vParamAplic.TipoCaja5 <> "")
    Text1(9).Enabled = (vParamAplic.TipoCaja5 <> "")
    Text1(9).visible = (vParamAplic.TipoCaja5 <> "")
    Text1(29).Enabled = (vParamAplic.TipoCaja5 <> "")
    Text1(29).visible = (vParamAplic.TipoCaja5 <> "")
    Text1(34).Enabled = (vParamAplic.TipoCaja5 <> "")
    Text1(34).visible = (vParamAplic.TipoCaja5 <> "")
    
End Sub


Private Function HanModificadoCajas() As Boolean
Dim i As Integer
    HanModificadoCajas = False
    
    NumCajas = 0
'    For I = 13 To 17
'        If Text1(I).Text <> "" Then
'            NumCajas = NumCajas + CCur(ComprobarCero(Text1(I).Text))
'        End If
'    Next I
    If Text1(13).Text <> "" And vParamAplic.EsCaja1 Then NumCajas = NumCajas + CCur(ComprobarCero(Text1(13).Text))
    If Text1(14).Text <> "" And vParamAplic.EsCaja2 Then NumCajas = NumCajas + CCur(ComprobarCero(Text1(14).Text))
    If Text1(15).Text <> "" And vParamAplic.EsCaja3 Then NumCajas = NumCajas + CCur(ComprobarCero(Text1(15).Text))
    If Text1(16).Text <> "" And vParamAplic.EsCaja4 Then NumCajas = NumCajas + CCur(ComprobarCero(Text1(16).Text))
    If Text1(17).Text <> "" And vParamAplic.EsCaja5 Then NumCajas = NumCajas + CCur(ComprobarCero(Text1(17).Text))

    HanModificadoCajas = (NumCajas <> NumCajasAnt)

End Function

Private Function HanModificadoKilos() As Boolean
Dim i As Integer
    
    HanModificadoKilos = (DBLet(Text1(11).Text, "N") <> NumKilosAnt)

End Function



Private Sub CrearPaletizacion()
Dim Sql As String

    Sql = "delete from trzpalets where numnotac = " & Trim(Data1.Recordset!NumNotac)
    conn.Execute Sql
    
    mnPaletizacion_Click

End Sub

Private Sub ActualizarPaletizacion()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim SQL1 As String
Dim KilosTotal As Currency
Dim KilosNeto As Currency
Dim KilosLinea As Currency
Dim Numlineas As Currency
Dim IdPalet As Currency

    If vParamAplic.HayTraza = False Then Exit Sub
    
    Sql = "select numcajones, numkilos, idpalet from trzpalets where numnotac = " & Trim(Data1.Recordset!NumNotac)
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenDynamic, adLockReadOnly, adCmdText
    
    If Not Rs.EOF Then
        Rs.MoveFirst
        
        KilosNeto = DBLet(Data1.Recordset!KilosNet, "N")
'        NumCajas = DBLet(Data1.Recordset!numcajo1, "N") + _
'                   DBLet(Data1.Recordset!numcajo2, "N") + _
'                   DBLet(Data1.Recordset!numcajo3, "N") + _
'                   DBLet(Data1.Recordset!numcajo4, "N") + _
'                   DBLet(Data1.Recordset!numcajo5, "N")

        NumCajas = 0
        If vParamAplic.EsCaja1 Then NumCajas = NumCajas + DBLet(Data1.Recordset!numcajo1, "N")
        If vParamAplic.EsCaja2 Then NumCajas = NumCajas + DBLet(Data1.Recordset!numcajo2, "N")
        If vParamAplic.EsCaja3 Then NumCajas = NumCajas + DBLet(Data1.Recordset!numcajo3, "N")
        If vParamAplic.EsCaja4 Then NumCajas = NumCajas + DBLet(Data1.Recordset!numcajo4, "N")
        If vParamAplic.EsCaja5 Then NumCajas = NumCajas + DBLet(Data1.Recordset!numcajo5, "N")
        
        If NumCajas = 0 Then 'vamos por palots y debemos ver cuantos registros=palots tenemos
            SQL1 = "select count(*) from trzpalets where numnotac = " & Trim(Data1.Recordset!NumNotac)
            
            Numlineas = TotalRegistros(SQL1)
        End If
        
        KilosTotal = 0
        While Not Rs.EOF
            If NumCajas <> 0 Then ' estamos por palet
                KilosLinea = (KilosNeto * DBLet(Rs.Fields(0).Value, "N")) \ NumCajas
            Else ' estamos por palot
                KilosLinea = KilosNeto \ Numlineas
            End If
            
            SQL1 = "update trzpalets set numkilos = " & DBSet(KilosLinea, "N")
            SQL1 = SQL1 & " where idpalet = " & DBSet(Rs.Fields(2).Value, "N")
            
            conn.Execute SQL1
            
            KilosTotal = KilosTotal + KilosLinea
        
            IdPalet = DBLet(Rs.Fields(2).Value, "N")
            
            Rs.MoveNext
        Wend
        
        If KilosTotal <> KilosNeto Then ' en el ultimo registro metemos el restante
            SQL1 = "update trzpalets set numkilos = numkilos + " & DBSet(KilosNeto - KilosTotal, "N")
            SQL1 = SQL1 & " where idpalet = " & DBSet(IdPalet, "N")
            
            conn.Execute SQL1
        End If
    End If
    
    Set Rs = Nothing
        
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
    
    Dim Rs As ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    Dim Rs4 As ADODB.Recordset
    Dim Sql As String
    Dim Lin As String ' l�nea de impresi�n
    Dim i As Integer
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
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly
    If Not Rs.EOF Then
                '-- Impresi�n de la cabecera
'                Lin = "         1         2         3         4"
'                Printer.Print Lin
'                Lin = "1234567890123456789012345678901234567890"
'                Printer.Print Lin
    '    Lin = vEmpresa.nomempre
    
        Veces = Round2(CInt(Copias) / 2, 0)
    
        For i = 1 To Veces
    
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
            Lin = "Fecha   : " & Format(Rs!FechaEnt, "dd/mm/yyyy") & "  N." & Format(NumNota, "00000000") & " Hora:" & Format(Rs!horaentr, "hh:mm") & _
                   Space(2) & _
                  "Fecha   : " & Format(Rs!FechaEnt, "dd/mm/yyyy") & "  N." & Format(NumNota, "00000000") & " Hora:" & Format(Rs!horaentr, "hh:mm")
'                  1234567890                         1234567890      1234                     56789012      345678                         90123
            Printer.Print Lin
            
            ' LINEA 5
            If i = 1 Or i = 2 Then
                Lin = "Socio   : " & Rs!nomsocio
            Else
                Lin = ""
            End If
            Printer.Print Lin
            
            ' LINEA 6
            Partida = DevuelveValor("select nomparti from rcampos, rpartida where rcampos.codparti = rpartida.codparti and rcampos.codcampo = " & DBSet(Rs!codCampo, "N"))
            
            Lin = RellenaABlancos("Huerto  : " & Format(Rs!codCampo, "00000000") & "-" & Mid(Partida, 1, 24), True, 43) & _
                   Space(2) & _
                  RellenaABlancos("Huerto  : " & Format(Rs!codCampo, "00000000") & "-" & Mid(Partida, 1, 24), True, 43)
            Printer.Print Lin
'                  1234567890                         12345678      9    012345678901234567890123

            
            ' LINEA 7
            Situacion = ""
            Situacion = DevuelveValor("select nomsitua from rsituacioncampo, rcampos where rcampos.codsitua = rsituacioncampo.codsitua and rcampos.codsitua <> 0 and rcampos.codcampo = " & DBSet(Rs!codCampo, "N"))
            
            Lin = RellenaABlancos("Variedad: " & Format(Rs!codvarie, "0000") & " " & DBLet(Rs!nomvarie, "T") & " " & Situacion, True, 43) & _
                  Space(2) & _
                  RellenaABlancos("Variedad: " & Format(Rs!codvarie, "0000") & " " & DBLet(Rs!nomvarie, "T") & " " & Situacion, True, 43)
            Printer.Print Lin

            ' LINEA 8
            Clase = ""
            Clase = DevuelveValor("select nomclase from clases, variedades where variedades.codvarie = " & DBSet(Rs!codvarie, "N") & " and variedades.codclase = clases.codclase ")
            
            Lin = RellenaABlancos("Grupo   : " & Clase, True, 43) & Space(2) & RellenaABlancos("Grupo   : " & Clase, True, 43) & Space(2)
            Printer.Print Lin
            
            ' LINEA 9
            Lin = ""
            Printer.Print Lin
            
            ' LINEA 10
'            Cajas = DBLet(Rs!numcajo1, "N") + DBLet(Rs!numcajo2, "N") + DBLet(Rs!numcajo3, "N") + DBLet(Rs!numcajo4, "N") + DBLet(Rs!numcajo5, "N")
            cajas = 0
            If vParamAplic.EsCaja1 Then cajas = cajas + DBLet(Rs!numcajo1, "N")
            If vParamAplic.EsCaja2 Then cajas = cajas + DBLet(Rs!numcajo2, "N")
            If vParamAplic.EsCaja3 Then cajas = cajas + DBLet(Rs!numcajo3, "N")
            If vParamAplic.EsCaja4 Then cajas = cajas + DBLet(Rs!numcajo4, "N")
            If vParamAplic.EsCaja5 Then cajas = cajas + DBLet(Rs!numcajo5, "N")

            Tara = DBLet(Rs!taracaja1, "N") + DBLet(Rs!taracaja2, "N") + DBLet(Rs!taracaja3, "N") + DBLet(Rs!taracaja4, "N") + DBLet(Rs!taracaja5, "N") + DBLet(Rs!TaraVehi, "N")
            
            
            Lin = RellenaABlancos("Nro.Cajas : " & RellenaABlancos(Format(cajas, "###,##0"), False, 6) & "    " & "Total Tara: " & RellenaABlancos(Format(Tara, "###,##0"), False, 6), True, 43) & _
                  Space(2) & _
                  RellenaABlancos("Nro.Cajas : " & RellenaABlancos(Format(cajas, "###,##0"), False, 6) & "    " & "Total Tara: " & RellenaABlancos(Format(Tara, "###,##0"), False, 6), True, 43)

            Printer.Print Lin
            
            ' LINEA 11
            Lin = RellenaABlancos("Peso Bruto: " & RellenaABlancos(Format(Rs!KilosBru, "###,##0"), False, 6) & "    " & "Peso Neto : " & RellenaABlancos(Format(Rs!KilosNet, "###,##0"), False, 6), True, 43) & _
                  Space(2) & _
                  RellenaABlancos("Peso Bruto: " & RellenaABlancos(Format(Rs!KilosBru, "###,##0"), False, 6) & "    " & "Peso Neto : " & RellenaABlancos(Format(Rs!KilosNet, "###,##0"), False, 6), True, 43)
            Printer.Print Lin
            
            
            Lin = ""
            Printer.Print Lin
'            Printer.Print Lin
                
        Next i
                
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
    
    Dim Rs As ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    Dim Rs4 As ADODB.Recordset
    Dim Sql As String
    Dim Lin As String ' l�nea de impresi�n
    Dim i As Integer
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
        'Open "d:\t1.txt" For Output As #NF
        Open "LPT1" For Output As #NF
            
        For i = 1 To Lineas.Count
            Print #NF, Lineas(i)
        Next i
        
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
Dim Rs As ADODB.Recordset
Dim Lin As String ' l�nea de impresi�n
Dim i As Integer
Dim N As Integer
Dim ImporteIva As Currency
Dim EnEfectivo As Boolean
    
Dim Veces As Integer
Dim Partida As String
Dim Situacion As String
Dim Clase As String
Dim Tara As Currency
Dim cajas As Currency
    
    On Error GoTo eCargarLineas
    
    CargarLineas = False
    
    
    Sql = "select rentradas.*, rsocios.nomsocio, variedades.nomvarie from rentradas, rsocios, variedades "
    Sql = Sql & " where numnotac = " & DBSet(NumNota, "N")
    Sql = Sql & " and rentradas.codsocio = rsocios.codsocio "
    Sql = Sql & " and rentradas.codvarie = variedades.codvarie "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly
    If Not Rs.EOF Then
    
        Veces = Round2(CInt(Copias) / 2, 0)
    
        For i = 1 To Veces
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
            Lin = "Fecha   : " & Format(Rs!FechaEnt, "dd/mm/yyyy") & "  N." & Format(NumNota, "00000000") & " Hora:" & Format(Rs!horaentr, "hh:mm") & _
                   Space(2) & _
                  "Fecha   : " & Format(Rs!FechaEnt, "dd/mm/yyyy") & "  N." & Format(NumNota, "00000000") & " Hora:" & Format(Rs!horaentr, "hh:mm")
'                  1234567890                         1234567890      1234                     56789012      345678                         90123
            Lineas.Add Lin
            
            ' LINEA 5
            If i = 1 Or i = 2 Then
                Lin = "Socio   : " & Rs!nomsocio
            Else
                Lin = ""
            End If
            Lineas.Add Lin
            
            ' LINEA 6
            Partida = DevuelveValor("select nomparti from rcampos, rpartida where rcampos.codparti = rpartida.codparti and rcampos.codcampo = " & DBSet(Rs!codCampo, "N"))
            
            Lin = RellenaABlancos("Huerto  : " & Format(Rs!codCampo, "00000000") & "-" & Mid(Partida, 1, 24), True, 43) & _
                   Space(2) & _
                  RellenaABlancos("Huerto  : " & Format(Rs!codCampo, "00000000") & "-" & Mid(Partida, 1, 24), True, 43)
'                  1234567890                         12345678      9    012345678901234567890123
            Lineas.Add Lin

            
            ' LINEA 7
            Situacion = ""
            Situacion = DevuelveValor("select nomsitua from rsituacioncampo, rcampos where rcampos.codsitua = rsituacioncampo.codsitua and rcampos.codsitua <> 0 and rcampos.codcampo = " & DBSet(Rs!codCampo, "N"))
            
            Lin = RellenaABlancos("Variedad: " & Format(Rs!codvarie, "0000") & " " & DBLet(Rs!nomvarie, "T") & " " & Situacion, True, 43) & _
                  Space(2) & _
                  RellenaABlancos("Variedad: " & Format(Rs!codvarie, "0000") & " " & DBLet(Rs!nomvarie, "T") & " " & Situacion, True, 43)
            Lineas.Add Lin

            ' LINEA 8
            Clase = ""
            Clase = DevuelveValor("select nomclase from clases, variedades where variedades.codvarie = " & DBSet(Rs!codvarie, "N") & " and variedades.codclase = clases.codclase ")
            
            Lin = RellenaABlancos("Grupo   : " & Clase, True, 43) & Space(2) & RellenaABlancos("Grupo   : " & Clase, True, 43) & Space(2)
            Lineas.Add Lin
            
            ' LINEA 9
            Lin = ""
            Lineas.Add Lin
            
            ' LINEA 10
            'Cajas = DBLet(Rs!numcajo1, "N") + DBLet(Rs!numcajo2, "N") + DBLet(Rs!numcajo3, "N") + DBLet(Rs!numcajo4, "N") + DBLet(Rs!numcajo5, "N")
            cajas = 0
            If vParamAplic.EsCaja1 Then cajas = cajas + DBLet(Rs!numcajo1, "N")
            If vParamAplic.EsCaja2 Then cajas = cajas + DBLet(Rs!numcajo2, "N")
            If vParamAplic.EsCaja3 Then cajas = cajas + DBLet(Rs!numcajo3, "N")
            If vParamAplic.EsCaja4 Then cajas = cajas + DBLet(Rs!numcajo4, "N")
            If vParamAplic.EsCaja5 Then cajas = cajas + DBLet(Rs!numcajo5, "N")
            
            Tara = DBLet(Rs!taracaja1, "N") + DBLet(Rs!taracaja2, "N") + DBLet(Rs!taracaja3, "N") + DBLet(Rs!taracaja4, "N") + DBLet(Rs!taracaja5, "N") + DBLet(Rs!TaraVehi, "N")
            
            Lin = RellenaABlancos("Nro.Cajas : " & RellenaABlancos(Format(cajas, "###,##0"), False, 6) & "    " & "Total Tara: " & RellenaABlancos(Format(Tara, "###,##0"), False, 6), True, 43) & _
                  Space(2) & _
                  RellenaABlancos("Nro.Cajas : " & RellenaABlancos(Format(cajas, "###,##0"), False, 6) & "    " & "Total Tara: " & RellenaABlancos(Format(Tara, "###,##0"), False, 6), True, 43)
            Lineas.Add Lin

            
            ' LINEA 11
            Lin = RellenaABlancos("Peso Bruto: " & RellenaABlancos(Format(Rs!KilosBru, "###,##0"), False, 6) & "    " & "Peso Neto : " & RellenaABlancos(Format(Rs!KilosNet, "###,##0"), False, 6), True, 43) & _
                  Space(2) & _
                  RellenaABlancos("Peso Bruto: " & RellenaABlancos(Format(Rs!KilosBru, "###,##0"), False, 6) & "    " & "Peso Neto : " & RellenaABlancos(Format(Rs!KilosNet, "###,##0"), False, 6), True, 43)
            Lineas.Add Lin
            
            
            Lin = ""
            Lineas.Add Lin
'            Lineas.Add Lin
                
        Next i
    End If

    CargarLineas = True
    Exit Function
    
eCargarLineas:
    MuestraError Err.Number, "Cargando las lineas de impresi�n:", Err.Description
End Function




Private Function ActualizarChivato(Mens As String, Operacion As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Sql2 As String
Dim Rs1 As ADODB.Recordset
Dim cadena As String
Dim Producto As String
Dim NumF As String

    On Error GoTo eActualizarChivato

    ActualizarChivato = False
    
    Sql = "select codvarie, numcajo1, numnotac, codsocio, codcampo, codcapat, codtarif, "
    Sql = Sql & "kilosbru, kilosnet, tipoentr, fechaent, codtrans, nropesada, numcajo2, numcajo3, numcajo4, numcajo5, zona, altura "
    Sql = Sql & "from rentradas"
    Sql = Sql & " where numnotac = " & DBSet(Text1(0).Text, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    If Not Rs.EOF Then
        '[Monica]26/09/2011: cuando la variedad es de almazara no he de actualizar chivato
        If EsVariedadGrupo5(CStr(DBLet(Rs!codvarie, "N"))) Then
            ActualizarChivato = True
            Set Rs = Nothing
            Exit Function
        End If
        
        Producto = DevuelveValor("select codprodu from variedades where codvarie = " & DBSet(Rs!codvarie, "N"))
        
        cadena = v_cadena & "<ROW notacamp=" & """" & Format(DBLet(Rs!NumNotac, "N"), "######0") & """"
        cadena = cadena & " fechaent=" & """" & Format(Rs!FechaEnt, "yyyymmdd") & """"
        cadena = cadena & " codprodu=" & """" & Format(DBLet(Producto, "N"), "#####0") & """"
        cadena = cadena & " codvarie=" & """" & Format(DBLet(Rs!codvarie, "N"), "#####0") & """"
        cadena = cadena & " codsocio=" & """" & Format(DBLet(Rs!Codsocio, "N"), "#####0") & """"
        cadena = cadena & " codcampo=" & """" & Format(DBLet(Rs!codCampo, "N"), "#######0") & """"
        cadena = cadena & " kilosbru=" & """" & Format(DBLet(Rs!KilosBru, "N"), "###0") & """"
        cadena = cadena & " kilosnet=" & """" & Format(DBLet(Rs!KilosNet, "N"), "###0") & """"
        cadena = cadena & " numcajo1=" & """" & Format(DBLet(Rs!numcajo1, "N"), "##0") & """"
        cadena = cadena & " numcajo2=" & """" & Format(DBLet(Rs!numcajo2, "N"), "##0") & """"
        cadena = cadena & " numcajo3=" & """" & Format(DBLet(Rs!numcajo3, "N"), "##0") & """"
        cadena = cadena & " numcajo4=" & """" & Format(DBLet(Rs!numcajo4, "N"), "##0") & """"
        cadena = cadena & " numcajo5=" & """" & Format(DBLet(Rs!numcajo5, "N"), "##0") & """"
        cadena = cadena & " matricul=" & """" & DBLet(Rs!codTrans, "T") & """"
        cadena = cadena & " codcapat=" & """" & Format(DBLet(Rs!codcapat, "N"), "###0") & """"
        cadena = cadena & " identifi=" & """" & Format(0, "#####0") & """"
        cadena = cadena & " altura=" & """" & Format(DBLet(Rs!altura, "N"), "##0") & """"
        cadena = cadena & " zona=" & """" & Format(DBLet(Rs!Zona, "N"), "#########0") & """"
        cadena = cadena & " /></ROWDATA></DATAPACKET>"
    
            
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
        Sql = Sql & DBSet(Rs!NumNotac, "N") & ","
        Sql = Sql & DBSet(Rs!NumNotac, "N") & ","
        Sql = Sql & DBSet(cadena, "T") & ","
        Sql = Sql & ValorNulo & ","
        Sql = Sql & ValorNulo & ","
        Sql = Sql & "'" & Format(Now, "hh:mm:ss") & "',"
        Sql = Sql & ValorNulo & ","
        Sql = Sql & ValorNulo & ")"
        
        conn.Execute Sql
            
    End If
    
    Set Rs = Nothing
    
    ActualizarChivato = True
    Exit Function
    
eActualizarChivato:
    Mens = Mens & Err.Description
End Function


Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 ' tarar tractor
            mnTararTractor_Click
        Case 2 ' Pendiente de Tarar
            mnPendiente_Click
            'mnPaletizacion_Click
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

