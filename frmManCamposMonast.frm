VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmManCamposMonast 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Propiedades"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   8700
   Icon            =   "frmManCamposMonast.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   8700
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
      Index           =   0
      Left            =   6345
      TabIndex        =   32
      Top             =   270
      Width           =   1605
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3780
      TabIndex        =   30
      Top             =   30
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   31
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
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   150
      TabIndex        =   28
      Top             =   30
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   29
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
   Begin VB.Frame Frame2 
      Height          =   900
      Index           =   0
      Left            =   150
      TabIndex        =   13
      Top             =   780
      Width           =   8370
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Left            =   1350
         MaxLength       =   8
         TabIndex        =   0
         Tag             =   "C�digo Campo|N|N|1|9999|rcampos|codcampo|0000|S|"
         Top             =   315
         Width           =   960
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
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   0
         Left            =   225
         TabIndex        =   14
         Top             =   360
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   135
      TabIndex        =   11
      Top             =   7650
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
         TabIndex        =   12
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
      Left            =   7470
      TabIndex        =   10
      Top             =   7740
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
      Left            =   6300
      TabIndex        =   9
      Top             =   7740
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
      Left            =   7470
      TabIndex        =   15
      Top             =   7740
      Visible         =   0   'False
      Width           =   1065
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5835
      Left            =   150
      TabIndex        =   16
      Top             =   1755
      Width           =   8370
      _ExtentX        =   14764
      _ExtentY        =   10292
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   11
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
      TabPicture(0)   =   "frmManCamposMonast.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label28"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "imgZoom(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label29"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "imgBuscar(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "imgBuscar(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text1(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text1(8)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "FrameDatosDtoAdministracion"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text5(3)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text1(6)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text1(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text2(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text4(3)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text1(1)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text2(1)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
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
         Left            =   2325
         MaxLength       =   40
         TabIndex        =   35
         Top             =   495
         Width           =   5625
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
         Left            =   1305
         MaxLength       =   6
         TabIndex        =   1
         Tag             =   "C�digo Socio|N|N|1|999999|rcampos|codsocio|000000|N|"
         Top             =   495
         Width           =   975
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
         Left            =   1305
         MaxLength       =   30
         TabIndex        =   33
         Top             =   1485
         Width           =   960
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
         Left            =   2325
         MaxLength       =   30
         TabIndex        =   24
         Top             =   1005
         Width           =   5625
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
         Left            =   1290
         MaxLength       =   4
         TabIndex        =   2
         Tag             =   "Partida|N|N|1|9999|rcampos|codparti|0000||"
         Top             =   1005
         Width           =   990
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
         Height          =   1920
         Index           =   6
         Left            =   180
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Tag             =   "Observaciones|T|S|||rcampos|observac|||"
         Top             =   3555
         Width           =   7755
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
         Left            =   2325
         MaxLength       =   30
         TabIndex        =   20
         Top             =   1485
         Width           =   5625
      End
      Begin VB.Frame FrameDatosDtoAdministracion 
         Caption         =   "Datos Administrativos"
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
         Height          =   1185
         Left            =   180
         TabIndex        =   17
         Top             =   1935
         Width           =   7755
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
            Index           =   4
            Left            =   1710
            MaxLength       =   10
            TabIndex        =   3
            Tag             =   "Fecha Alta|F|N|||rcampos|fecaltas|dd/mm/yyyy||"
            Top             =   330
            Width           =   1300
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
            Left            =   5190
            MaxLength       =   10
            TabIndex        =   4
            Tag             =   "Fecha Baja|F|S|||rcampos|fecbajas|dd/mm/yyyy||"
            Top             =   330
            Width           =   1300
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
            Left            =   2490
            TabIndex        =   18
            Top             =   720
            Width           =   4020
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
            Left            =   1710
            MaxLength       =   2
            TabIndex        =   5
            Tag             =   "C�digo Situacion|N|N|0|99|rcampos|codsitua|00||"
            Top             =   720
            Width           =   765
         End
         Begin VB.CheckBox chkAbonos 
            Caption         =   "Entrega Ficha Cultivo"
            Height          =   315
            Index           =   1
            Left            =   3690
            TabIndex        =   8
            Top             =   750
            Width           =   2445
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
            Left            =   150
            TabIndex        =   26
            Top             =   330
            Width           =   1245
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
            Left            =   3750
            TabIndex        =   25
            Top             =   330
            Width           =   1125
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   0
            Left            =   1410
            Picture         =   "frmManCamposMonast.frx":0028
            ToolTipText     =   "Buscar fecha"
            Top             =   330
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   1
            Left            =   4890
            Picture         =   "frmManCamposMonast.frx":00B3
            ToolTipText     =   "Buscar fecha"
            Top             =   330
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
            Left            =   150
            TabIndex        =   19
            Top             =   720
            Width           =   945
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   0
            Left            =   1410
            ToolTipText     =   "Buscar Situaci�n"
            Top             =   720
            Width           =   240
         End
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
         Left            =   1350
         MaxLength       =   4
         TabIndex        =   34
         Tag             =   "Zonas|N|S|||rcampos|codzonas|0000||"
         Top             =   1005
         Width           =   855
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
         Left            =   1305
         MaxLength       =   6
         TabIndex        =   7
         Tag             =   "C�digo Propietario|N|S|||rcampos|codpropiet|000000|N|"
         Top             =   495
         Width           =   975
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   990
         ToolTipText     =   "Buscar Socio"
         Top             =   495
         Width           =   240
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
         Left            =   180
         TabIndex        =   36
         Top             =   495
         Width           =   660
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   990
         ToolTipText     =   "Buscar Partida"
         Top             =   1020
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Calle"
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
         Left            =   180
         TabIndex        =   23
         Top             =   1005
         Width           =   735
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
         Left            =   255
         TabIndex        =   22
         Top             =   3195
         Width           =   1530
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   1815
         ToolTipText     =   "Zoom descripci�n"
         Top             =   3195
         Width           =   240
      End
      Begin VB.Label Label28 
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
         Left            =   180
         TabIndex        =   21
         Top             =   1485
         Width           =   1035
      End
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   8100
      TabIndex        =   27
      Top             =   225
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3105
      Top             =   7335
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
         Caption         =   "Asignaci�n GlobalGap"
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
Attribute VB_Name = "frmManCamposMonast"
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

'Dim T1 As Single
Private Const IdPrograma = 2021


Public DatosADevolverBusqueda As String    'Tindr� el n� de text que vol que torne, empipat
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
Private WithEvents frmCamPrev As frmManCamposMonastPrev ' campos vista previa
Attribute frmCamPrev.VB_VarHelpID = -1

Private WithEvents frmPar As frmManPartidas 'partidas
Attribute frmPar.VB_VarHelpID = -1
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
        Case 1  'B�SQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOK Then
                Text1(7).Text = Text1(1).Text
                Text1(8).Text = 1
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
            If DatosOK Then
                Modificar
'                If ModificaDesdeFormulario2(Me, 1) Then
'                    TerminaBloquear
'                    PosicionarData
'                    CargaGrid 1, True
'                End If
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



'Private Sub cmdAux_Click(Index As Integer)
'    TerminaBloquear
'    Select Case Index
'        Case 1 'Calidades de la variedad de cabecera
'            Set frmCalid = New frmManCalidades
'            frmCalid.DatosADevolverBusqueda = "0|1|2|3|"
'            frmCalid.CodigoActual = txtAux1(1).Text
'            frmCalid.ParamVariedad = txtAux1(2).Text
'            frmCalid.Show vbModal
'            Set frmCalid = Nothing
'            PonerFoco txtAux1(1)
'
'        Case 0 ' Socios coopropietarios
'            Set frmSoc1 = New frmManSocios
'            frmSoc1.DatosADevolverBusqueda = "0|1|"
'            frmSoc1.Show vbModal
'            Set frmSoc1 = Nothing
'            PonerFoco txtAux3(2)
'
'        Case 2 ' Incidencias
'            Set frmInc = New frmManInciden
'            frmInc.DatosADevolverBusqueda = "0|1|"
'            frmInc.Show vbModal
'            Set frmInc = Nothing
'            PonerFoco txtAux5(3)
'
'        Case 3, 4 ' fecha de incidencia de agroseguro
'           Screen.MousePointer = vbHourglass
'
'           Dim esq As Long
'           Dim dalt As Long
'           Dim menu As Long
'           Dim obj As Object
'
'           Set frmC2 = New frmCal
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
'           frmC2.Left = esq + cmdAux(Index).Parent.Left + 30
'           frmC2.Top = dalt + cmdAux(Index).Parent.Top + cmdAux(Index).Height + menu - 40
'
'           frmC2.NovaData = Now
'           Select Case Index
'                Case 3
'                    Indice = 2
'                Case 4
'                    Indice = 6
'           End Select
'
'           Me.cmdAux(0).Tag = Indice
'
'           PonerFormatoFecha txtAux5(Indice)
'           If txtAux5(Indice).Text <> "" Then frmC2.NovaData = CDate(txtAux5(Indice).Text)
'
'           Screen.MousePointer = vbDefault
'           frmC2.Show vbModal
'           Set frmC2 = Nothing
'           PonerFoco txtAux5(Indice)
'
'        Case 9 ' concepto de gasto
'            Set frmGto = New frmManConcepGasto
'            frmGto.DatosADevolverBusqueda = "0|1|"
'            frmGto.Show vbModal
'            Set frmGto = Nothing
'            PonerFoco txtaux7(2)
'
'        Case 10 ' fecha de concepto de gasto
'           Screen.MousePointer = vbHourglass
'
'           Set frmC3 = New frmCal
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
'           frmC3.Left = esq + cmdAux(Index).Parent.Left + 30
'           frmC3.Top = dalt + cmdAux(Index).Parent.Top + cmdAux(Index).Height + menu - 40
'
'           frmC3.NovaData = Now
'
'           Indice = 3
'
'           Me.cmdAux(0).Tag = Indice
'
'           PonerFormatoFecha txtaux7(Indice)
'           If txtaux7(Indice).Text <> "" Then frmC3.NovaData = CDate(txtaux7(Indice).Text)
'
'           Screen.MousePointer = vbDefault
'           frmC3.Show vbModal
'           Set frmC3 = Nothing
'           PonerFoco txtaux7(Indice)
'
'
'        Case 11 ' fecha de impresion de orden de confeccion
'           Screen.MousePointer = vbHourglass
'
'           Set frmC4 = New frmCal
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
'           frmC4.Left = esq + cmdAux(Index).Parent.Left + 30
'           frmC4.Top = dalt + cmdAux(Index).Parent.Top + cmdAux(Index).Height + menu - 40
'
'           frmC4.NovaData = Now
'
'           Indice = 2
'
'           Me.cmdAux(0).Tag = Indice
'
'           PonerFormatoFecha txtaux8(Indice)
'           If txtaux8(Indice).Text <> "" Then frmC4.NovaData = CDate(txtaux8(Indice).Text)
'
'           Screen.MousePointer = vbDefault
'           frmC4.Show vbModal
'           Set frmC4 = Nothing
'           PonerFoco txtaux8(Indice)
'
'        Case 12 ' fecha de revision del campo
'           Screen.MousePointer = vbHourglass
'
'           Set frmC5 = New frmCal
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
'           frmC5.Left = esq + cmdAux(Index).Parent.Left + 30
'           frmC5.Top = dalt + cmdAux(Index).Parent.Top + cmdAux(Index).Height + menu - 40
'
'           frmC5.NovaData = Now
'
'           Indice = 2
'
'           Me.cmdAux(0).Tag = Indice
'
'           PonerFormatoFecha txtaux8(Indice)
'           If txtAux9(Indice).Text <> "" Then frmC5.NovaData = CDate(txtAux9(Indice).Text)
'
'           Screen.MousePointer = vbDefault
'           frmC5.Show vbModal
'           Set frmC5 = Nothing
'           PonerFoco txtAux9(Indice)
'
'
'    End Select
'
'    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
'End Sub

' *** si n'hi han combos a la cap�alera ***
'Private Sub Combo1_GotFocus(Index As Integer)
'    If Modo = 1 Then Combo1(Index).BackColor = vbLightBlue
'End Sub
'
'Private Sub Combo1_LostFocus(Index As Integer)
'    If Combo1(Index).BackColor = vbLightBlue Then Combo1(Index).BackColor = vbWhite
'End Sub
'
'Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
'    KEYpress KeyAscii
'End Sub

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
    '[Monica]03/10/2011: a�adido el modo = 3 para solucionar problema de Picassent
    If Modo = 3 Or Modo = 4 Or Modo = 5 Then TerminaBloquear
    
    Set dbAriagro = Nothing

    '[Monica]28/11/2011: cliente de ariges
    If vParamAplic.BDAriges <> "" Then CerrarConexionAriges
End Sub

Private Sub Form_Load()
Dim I As Integer
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    
    ' ICONETS DE LA BARRA
    btnPrimero = 23 'index del bot� "primero"
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
        'el 13 i el 14 son separadors
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
'    For I = 0 To ToolAux.Count - 1
'        With Me.ToolAux(I)
'            .HotImageList = frmPpal.imgListComun_OM16
'            .DisabledImageList = frmPpal.imgListComun_BN16
'            .ImageList = frmPpal.imgListComun16
'            .Buttons(1).Image = 3   'Insertar
'            .Buttons(2).Image = 4   'Modificar
'            .Buttons(3).Image = 5   'Borrar
'        End With
'
'        If I = 5 Then ' boton de contabilizar un gasto de campo
'            With Me.ToolAux(I)
'                .HotImageList = frmPpal.imgListComun_OM16
'                .DisabledImageList = frmPpal.imgListComun_BN16
'                .ImageList = frmPpal.imgListComun16
'                .Buttons(4).Image = 13   'Contabilizar
'            End With
'        End If
'
'        If I = 7 Then
'            With Me.ToolAux(I)
'                .HotImageList = frmPpal.imgListComun_OM16
'                .DisabledImageList = frmPpal.imgListComun_BN16
'                .ImageList = frmPpal.imgListComun16
'                .Buttons(4).Image = 10   'Impresion de revisiones de campos
'            End With
'        End If
'    Next I
    ' ***********************************
    '[Monica]03/02/2015: solo para el caso de eescalona ponemos Arrendador
    If vParamAplic.Cooperativa = 10 Then
        Label4.Caption = "Arrendador"
        Text1(1).Text = "C�digo Arrendador|N|N|1|999999|rcampos|codsocio|000000|N|"
    End If
    
    
'    Me.imgDoc(1).Picture = frmPpal.imgListPpal.ListImages(24).Picture
'    Me.imgDoc(1).visible = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16)
'    Me.imgDoc(1).Enabled = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16)
    
    
    
    
    'cargar IMAGES de busqueda
    For I = 0 To 2
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    
'    'carga IMAGES de mail
'    For i = 0 To Me.imgMail.Count - 1
'        Me.imgMail(i).Picture = frmPpal.imgListImages16.ListImages(2).Picture
'    Next i
    'IMAGES para zoom
    For I = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(I).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next I
    
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
'    ' ******* si n'hi han ll�nies *******
'    DataGridAux(0).ClearFields
'    DataGridAux(1).ClearFields
    
    '*** canviar el nom de la taula i l'ordenaci� de la cap�alera ***
    NombreTabla = "rcampos"
    Ordenacion = " ORDER BY codcampo"
    '************************************************
    
    'Mirem com est� guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    CadenaConsulta = "Select * from " & NombreTabla
    
    If NroCampo <> "" Then
        CadenaConsulta = CadenaConsulta & " where codcampo = " & DBSet(NroCampo, "N")
    Else
        CadenaConsulta = CadenaConsulta & " where codcampo = -1 "
    End If
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    '***** canviar el nom de la PK de la cap�alera; repasar codEmpre *************
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
       
    
    ModoLineas = 0
       
         
    '[Monica]14/02/2013: Totales de parcelas solo para Picassent
'    Label43.visible = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16)
'    For I = 10 To 12
'        txtAux2(I).visible = (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16)
'    Next I
         
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
    
    ' Para el chivato
    Set dbAriagro = New BaseDatos
    dbAriagro.abrir_MYSQL vConfig.SERVER, vUsu.CadenaConexion, vConfig.User, vConfig.password
    
    
    '[Monica]23/09/2014: en el caso de alzira el campo poda lo usaran para indicar si el campo est� sin Placa Identificativa
    If vParamAplic.Cooperativa = 4 Then
        chkAbonos(3).Tag = "Sin Placa Identif.|N|N|||rcampos|conpoda||N|"
        chkAbonos(3).Caption = "Sin Placa Identif."
    End If
    

    
End Sub


Private Sub LimpiarCampos()
Dim I As Integer

    On Error Resume Next
    
    limpiar Me   'M�tode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
    ' *** si n'hi han combos a la cap�alera ***
    ' *****************************************
'    For I = 0 To Combo1.Count - 1
'        Combo1(I).ListIndex = -1
'    Next I
    For I = 0 To chkAbonos.Count - 1
        Me.chkAbonos(I).Value = 0
    Next I
'    Me.chkAux(0).Value = 0

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
Dim I As Integer, NumReg As Byte
Dim B As Boolean

    On Error GoTo EPonerModo
 
 
 
    'Actualisa Iconos Insertar,Modificar,Eliminar
    'ActualizarToolbar Modo, Kmodo
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo, ModoLineas
    
    BuscaChekc = ""
       
    'Modo 2. N'hi han datos i estam visualisant-los
    '=========================================
    'Posem visible, si es formulari de b�squeda, el bot� "Regresar" quan n'hi han datos
    If DatosADevolverBusqueda <> "" Or NroCampo <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    
    '=======================================
    B = (Modo = 2)
    'Posar Fleches de desplasament visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Nom�s es per a saber que n'hi ha + d'1 registre
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    DesplazamientoVisible B And Data1.Recordset.RecordCount > 1
    
    '---------------------------------------------
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = B
    cmdAceptar.visible = B
       
    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a m�s neteja els camps Text1
    BloquearText1 Me, Modo
    BloquearCombo Me, Modo
    For I = 0 To 7
        BloquearChk Me.chkAbonos(I), (Modo = 0 Or Modo = 2 Or Modo = 5)
    Next I
    
    
    ' ***** bloquejar tots els controls visibles de la clau primaria de la cap�alera ***
'    If Modo = 4 Then _
'        BloquearTxt Text1(0), True 'si estic en  modificar, bloqueja la clau primaria
    ' **********************************************************************************
    
    ' **** si n'hi han imagens de buscar en la cap�alera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    For I = 0 To imgFec.Count - 1
        BloquearImgFec Me, I, Modo
    Next I
'    BloquearImgFec Me, 25, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
    ' ********************************************************
    ' *** si n'hi han ll�nies i imagens de buscar que no estiguen als grids ******
    'Ll�nies Departaments
    B = (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) And (NumTabMto = 0))
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
        CargaGrid 3, False
        CargaGrid 4, False
        CargaGrid 5, False
        CargaGrid 7, False
        '[Monica]30/09/2013
        'CargaGrid 6, False
    End If
    
    B = (Modo = 4) Or (Modo = 2)
    

'    DataGridAux(0).Enabled = B
'    DataGridAux(1).Enabled = B
'    DataGridAux(2).Enabled = B
'    DataGridAux(3).Enabled = B
'    DataGridAux(4).Enabled = B
'    DataGridAux(5).Enabled = B
'    DataGridAux(7).Enabled = B
    '[Monica]30/09/2013
    'DataGridAux(6).Enabled = b
    
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
    B = (Modo = 5) And (NumTabMto = 1) 'And (ModoLineas <> 3)
'    For I = 1 To txtAux1.Count - 1
'        BloquearTxt txtAux1(I), Not B
'    Next I
    B = (Modo = 5) And (NumTabMto = 1) And ModoLineas = 2
'    BloquearTxt txtAux1(1), B
'    BloquearBtn cmdAux(1), B
     '-----------------------------
     
    PonerModoOpcionesMenu (Modo) 'Activar opcions men� seg�n modo
    PonerOpcionesMenu   'Activar opcions de men� seg�n nivell
                        'de permisos de l'usuari


    ' bloqueo de todos los datos excepto de datos tecnicos cuando no es administrador y estamos modificando
    B = (Modo = 4) And vUsu.Nivel > 1
    
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
Dim B As Boolean, bAux As Boolean
Dim I As Byte
    
    'Barra de CAP�ALERA
    '------------------------------------------
    'b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    B = (Modo = 2 Or Modo = 0) And NroCampo = ""
    'Buscar
    Toolbar1.Buttons(5).Enabled = B
    Me.mnBuscar.Enabled = B
    'Vore Tots
    Toolbar1.Buttons(6).Enabled = B
    Me.mnVerTodos.Enabled = B
    
    'Insertar
    Toolbar1.Buttons(1).Enabled = B
    Me.mnNuevo.Enabled = B
    
    B = (Modo = 2 And Data1.Recordset.RecordCount > 0 And NroCampo = "") 'And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    Me.mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(3).Enabled = B
    Me.mnEliminar.Enabled = B
    
    
'    'Verificacion de errores
'    Toolbar2.Buttons(1).Enabled = B
'    Me.mnVerificacionErr.Enabled = B
'    'Sigpac
'    Toolbar2.Buttons(2).Enabled = B
'    Me.mnSigpac.Enabled = B
'    'Goolzoom
'    Toolbar2.Buttons(3).Enabled = B
'    Me.mnGoolzoom.Enabled = B
'
'    'Chequeo del Nro de Orden
'    Toolbar2.Buttons(4).Enabled = (Modo = 0 Or Modo = 2) And vParamAplic.Cooperativa = 4
'    Me.mnChequeoNroOrden.Enabled = (Modo = 0 Or Modo = 2) And vParamAplic.Cooperativa = 4
'
'    'Cambio de socio de un campo
'    Toolbar2.Buttons(5).Enabled = B
'    Me.mnCambioSocio.Enabled = B
'
'    'Gastos Pendientes de Integrar
'    Toolbar2.Buttons(6).Enabled = B
'    Me.mnGastosCampos.Enabled = B
'
'    '[Monica]10/11/2015. nuevo punto de menu de recalculo de globalgap
'    Toolbar2.Buttons(7).Enabled = (Modo = 0 Or Modo = 2)
'    Me.mnGlobalGap.Enabled = (Modo = 0 Or Modo = 2)
    
    'Imprimir
    Toolbar1.Buttons(8).Enabled = (Modo = 2 And Data1.Recordset.RecordCount > 0)
    Me.mnImprimir.Enabled = (Modo = 2 And Data1.Recordset.RecordCount > 0)
    
'    '[Monica]14/02/2013: Actualizacion de las superficies solo para Picassent
'    imgDoc(1).visible = (Modo = 2 And Data1.Recordset.RecordCount > 0 And (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16))
'    imgDoc(1).Enabled = (Modo = 2 And Data1.Recordset.RecordCount > 0 And (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16))
    
    ' *** si n'hi han ll�nies que tenen grids (en o sense tab) ***
    B = (Modo = 3 Or Modo = 4 Or Modo = 2 And NroCampo = "")
'    For I = 0 To ToolAux.Count - 1 '[Monica]30/09/2013: antes - 1
'        If I <> 6 Then
'            ToolAux(I).Buttons(1).Enabled = B
'            If B Then bAux = (B And Me.Adoaux(I).Recordset.RecordCount > 0)
'            ToolAux(I).Buttons(2).Enabled = bAux
'            ToolAux(I).Buttons(3).Enabled = bAux
'        End If
'    Next I
    
'    ToolAux(4).Buttons(1).Enabled = B And vUsu.Login = "root"
'    If B Then bAux = (B And Me.Adoaux(4).Recordset.RecordCount > 0)
'    ToolAux(4).Buttons(2).Enabled = bAux And vUsu.Login = "root"
'    ToolAux(4).Buttons(3).Enabled = bAux And vUsu.Login = "root"
    
    ' boton de integracion contable
'    bAux = B And Me.Adoaux(5).Recordset.RecordCount > 0
'    If Me.Adoaux(5).Recordset.RecordCount > 0 Then
'        bAux = bAux And CInt(Adoaux(5).Recordset.Fields(6).Value) = 0
'    End If
        
'    ToolAux(5).Buttons(4).Enabled = bAux
    
    ' boton de impresion de revisiones de campos
'    ToolAux(7).Buttons(4).Enabled = True
    
    
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
            tabla = "rcampos_clasif"
            Sql = "SELECT rcampos_clasif.codcampo, rcampos_clasif.codvarie, rcampos_clasif.codcalid, rcalidad.nomcalid, rcampos_clasif.muestra "
            Sql = Sql & " FROM " & tabla & " INNER JOIN rcalidad ON rcampos_clasif.codvarie = rcalidad.codvarie "
            Sql = Sql & " and rcampos_clasif.codcalid = rcalidad.codcalid "
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE rcampos_clasif.codcampo = -1"
            End If
            Sql = Sql
            Sql = Sql & " ORDER BY " & tabla & ".codcalid "
            
       Case 2 ' parcelas
            tabla = "rcampos_parcelas"
            Sql = "SELECT rcampos_parcelas.codcampo, rcampos_parcelas.numlinea, rcampos_parcelas.poligono,rcampos_parcelas.parcela,rcampos_parcelas.subparce, "
            Sql = Sql & "rcampos_parcelas.recintos,rcampos_parcelas.codsigpa,rcampos_parcelas.supsigpa,"
            Sql = Sql & "rcampos_parcelas.supcultsigpa,rcampos_parcelas.supcatas,rcampos_parcelas.supcultcatas"
            Sql = Sql & " FROM " & tabla
            If enlaza Then
                Sql = Sql & ObtenerWhereCab2(True)
            Else
                Sql = Sql & " WHERE rcampos_parcelas.codcampo = -1"
            End If
            Sql = Sql & " ORDER BY " & tabla & ".numlinea "
            
       Case 3 ' agroseguro
            tabla = "rcampos_seguros"
            Sql = "SELECT rcampos_seguros.codcampo, rcampos_seguros.numlinea, rcampos_seguros.fecha, rcampos_seguros.codincid, rincidencia.nomincid, "
            Sql = Sql & "rcampos_seguros.kilos,rcampos_seguros.kilosaportacion, rcampos_seguros.importe,rcampos_seguros.fechapago,"
            Sql = Sql & "rcampos_seguros.essiniestro , IF(essiniestro=1,'*','') as dsiniestro "
            Sql = Sql & " FROM " & tabla & " INNER JOIN rincidencia ON rcampos_seguros.codincid = rincidencia.codincid "
            If enlaza Then
                Sql = Sql & ObtenerWhereCab3(True)
            Else
                Sql = Sql & " WHERE rcampos_seguros.codcampo = -1"
            End If
            Sql = Sql & " ORDER BY " & tabla & ".numlinea "
            
        Case 4 ' hco del campo
            tabla = "rcampos_hco"
            Sql = "SELECT rcampos_hco.codcampo, rcampos_hco.numlinea, rcampos_hco.codsocio, rsocios.nomsocio, rcampos_hco.fechaalta, "
            Sql = Sql & "rcampos_hco.fechabaja, rcampos_hco.codincid, rincidencia.nomincid"
            Sql = Sql & " FROM (" & tabla & " INNER JOIN rincidencia ON rcampos_hco.codincid = rincidencia.codincid) "
            Sql = Sql & " INNER JOIN rsocios ON rcampos_hco.codsocio = rsocios.codsocio "
            If enlaza Then
                Sql = Sql & ObtenerWhereCab4(True)
            Else
                Sql = Sql & " WHERE rcampos_hco.codcampo = -1"
            End If
            Sql = Sql & " ORDER BY " & tabla & ".numlinea "
        
        Case 5 ' gastos del campo
            tabla = "rcampos_gastos"
            Sql = "SELECT rcampos_gastos.codcampo, rcampos_gastos.numlinea, rcampos_gastos.codgasto, rconcepgasto.nomgasto, rcampos_gastos.fecha, "
            Sql = Sql & "rcampos_gastos.importe, rcampos_gastos.contabilizado, IF(contabilizado=1,'*','') as dcontabilizado "
            Sql = Sql & " FROM " & tabla & " INNER JOIN rconcepgasto ON rcampos_gastos.codgasto = rconcepgasto.codgasto "
            
            If enlaza Then
                Sql = Sql & ObtenerWhereCab5(True)
            Else
                Sql = Sql & " WHERE rcampos_gastos.codcampo = -1"
            End If
            Sql = Sql & " ORDER BY " & tabla & ".numlinea "
    
        Case 6 ' impresion de ordenes de recoleccion del campo
            tabla = "rcampos_ordrec"
            Sql = "SELECT rcampos_ordrec.codcampo, rcampos_ordrec.nroorden, rcampos_ordrec.fecimpre "
            Sql = Sql & " FROM " & tabla
            
            If enlaza Then
                Sql = Sql & ObtenerWhereCab6(True)
            Else
                Sql = Sql & " WHERE rcampos_ordrec.codcampo = -1"
            End If
            Sql = Sql & " ORDER BY " & tabla & ".nroorden "
    
        Case 7 ' revisiones
            tabla = "rcampos_revision"
            Sql = "SELECT rcampos_revision.codcampo, rcampos_revision.numlinea, rcampos_revision.fecha, rcampos_revision.tecnico, rcampos_revision.observac "
            Sql = Sql & " FROM " & tabla
            
            If enlaza Then
                Sql = Sql & ObtenerWhereCab7(True)
            Else
                Sql = Sql & " WHERE rcampos_revision.codcampo = -1"
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

'Private Sub frmC_Selec(vFecha As Date)
'Dim Indice As Byte
'    Indice = CByte(Me.cmdAux(0).Tag + 2)
'    txtAux1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
'End Sub

Private Sub frmC1_Selec(vFecha As Date)
Dim Indice As Byte
    Indice = CByte(Me.imgFec(0).Tag)
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

'Private Sub frmC2_Selec(vFecha As Date)
'Dim Indice As Byte
'    Indice = CByte(Me.cmdAux(0).Tag)
'    txtAux5(Indice).Text = Format(vFecha, "dd/mm/yyyy")
'End Sub
'
'Private Sub frmC3_Selec(vFecha As Date)
'Dim Indice As Byte
'    Indice = CByte(Me.cmdAux(0).Tag)
'    txtaux7(Indice).Text = Format(vFecha, "dd/mm/yyyy")
'End Sub
'
'Private Sub frmC4_Selec(vFecha As Date)
'Dim Indice As Byte
'    Indice = CByte(Me.cmdAux(0).Tag)
'    txtaux8(Indice).Text = Format(vFecha, "dd/mm/yyyy")
'End Sub
'
'Private Sub frmC5_Selec(vFecha As Date)
'Dim Indice As Byte
'    Indice = CByte(Me.cmdAux(0).Tag)
'    txtAux9(Indice).Text = Format(vFecha, "dd/mm/yyyy")
'End Sub

'Private Sub frmCalid_DatoSeleccionado(CadenaSeleccion As String)
'    txtAux1(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo variedad
'    txtAux1(1).Text = RecuperaValor(CadenaSeleccion, 3) 'codigo calidad
'    txtAux2(1).Text = RecuperaValor(CadenaSeleccion, 4) 'nombre calidad
'End Sub

Private Sub frmCamPrev_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
    
    If CadenaSeleccion <> "" Then
        CadB = "codcampo = " & DBSet(RecuperaValor(CadenaSeleccion, 1), "N")
        
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub frmDesa_DatoSeleccionado(CadenaSeleccion As String)
    Text1(26).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo desarrollo vegetativo
    FormateaCampo Text1(26)
    Text2(26).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre desarrollo vegetativo
End Sub

Private Sub frmGlo_DatoSeleccionado(CadenaSeleccion As String)
'globalgap
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo de globalgap
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

'Private Sub frmGto_DatoSeleccionado(CadenaSeleccion As String)
'    txtaux7(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo concepto de gasto
'    FormateaCampo txtaux7(2)
'    txtAux2(5).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre concepto de gasto
'End Sub

'Private Sub frmInc_DatoSeleccionado(CadenaSeleccion As String)
'    txtAux5(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo incidencia
'    FormateaCampo txtAux5(3)
'    txtAux2(2).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre incidencia
'End Sub

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

    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo de partida
    FormateaCampo Text1(3)
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre de partida
    
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
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo situacion
    FormateaCampo Text1(3)
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre situacion
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    Text1(1).Text = RecuperaValor(CadenaSeleccion, 1) 'codsocio
    FormateaCampo Text1(Indice)
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2) 'nomsocio
End Sub

'Private Sub frmSoc1_DatoSeleccionado(CadenaSeleccion As String)
'    txtAux3(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codsocio
'    FormateaCampo txtAux3(2)
'    txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'nomsocio
'End Sub


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
    If Indice = 6 Then
        Text1(Indice).Text = vCampo
    Else
'        txtAux9(Indice).Text = vCampo
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
        
       menu = Me.Height - Me.ScaleHeight 'ac� tinc el heigth del men� i de la toolbar
    
       frmC1.Left = esq + imgFec(Index).Parent.Left + 30
       frmC1.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40
    
       
       frmC1.NovaData = Now
       Select Case Index
            Case 0, 1
                Indice = Index + 4
            Case 2, 3
                Indice = Index + 40
       End Select
       
       Me.imgFec(0).Tag = Indice
       
       PonerFormatoFecha Text1(Indice)
       If Text1(Indice).Text <> "" Then frmC1.NovaData = CDate(Text1(Indice).Text)
    
       Screen.MousePointer = vbDefault
       frmC1.Show vbModal
       Set frmC1 = Nothing
       PonerFoco Text1(Indice)
    
End Sub

Private Sub imgZoom_Click(Index As Integer)
    Set frmZ = New frmZoom

    Select Case Index
        Case 0
            Indice = 6
            frmZ.pTitulo = "Observaciones del Campo"
            frmZ.pValor = Text1(Indice).Text
            frmZ.pModo = Modo
            frmZ.Show vbModal
            Set frmZ = Nothing
            PonerFoco Text1(Indice)
            
        Case 1
            Indice = 4
            frmZ.pTitulo = "Observaciones de la Revisi�n"
            If Modo = 5 Then
                frmZ.pModo = 3
'                frmZ.pValor = txtAux9(Indice).Text
            Else
                frmZ.pModo = Modo
'                frmZ.pValor = DBLet(Me.Adoaux(7).Recordset!Observac, "T")
            End If
            frmZ.Show vbModal
            Set frmZ = Nothing
            PonerFoco Text1(Indice)
            
            
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
        MsgBox "No tiene configurada en par�metros la direcci�n de Goolzoom. Llame a Soporte.", vbExclamation
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass

    '[Monica]23/12/2016: cambiado el lanzahomegnral por LanzaVisorMimeDocumento
    LanzaVisorMimeDocumento Me.hWnd, Direccion
    'If LanzaHomeGnral(Direccion) Then espera 2
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
                MsgBox "No existe el c�digo de poblacion de la partida", vbExclamation
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
        MsgBox "No tiene configurada en par�metros la direcci�n de Sigpac. Llame a Soporte.", vbExclamation
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
    
    'A�adir el parametro de Empresa
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
'Alvan�ar/Retrocedir els camps en les fleches de despla�ament del teclat.
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


Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 1 ' Verificacion de Errores
            mnVerificacionErr_Click
        Case 2
            mnSigpac_Click
        Case 3
            mnGoolzoom_Click
        Case 4
            mnChequeoNroOrden_Click
        Case 5
            mnCambioSocio_Click
        Case 6
            mnGastosCampos_Click
        Case 7
            mnGlobalGap_Click
    End Select
End Sub



Private Sub BotonBuscar()
Dim I As Integer
' ***** Si la clau primaria de la cap�alera no es Text1(0), canviar-ho en <=== *****
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonerFoco Text1(0) ' <===
        Text1(0).BackColor = vbLightBlue ' <===
        ' *** si n'hi han combos a la cap�alera ***
        
        EstablecerOrden True
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

    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
    If chkVistaPrevia(0) = 1 Then
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
'    Dim Cad As String
'    Dim NombreTabla1 As String
'
'    'Cridem al form
'    ' **************** arreglar-ho per a vore lo que es desije ****************
'    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
'    Cad = ""
'    Cad = Cad & "C�digo|rcampos.codcampo|N|000000|9�"
'    Cad = Cad & "Socio|rcampos.codsocio|N|000000|9�"
'    Cad = Cad & "Nombre|rsocios.nomsocio|T||30�"
'    Cad = Cad & "Variedad|variedades.nomvarie|T||16�"
'    Cad = Cad & "Partida|rpartida.nomparti|T||15�"
'    Cad = Cad & "Pol.|rcampos.poligono|T||5�"
'    Cad = Cad & "Parc.|rcampos.parcela|T||7�"
'    Cad = Cad & "Sp.|rcampos.subparce|T||4�"
'    Cad = Cad & "Nro|rcampos.nrocampo|T||5�"
'
'    NombreTabla1 = "((rcampos inner join rsocios on rcampos.codsocio = rsocios.codsocio) inner join variedades on rcampos.codvarie = variedades.codvarie) " & _
'                   " inner join rpartida on rcampos.codparti = rpartida.codparti "
'
'
'    If Cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = Cad
'        frmB.vtabla = NombreTabla1
'        frmB.vSQL = CadB
'        HaDevueltoDatos = False
'        frmB.vDevuelve = "0|" '*** els camps que volen que torne ***
'        frmB.vTitulo = "Campos" ' ***** repasa a��: t�tol de BuscaGrid *****
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

    Set frmCamPrev = New frmManCamposMonastPrev
    frmCamPrev.cWhere = CadB
    frmCamPrev.DatosADevolverBusqueda = "0|1|2|"
    frmCamPrev.Show vbModal
    
    Set frmCamPrev = Nothing



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
    Text1(0).Text = SugerirCodigoSiguienteStr("rcampos", "codcampo")
    FormateaCampo Text1(0)
       
    Text1(4).Text = Format(Now, "dd/mm/yyyy")
    Text1(3).Text = 1
    
    '[Monica]29/09/2014: comprobamos si vamos a dar de baja que no tenga fecha de alta en programa operativo
    FecBajaAnt = Text1(5).Text
       
    PonerFoco Text1(0) '*** 1r camp visible que siga PK ***
    ' ***********************************************************
    
    ' *** si n'hi han camps de descripci� a la cap�alera ***
    PosarDescripcions
    ' ******************************************************

    ' *** si n'hi han tabs, em posicione al 1r ***
    Me.SSTab1.Tab = 0
    ' ********************************************
    
    EstablecerOrden True
End Sub


Private Sub BotonModificar()

    PonerModo 4

    ' *** bloquejar els camps visibles de la clau primaria de la cap�alera ***
    BloquearTxt Text1(0), True
    ' *************************************************************************
    
    SocioAnt = Text1(1).Text
    '[Monica]29/09/2014: comprobamos si vamos a dar de baja que no tenga fecha de alta en programa operativo
    FecBajaAnt = Text1(5).Text
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    ' *********************************************************
    
    EstablecerOrden True
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
    cad = "�Seguro que desea eliminar el Campo?"
    cad = cad & vbCrLf & "C�digo: " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
    cad = cad & vbCrLf & "Socio : " & Data1.Recordset.Fields(1)
    cad = cad & vbCrLf & "Nombre: " & Text2(1).Text
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
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la cap�alera
    
    ' *** si n'hi han ll�nies en datagrids ***
'    For I = 0 To DataGridAux.Count - 1 '[Monica]30/09/2013: antes - 1
'        If I <> 6 Then
'            CargaGrid I, True
'            If Not Adoaux(I).Recordset.EOF Then _
'                PonerCamposForma2 Me, Adoaux(I), 2, "FrameAux" & I
'        End If
'    Next I
'    '[Monica]30/09/2013
    
    ' *******************************************

    ' *** si n'hi han ll�nies sense datagrid ***
    ' ***************************************
    
    ' ************* configurar els camps de les descripcions de la cap�alera *************
'    txtAux2(22).Text = PonerNombreDeCod(txtAux(22), "poblacio", "despobla", "codpobla", "N")

    PosarDescripcions

    '[Monica]23/10/2013: Solo si es Escalona o Utxera (o de momento montifrut) damos mensaje de que el socio tiene pagos pendientes
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Or vParamAplic.Cooperativa = 12 Then
        '[Monica]15/05/2013: Visualizamos los cobros pendientes del socio
        ComprobarCobrosSocio CStr(Data1.Recordset!Codsocio), ""
    End If


    
    '[Monica]09/05/2017: solo en el caso de que est� en documento de baja se visualiza
    B = EstaEnDocumentoBaja(Text1(0).Text)
    
    
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu   'Activar opcions de men� seg�n nivell
                        'de permisos de l'usuari
    
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
                
                
        Case 5 'LL�NIES
            Select Case ModoLineas
                Case 1 'afegir ll�nia
                    ModoLineas = 0
                    ' *** les ll�nies que tenen datagrid (en o sense tab) ***
                    If NumTabMto = 0 Or NumTabMto = 1 Or NumTabMto = 2 Or NumTabMto = 3 Or NumTabMto = 4 Or NumTabMto = 5 Or NumTabMto = 6 Or NumTabMto = 7 Then
'                        DataGridAux(NumTabMto).AllowAddNew = False
                        ' **** repasar si es diu Data1 l'adodc de la cap�alera ***
                        'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
                        ' ********************************************************
                        LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
'                        DataGridAux(NumTabMto).Enabled = True
'                        DataGridAux(NumTabMto).SetFocus

                        ' *** si n'hi han camps de descripci� dins del grid, els neteje ***
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

'                    If Not Adoaux(NumTabMto).Recordset.EOF Then
'                        Adoaux(NumTabMto).Recordset.MoveFirst
'                    End If

                Case 2 'modificar ll�nies
                    ModoLineas = 0

                    ' *** si n'hi han tabs ***
                    SituarTab (NumTabMto)
                    'SSTab1.Tab = 1
                    'SSTab2.Tab = NumTabMto
                    ' ***********************

                    PonerModo 4
'                    If Not Adoaux(NumTabMto).Recordset.EOF Then
'                        ' *** l'Index de Fields es el que canvie de la PK de ll�nies ***
'                        V = Adoaux(NumTabMto).Recordset.Fields(1) 'el 2 es el n� de llinia
'                        Adoaux(NumTabMto).Recordset.Find (Adoaux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
'                        ' ***************************************************************
'                    End If
                    ' ***  bloquejar els camps fora dels grids ***
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
            End Select
            

            PosicionarData

            ' *** si n'hi han ll�nies en grids i camps fora d'estos ***
'            If Not Adoaux(NumTabMto).Recordset.EOF Then
'                DataGridAux_RowColChange NumTabMto, 1, 1
'            Else
'                LimpiarCamposFrame NumTabMto
'            End If
            ' *********************************************************
    End Select
End Sub


Private Function DatosOK() As Boolean
Dim B As Boolean
Dim Sql As String
Dim cad As String
Dim Rs As ADODB.Recordset
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOK = False
    B = CompForm2(Me, 1)
    If Not B Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
        'comprobar si existe ya el cod. del campo clave primaria
        If ExisteCP(Text1(0)) Then B = False
    End If
    
    'miramos si hay otros campos con la misma ubicacion
    If B And (Modo = 3 Or Modo = 4) Then
        ' comprobamos que el socio no est� dado de baja
        If B Then
            Sql = "select fechabaja from rsocios where codsocio = " & DBSet(Text1(1).Text, "N")
            
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If DBLet(Rs.Fields(0).Value, "F") <> "" Then
                cad = "Este socio tiene fecha de baja. � Desea continuar ?"
                If MsgBox(cad, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                    B = False
                End If
            End If
        End If
        
        
        '[Monica]31/10/2014: si la fecha de alta es superior a la fecha de alta del socio de la seccion de horto damos un aviso
        If B Then
            Sql = "select fecalta from rsocios_seccion where codsocio = " & DBSet(Text1(1).Text, "N") & " and codsecci = " & DBSet(vParamAplic.Seccionhorto, "N")
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not Rs.EOF Then
                If DBLet(Rs.Fields(0).Value, "F") > CDate(Text1(4).Text) Then
                    Sql = "La fecha de alta del socio en la Seccion de Horto es superior a la fecha de alta del campo." & vbCrLf & vbCrLf
                    Sql = Sql & "                     � Desea continuar ?"
                    
                    If MsgBox(Sql, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                        B = False
                    End If
                End If
            Else
                Sql = "El socio no se encuentra en la Secci�n de Horto." & vbCrLf & vbCrLf
                Sql = Sql & "                     � Desea continuar ?"
                If MsgBox(Sql, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                    B = False
                End If
            End If
            Set Rs = Nothing
        End If
        
    End If
    
    
    
    
    ' ************************************************************************************
    DatosOK = B
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub PosicionarData()
Dim cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la cap�alera, no llevar els () ***
    cad = "(codcampo=" & Text1(0).Text & ")"
    ' ***************************************
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarData(Data1, cad, Indicador) Then
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
    vWhere = " WHERE codcampo=" & Data1.Recordset!codcampo
        ' ***********************************************************************
        
    ' ***** elimina les ll�nies ****
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


'    'Eliminar la CAP�ALERA
'    vWhere = " WHERE codsocio=" & Data1.Recordset!codsocio
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
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        
                
        Case 2 'PARTIDA
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rpartida", "nomparti", "codparti", "N")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Partida: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearla?" & vbCrLf
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
                
                
        Case 3 'SITUACION Campo
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rsituacioncampo", "nomsitua")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Situaci�n Campo: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearla?" & vbCrLf
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
                
        '[Monica]29/09/2014: campo 43, fecha de alta en programa operativo
        Case 4, 5 'Fecha no comprobaremos que est� dentro de campa�a
                    'Fecha de alta y fecha baja
            If Modo = 1 Then Exit Sub
            PonerFormatoFecha Text1(Index), False
            
    End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 7 Then
        If KeyAscii = teclaBuscar Then
            If Modo = 1 Or Modo = 3 Or Modo = 4 Then
                Select Case Index
                    Case 1: KEYBusqueda KeyAscii, 1 'socio
                    Case 2: KEYBusqueda KeyAscii, 3 'partida
                    Case 4: KEYFecha KeyAscii, 0 ' fecha alta
                    Case 5: KEYFecha KeyAscii, 1 ' fecha baja
                    Case 3: KEYBusqueda KeyAscii, 0 'situacion
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

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFec_Click (Indice)
End Sub

Private Sub KEYBusquedaAux(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
'    cmdAux_Click (Indice)
End Sub



' **** si n'hi han camps de descripci� a la cap�alera ****
Private Sub PosarDescripcions()
Dim NomEmple As String

    On Error GoTo EPosarDescripcions

    Text2(1).Text = PonerNombreDeCod(Text1(1), "rsocios", "nomsocio", "codsocio", "N")
    Text2(3).Text = PonerNombreDeCod(Text1(3), "rsituacioncampo", "nomsitua", "codsitua", "N")
    Text2(2).Text = PonerNombreDeCod(Text1(2), "rpartida", "nomparti", "codparti", "N")
    
    '[Monica]14/02/2013: sacamos el codigo de conselleria de las lineas
'    txtAux2(13).Text = DevuelveValor("select codconse from variedades where codvarie = " & DBSet(Text1(2).Text, "N"))
    
    PonerDatosPartida
    
    
    
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
'    frmListado.NumCod = "rcampos_gastos.codcampo = " & Adoaux(5).Recordset!codcampo & " and rcampos_gastos.numlinea = " & Adoaux(5).Recordset!numlinea
    frmListado.Show vbModal
    CargaGrid NumTabMto, True
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub BotonEliminarLinea(Index As Integer)
'Dim Sql As String
'Dim vWhere As String
'Dim Eliminar As Boolean
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
'    If Adoaux(Index).Recordset.EOF Then Exit Sub
'    If Not SepuedeBorrar(Index) Then Exit Sub
'    NumTabMto = Index
'    Eliminar = False
'
'    vWhere = ObtenerWhereCab(True)
'
'    ' ***** independentment de si tenen datagrid o no,
'    ' canviar els noms, els formats i el DELETE *****
'    Select Case Index
'        Case 0 'coopropietarios
'            Sql = "�Seguro que desea eliminar el coopropietario?"
'            Sql = Sql & vbCrLf & "Coopropietario: " & Adoaux(Index).Recordset!Codsocio & " - " & Adoaux(Index).Recordset!nomsocio
'            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
'                Eliminar = True
'                Sql = "DELETE FROM rcampos_cooprop"
'                Sql = Sql & " WHERE rcampos_cooprop.codcampo = " & DBLet(Adoaux(Index).Recordset!codcampo, "N")
'                Sql = Sql & " and codsocio = " & DBLet(Adoaux(Index).Recordset!Codsocio, "N")
'            End If
'
'        Case 1 'clasificacion
'            Sql = "�Seguro que desea eliminar la clasificaci�n?"
'            Sql = Sql & vbCrLf & "Clasificaci�n: " & Adoaux(Index).Recordset!codcalid & " - " & Adoaux(Index).Recordset!nomcalid
'            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
'                Eliminar = True
'                Sql = "DELETE FROM rcampos_clasif"
'                Sql = Sql & vWhere & " AND codvarie= " & DBLet(Adoaux(Index).Recordset!codvarie, "N")
'                Sql = Sql & " and codcalid = " & DBLet(Adoaux(Index).Recordset!codcalid, "N")
'            End If
'
'        Case 2 'parcelas
'            vWhere = ObtenerWhereCab2(True)
'
'            Sql = "�Seguro que desea eliminar la parcela?"
'            Sql = Sql & vbCrLf & "P�ligono: " & Adoaux(Index).Recordset!Poligono & " - Parcela : " & Adoaux(Index).Recordset!Parcela
'            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
'                Eliminar = True
'                Sql = "DELETE FROM rcampos_parcelas"
'                Sql = Sql & vWhere & " AND numlinea= " & DBLet(Adoaux(Index).Recordset!numlinea, "N")
'            End If
'
'        Case 3 'agroseguro
'            vWhere = ObtenerWhereCab3(True)
'
'            Sql = "�Seguro que desea eliminar la L�nea?"
'            Sql = Sql & vbCrLf & "Fecha: " & Adoaux(Index).Recordset!Fecha & " - Incidencia : " & Adoaux(Index).Recordset!nomincid
'            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
'                Eliminar = True
'                Sql = "DELETE FROM rcampos_seguros"
'                Sql = Sql & vWhere & " AND numlinea= " & DBLet(Adoaux(Index).Recordset!numlinea, "N")
'            End If
'
'        Case 4 'hco de campos
'            vWhere = ObtenerWhereCab4(True)
'
'            Sql = "�Seguro que desea eliminar la L�nea?" & vbCrLf
'            Sql = Sql & "Socio: " & Format(Adoaux(Index).Recordset!Codsocio, "000000") & " - " & Adoaux(Index).Recordset!nomsocio
'            Sql = Sql & vbCrLf & "Fecha Alta: " & Adoaux(Index).Recordset!FechaAlta
'            Sql = Sql & vbCrLf & "Fecha Baja: " & Adoaux(Index).Recordset!FechaBaja
'            Sql = Sql & vbCrLf & "Incidencia : " & Adoaux(Index).Recordset!nomincid
'            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
'                Eliminar = True
'                Sql = "DELETE FROM rcampos_hco"
'                Sql = Sql & vWhere & " AND numlinea= " & DBLet(Adoaux(Index).Recordset!numlinea, "N")
'            End If
'
'        Case 5 'gastos de campos
'            vWhere = ObtenerWhereCab5(True)
'
'            If Adoaux(Index).Recordset!contabilizado Then
'                Sql = "Este Gasto est� contabilizado. Si continua deber� modificar la contabilidad." & vbCrLf
'                Sql = Sql & " � Desea continuar ? "
'                If MsgBox(Sql, vbQuestion + vbYesNo) = vbNo Then Exit Sub
'            End If
'
'
'            Sql = "�Seguro que desea eliminar la L�nea?" & vbCrLf
'            Sql = Sql & "Concepto: " & Format(Adoaux(Index).Recordset!Codgasto, "00") & " - " & Adoaux(Index).Recordset!NomGasto
'            Sql = Sql & vbCrLf & "Fecha: " & Adoaux(Index).Recordset!Fecha
'            Sql = Sql & vbCrLf & "Importe: " & Adoaux(Index).Recordset!Importe
'            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
'                Eliminar = True
'                Sql = "DELETE FROM rcampos_gastos"
'                Sql = Sql & vWhere & " AND numlinea= " & DBLet(Adoaux(Index).Recordset!numlinea, "N")
'            End If
'
'        Case 6 'ordenes de recoleccion
'            vWhere = ObtenerWhereCab6(True)
'
'
'            Sql = "�Seguro que desea eliminar la L�nea?" & vbCrLf
'            Sql = Sql & "Orden: " & Format(Adoaux(Index).Recordset!nroorden, "0000000")
'            Sql = Sql & vbCrLf & "Fecha: " & Adoaux(Index).Recordset!fecimpre
'            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
'                Eliminar = True
'                Sql = "DELETE FROM rcampos_ordrec"
'                Sql = Sql & vWhere & " AND nroorden= " & DBLet(Adoaux(Index).Recordset!nroorden, "N")
'            End If
'
'        Case 7 'revisiones de campos
'            vWhere = ObtenerWhereCab7(True)
'
'
'            Sql = "�Seguro que desea eliminar la L�nea?" & vbCrLf
'            Sql = Sql & "Fecha: " & Adoaux(Index).Recordset!Fecha
'            Sql = Sql & vbCrLf & "T�cnico: " & Adoaux(Index).Recordset!tecnico
'            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
'                Eliminar = True
'                Sql = "DELETE FROM rcampos_revision"
'                Sql = Sql & vWhere & " AND numlinea= " & DBLet(Adoaux(Index).Recordset!numlinea, "N")
'            End If
'
'
'    End Select
'
'    If Eliminar Then
'        NumRegElim = Adoaux(Index).Recordset.AbsolutePosition
'        TerminaBloquear
'        conn.Execute Sql
'        ' *** si n'hi han tabs sense datagrid, posar l'If ***
''        If Index <> 3 Then
'            CargaGrid Index, True
'        ' ***************************************************
'        If Not SituarDataTrasEliminar(Adoaux(Index), NumRegElim, True) Then
'
'        End If
'        SumaTotalPorcentajes NumTabMto
'        ' *** si n'hi han tabs sense datagrid ***
''        If Index = 3 Then CargaFrame 3, True
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
'Dim vWhere As String, vtabla As String
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
'        Case 0: vtabla = "rcampos_cooprop"
'        Case 1: vtabla = "rcampos_clasif"
'        Case 2: vtabla = "rcampos_parcelas"
'        Case 3: vtabla = "rcampos_seguros"
'        Case 4: vtabla = "rcampos_hco"
'        Case 5: vtabla = "rcampos_gastos"
'        Case 6: vtabla = "rcampos_ordrec"
'        Case 7: vtabla = "rcampos_revision"
'    End Select
'    ' ********************************************************
'
'    vWhere = ObtenerWhereCab(False)
'
'    Select Case Index
'         Case 0, 1, 2, 3, 4, 5, 6, 7 'clasificacion
'            ' *** canviar la clau primaria de les ll�nies,
'            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
'            Select Case Index
'                Case 0
'                    NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", "rcampos_cooprop.codcampo = " & Val(Text1(0).Text))
'                Case 1
'                    NumF = ""
'                Case 2
'                    NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", "rcampos_parcelas.codcampo = " & Val(Text1(0).Text))
'                Case 3
'                    NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", "rcampos_seguros.codcampo = " & Val(Text1(0).Text))
'                Case 4
'                    NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", "rcampos_hco.codcampo = " & Val(Text1(0).Text))
'                Case 5
'                    NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", "rcampos_gastos.codcampo = " & Val(Text1(0).Text))
'                Case 7
'                    NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", "rcampos_revision.codcampo = " & Val(Text1(0).Text))
'            End Select
'            ' ***************************************************************
'
'            AnyadirLinea DataGridAux(Index), Adoaux(Index)
'
'            anc = DataGridAux(Index).Top
'            If DataGridAux(Index).Row < 0 Then
'                anc = anc + 240
'            Else
'                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
'            End If
'
'            LLamaLineas Index, ModoLineas, anc
'
'            Select Case Index
'                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
'                Case 1 'secciones
'                    For I = 0 To txtAux1.Count - 1
'                        txtAux1(I).Text = ""
'                    Next I
'                    txtAux1(0).Text = Text1(0).Text 'codcampo
'                    txtAux1(2).Text = Text1(2).Text 'codvariedad
'                    txtAux2(1).Text = ""
'                    PonerFoco txtAux1(1)
'
'                Case 0 'copropietarios
'                    For I = 0 To txtAux3.Count - 1
'                        txtAux3(I).Text = ""
'                    Next I
'                    txtAux2(0).Text = ""
'                    txtAux3(0).Text = Text1(0).Text 'codcampo
'                    txtAux3(1).Text = NumF 'numlinea
'                    txtAux3(2).Text = ""
'                    PonerFoco txtAux3(2)
'
'                Case 2 ' parcelas
'                    For I = 0 To txtAux4.Count - 1
'                        txtAux4(I).Text = ""
'                    Next I
'                    txtAux4(0).Text = Text1(0).Text 'codcampo
'                    txtAux4(1).Text = NumF 'numlinea
'                    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then txtAux4(6).Text = "0"
'                    PonerFoco txtAux4(2)
'
'                Case 3 ' seguros
'                    For I = 0 To txtAux5.Count - 1
'                        txtAux5(I).Text = ""
'                    Next I
'                    txtAux2(2).Text = ""
'
'                    txtAux5(0).Text = Text1(0).Text 'codcampo
'                    txtAux5(1).Text = NumF 'numlinea
'                    PonerFoco txtAux5(2)
'
'                    Me.chkAux(0).Value = 0
'
'                Case 4 ' hco de campos
'                    For I = 0 To txtaux6.Count - 1
'                        txtaux6(I).Text = ""
'                    Next I
'                    txtaux6(0).Text = Text1(0).Text 'codcampo
'                    txtaux6(1).Text = NumF 'numlinea
'                    PonerFoco txtaux6(2)
'
'                Case 5 ' gastos de  campos
'                    For I = 0 To txtaux7.Count - 1
'                        txtaux7(I).Text = ""
'                    Next I
'                    txtAux2(5).Text = ""
'                    txtaux7(0).Text = Text1(0).Text 'codcampo
'                    txtaux7(1).Text = NumF 'numlinea
'                    PonerFoco txtaux7(2)
'
'                    Me.chkAux(1).Value = 0
'
'                Case 6 ' ordenes de recoleccion de campos
'                    For I = 0 To txtaux8.Count - 1
'                        txtaux8(I).Text = ""
'                    Next I
'                    txtaux8(0).Text = Text1(0).Text 'codcampo
'                    PonerFoco txtaux8(1)
'
'                Case 7 ' revisiones de campo
'                    For I = 0 To txtAux9.Count - 1
'                        txtAux9(I).Text = ""
'                    Next I
'                    txtAux9(0).Text = Text1(0).Text 'codcampo
'                    txtAux9(1).Text = NumF 'numlinea
'                    txtAux9(2).Text = Now ' fecha
'                    PonerFoco txtAux9(2)
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
'    Dim Sql As String
'
'
'    If Adoaux(Index).Recordset.EOF Then Exit Sub
'    If Adoaux(Index).Recordset.RecordCount < 1 Then Exit Sub
'
'    If Index = 5 Then
'        If CInt(Adoaux(Index).Recordset!contabilizado) = 1 Then
'            Sql = "Este Gasto est� contabilizado, deber� modificar la contabilidad." & vbCrLf
'            Sql = Sql & " � Desea continuar ? "
'            If MsgBox(Sql, vbQuestion + vbYesNo) = vbNo Then Exit Sub
'        End If
'    End If
'
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
'
'
'    Select Case Index
'        Case 0, 1, 2, 3, 4, 5, 6, 7 ' *** pose els index de ll�nies que tenen datagrid (en o sense tab) ***
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
'        Case 0 'coopropietarios
'            txtAux3(0).Text = DataGridAux(Index).Columns(0).Text
'            txtAux3(1).Text = DataGridAux(Index).Columns(1).Text
'            txtAux3(2).Text = DataGridAux(Index).Columns(2).Text
'
'            txtAux2(0).Text = DataGridAux(Index).Columns(3).Text
'            txtAux3(3).Text = DataGridAux(Index).Columns(4).Text
'
'        Case 1 'clasificacion
'            txtAux1(0).Text = DataGridAux(Index).Columns(0).Text
'            txtAux1(1).Text = DataGridAux(Index).Columns(2).Text
'            txtAux1(2).Text = DataGridAux(Index).Columns(1).Text
'
'            txtAux2(1).Text = DataGridAux(Index).Columns(3).Text
'            txtAux1(3).Text = DataGridAux(Index).Columns(4).Text
'
'        Case 2 'parcelas
'            For I = 0 To 10
'                txtAux4(I).Text = DataGridAux(Index).Columns(I).Text
'            Next I
'
'        Case 3 'seguros
'            For I = 0 To 3
'                txtAux5(I).Text = DataGridAux(Index).Columns(I).Text
'            Next I
'            txtAux2(2).Text = DataGridAux(Index).Columns(4).Text
'            '[Monica]26/01/2016: a�adida nueva columna de kilos de aportacion
''            For I = 4 To 6
''                txtAux5(I).Text = DataGridAux(Index).Columns(I + 1).Text
''            Next I
'            txtAux5(4).Text = DataGridAux(Index).Columns(5).Text
'            txtAux5(7).Text = DataGridAux(Index).Columns(6).Text
'            txtAux5(5).Text = DataGridAux(Index).Columns(7).Text
'            txtAux5(6).Text = DataGridAux(Index).Columns(8).Text
'
'            Me.chkAux(0).Value = Me.Adoaux(3).Recordset!essiniestro
'
'
'        Case 4 'hco de campos
'            For I = 0 To 2
'                txtaux6(I).Text = DataGridAux(Index).Columns(I).Text
'            Next I
'            txtAux2(4).Text = DataGridAux(Index).Columns(3).Text
'            For I = 3 To 5
'                txtaux6(I).Text = DataGridAux(Index).Columns(I + 1).Text
'            Next I
'            txtAux2(3).Text = DataGridAux(Index).Columns(7).Text
'
'        Case 5 'gastos de campos
'            For I = 0 To 2
'                txtaux7(I).Text = DataGridAux(Index).Columns(I).Text
'            Next I
'            txtAux2(5).Text = DataGridAux(Index).Columns(3).Text
'            For I = 3 To 4
'                txtaux7(I).Text = DataGridAux(Index).Columns(I + 1).Text
'            Next I
'            chkAux(1).Value = DataGridAux(Index).Columns(6).Text
'
'        Case 6 'ordenes de recoleccion de campos
'            For I = 0 To 2
'                txtaux8(I).Text = DataGridAux(Index).Columns(I).Text
'            Next I
'
'        Case 7 'revisiones de campo
'            For I = 0 To txtAux9.Count - 1
'                txtAux9(I).Text = DataGridAux(Index).Columns(I).Text
'            Next I
'
'    End Select
'
'    LLamaLineas Index, ModoLineas, anc
'
'    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
'    Select Case Index
'        Case 0 'coopropietarios
'            PonerFoco txtAux3(2)
'        Case 1 'clasificacion
'            PonerFoco txtAux1(3)
'        Case 2 'parcelas
'            PonerFoco txtAux4(2)
'        Case 3 'agroseguro
'            PonerFoco txtAux5(2)
'        Case 4 'hco
'            PonerFoco txtaux6(2)
'        Case 5 'gastos de campos
'            PonerFoco txtaux7(2)
'        Case 6 'ordenes de recoleccion de campos
'            PonerFoco txtaux8(1)
'        Case 7 'revisiones de campo
'            PonerFoco txtAux9(2)
'    End Select
'    ' ***************************************************************************************
End Sub


Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
'Dim jj As Integer
'Dim B As Boolean
'
'    ' *** si n'hi han tabs sense datagrid posar el If ***
'    If Index <> 3 Then DeseleccionaGrid DataGridAux(Index)
'    ' ***************************************************
'
'    B = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Ll�nies
'    Select Case Index
'        Case 1 'clasificacion
'            For jj = 1 To txtAux1.Count - 1
'                If jj <> 2 Then
'                    txtAux1(jj).visible = B
'                    txtAux1(jj).Top = alto
'                End If
'            Next jj
'
'            txtAux2(1).visible = B
'            txtAux2(1).Top = alto
'
'            For jj = 1 To 1
'                cmdAux(jj).visible = B
'                cmdAux(jj).Top = txtAux1(3).Top
'                cmdAux(jj).Height = txtAux1(3).Height
'            Next jj
'        Case 0 ' coopropietarios
'            For jj = 2 To txtAux3.Count - 1
'                txtAux3(jj).visible = B
'                txtAux3(jj).Top = alto
'            Next jj
'            txtAux2(0).visible = B
'            txtAux2(0).Top = alto
'            cmdAux(0).visible = B
'            cmdAux(0).Top = txtAux3(2).Top
'            cmdAux(0).Height = txtAux3(2).Height
'
'        Case 2 'parcelas
'            For jj = 2 To txtAux4.Count - 1
'                If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
'                    If jj <> 6 Then
'                        txtAux4(jj).visible = B
'                        txtAux4(jj).Top = alto
'                    End If
'                Else
'                    txtAux4(jj).visible = B
'                    txtAux4(jj).Top = alto
'                End If
'            Next jj
'
'        Case 3 'seguros
'            For jj = 2 To txtAux5.Count - 1
'                txtAux5(jj).visible = B
'                txtAux5(jj).Top = alto
'            Next jj
'            txtAux2(2).visible = B
'            txtAux2(2).Top = alto
'
'            For jj = 2 To 4
'                cmdAux(jj).visible = B
'                cmdAux(jj).Top = txtAux5(3).Top
'                cmdAux(jj).Height = txtAux5(3).Height
'            Next jj
'
'            chkAux(0).visible = B
'            chkAux(0).Top = txtAux5(3).Top
'            chkAux(0).Height = txtAux5(3).Height
'
'        Case 4 'hco de campos
'            For jj = 2 To txtaux6.Count - 1
'                txtaux6(jj).visible = B
'                txtaux6(jj).Top = alto
'            Next jj
'            txtAux2(3).visible = B
'            txtAux2(3).Top = alto
'            txtAux2(4).visible = B
'            txtAux2(4).Top = alto
'
'            For jj = 5 To 8
'                cmdAux(jj).visible = B
'                cmdAux(jj).Top = txtaux6(2).Top
'                cmdAux(jj).Height = txtaux6(2).Height
'            Next jj
'
'        Case 5 'gastos de campos
'            For jj = 2 To txtaux7.Count - 1
'                txtaux7(jj).visible = B
'                txtaux7(jj).Top = alto
'            Next jj
'            txtAux2(5).visible = B
'            txtAux2(5).Top = alto
'
'            For jj = 9 To 10
'                cmdAux(jj).visible = B
'                cmdAux(jj).Top = txtaux7(2).Top
'                cmdAux(jj).Height = txtaux7(2).Height
'            Next jj
'
'        Case 6 'ordenes de recoleccion
'            For jj = 1 To txtaux8.Count - 1
'                txtaux8(jj).visible = B
'                txtaux8(jj).Top = alto
'            Next jj
'
'            For jj = 11 To 11
'                cmdAux(jj).visible = B
'                cmdAux(jj).Top = txtaux8(2).Top
'                cmdAux(jj).Height = txtaux8(2).Height
'            Next jj
'
'        Case 7 'revisiones de campos
'            For jj = 2 To txtAux9.Count - 1
'                txtAux9(jj).visible = B
'                txtAux9(jj).Top = alto
'            Next jj
'
'            For jj = 12 To 12
'                cmdAux(jj).visible = B
'                cmdAux(jj).Top = txtAux9(2).Top
'                cmdAux(jj).Height = txtAux9(2).Height
'            Next jj
'
'
'    End Select
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

Private Sub TxtAux3_LostFocus(Index As Integer)
'Dim cadMen As String
'Dim Nuevo As Boolean
'
'    If Not PerderFocoGnral(txtAux3(Index), Modo) Then Exit Sub
'
'    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
'    'mostrar mensajes ni hacer nada
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
'
'
'    ' ******* configurar el LostFocus dels camps de ll�nies (dins i fora grid) ********
'    Select Case Index
'        Case 2 'NIF
'            If PonerFormatoEntero(txtAux3(Index)) Then
'                txtAux2(0).Text = PonerNombreDeCod(txtAux3(Index), "rsocios", "nomsocio")
'                If txtAux2(0).Text = "" Then
'                    cadMen = "No existe el Socio: " & Text1(Index).Text & vbCrLf
'                    cadMen = cadMen & "�Desea crearlo?" & vbCrLf
'                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
'                        Set frmSoc1 = New frmManSocios
'                        frmSoc1.DatosADevolverBusqueda = "0|1|"
''                        frmVar.NuevoCodigo = Text1(Index).Text
'                        txtAux3(Index).Text = ""
'                        TerminaBloquear
'                        frmSoc1.Show vbModal
'                        Set frmSoc1 = Nothing
'                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
'                    Else
'                        txtAux3(Index).Text = ""
'                    End If
'                    PonerFoco txtAux3(Index)
'                Else
'                    ' comprobamos que el socio no est� dado de baja
'                    If Not EstaSocioDeAlta(txtAux3(Index).Text) Then
'                        If MsgBox("Este socio tiene fecha de baja. � Continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
'                            txtAux3(Index).Text = ""
'                            txtAux2(0).Text = ""
'                            PonerFoco txtAux3(Index)
'                        End If
'                    End If
'                End If
'            Else
'                txtAux2(0).Text = ""
'            End If
'
'        Case 3 'porcentaje de
'            PonerFormatoDecimal txtAux3(Index), 4
'            If txtAux3(2).Text <> "" Then cmdAceptar.SetFocus
'
'    End Select
'
'    ' ******************************************************************************
End Sub

Private Sub TxtAux3_GotFocus(Index As Integer)
'    If Not txtAux3(Index).MultiLine Then ConseguirFocoLin txtAux3(Index)
End Sub

Private Sub TxtAux3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    If Not txtAux3(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub TxtAux3_KeyPress(Index As Integer, KeyAscii As Integer)
'    If Not txtAux3(Index).MultiLine Then
'        If KeyAscii = teclaBuscar Then
'            If Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) Then
'                Select Case Index
'                    Case 2: KEYBusquedaAux KeyAscii, 0 'socio
'                End Select
'            End If
'        Else
'            KEYpress KeyAscii
'        End If
'    End If
End Sub


Private Function DatosOkLlin(nomframe As String) As Boolean
'Dim Rs As ADODB.Recordset
'Dim Sql As String
'Dim B As Boolean
'Dim Cant As Integer
'Dim Mens As String
'Dim vFact As Byte, vDocum As Byte
'
'    DatosOkLlin = True
'
'    On Error GoTo EDatosOKLlin
'
'    Mens = ""
'    DatosOkLlin = False
'
'    B = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
'    If Not B Then Exit Function
'
'    If B And (Modo = 5 And ModoLineas = 1) And nomframe = "FrameAux1" Then  'insertar
'        'comprobar si existe ya el cod. de la calidad para ese campo
'        Sql = ""
'        Sql = DevuelveDesdeBDNew(cAgro, "rcampos_clasif", "codcalid", "codcampo", txtAux1(0).Text, "N", , "codvarie", txtAux1(2).Text, "N", "codcalid", txtAux1(1).Text, "N")
'        If Sql <> "" Then
'            MsgBox "Ya existe la calidad para el campo.", vbExclamation
'            PonerFoco txtAux1(1)
'            B = False
'        End If
'    End If
'
'    If B And (Modo = 5 And ModoLineas = 1) And nomframe = "FrameAux0" Then  'insertar
'        'comprobar que el porcentaje sea distinto de cero
'        If txtAux3(3).Text = "" Then
'            MsgBox "El porcentaje de coopropiedad debe ser superior a 0.", vbExclamation
'            PonerFoco txtAux3(3)
'            B = False
'        Else
'            If CInt(txtAux3(3).Text) = 0 Then
'                MsgBox "El porcentaje de coopropiedad debe ser superior a 0.", vbExclamation
'                PonerFoco txtAux3(3)
'                B = False
'            End If
'        End If
'    End If
'
'
'
''
''    ' *** si cal fer atres comprovacions a les ll�nies (en o sense tab) ***
''    Select Case NumTabMto
''        Case 0  'CUENTAS BANCARIAS
''            SQL = "SELECT COUNT(ctaprpal) FROM cltebanc "
''            SQL = SQL & ObtenerWhereCab(True) & " AND ctaprpal=1"
''            If ModoLineas = 2 Then SQL = SQL & " AND numlinea<> " & AdoAux(NumTabMto).Recordset!numlinea
''            Set RS = New ADODB.Recordset
''            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
''            Cant = IIf(Not RS.EOF, RS.Fields(0).Value, 0)
''
''            RS.Close
''            Set RS = Nothing
'''yo
'''            'no n'hi ha cap conter principal i ha seleccionat que no
'''            If (Cant = 0) And (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 0) Then
'''                Mens = "Debe una haber una cuenta principal"
'''            ElseIf (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 1) And (cmbAux(9).ItemData(cmbAux(9).ListIndex) = 0) Then
'''                Mens = "Debe seleccionar que esta cuenta est� activa si desea que sea la principal"
'''            End If
''
'''            'No puede haber m�s de una cuenta principal
'''            If cant > 0 And (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 1) Then
'''                Mens = "No puede haber m�s de una cuenta principal."
'''            End If
'''yo
'''            'No pueden haber registros con el mismo: codbanco-codsucur-digcontr-ctabanc
'''            If Mens = "" Then
'''                SQL = "SELECT count(codclien) FROM cltebanc "
'''                SQL = SQL & " WHERE codclien=" & text1(0).Text & " AND codempre= " & vSesion.Empresa
'''                If ModoLineas = 2 Then SQL = SQL & " AND numlinea<> " & AdoAux(NumTabMto).Recordset!numlinea
'''                SQL = SQL & " AND codnacio=" & cmbAux(0).ItemData(cmbAux(0).ListIndex)
'''                SQL = SQL & " AND codbanco=" & txtaux(3).Text & " AND codsucur=" & txtaux(4).Text
'''                SQL = SQL & " AND digcontr='" & txtaux(5).Text & "' AND ctabanco='" & txtaux(6).Text & "'"
'''                Set RS = New ADODB.Recordset
'''                RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'''                Cant = IIf(Not RS.EOF, RS.Fields(0).Value, 0)
'''                If Cant > 0 Then
'''                    Mens = "Ya Existe la cuenta bancaria: " & cmbAux(0).List(cmbAux(0).ListIndex) & " - " & txtaux(3).Text & "-" & txtaux(4).Text & "-" & txtaux(5).Text & "-" & txtaux(6).Text
'''                End If
'''                RS.Close
'''                Set RS = Nothing
'''            End If
'''
'''            If Mens <> "" Then
'''                Screen.MousePointer = vbNormal
'''                MsgBox Mens, vbExclamation
'''                DatosOkLlin = False
'''                'PonerFoco txtAux(3)
'''                Exit Function
'''            End If
'''
''    End Select
''    ' ******************************************************************************
'    DatosOkLlin = B
'
'EDatosOKLlin:
'    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
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

' *** si n'hi han formularis de buscar codi a les ll�nies ***
Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
    
    Select Case Index
        Case 0 'situacion
            Set frmSit = New frmManSituCamp
            frmSit.DatosADevolverBusqueda = "0|1|"
            frmSit.CodigoActual = Text1(6).Text
            frmSit.DeConsulta = False
            frmSit.Show vbModal
            Set frmSit = Nothing
            PonerFoco Text1(6)
        
       Case 1 'Socios
            Indice = 1
            Set frmSoc = New frmManSocios
'            frmSoc.DeConsulta = True
            frmSoc.DatosADevolverBusqueda = "0|1|"
'            frmSoc.CodigoActual = Text1(1).Text
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco Text1(Indice)
    
       Case 2 'Partidas
            Set frmPar = New frmManPartidas
            frmPar.DeConsulta = True
            frmPar.DatosADevolverBusqueda = "0|1|2|3|4|5|"
            frmPar.CodigoActual = Text1(2).Text
            frmPar.Show vbModal
            Set frmPar = Nothing
            PonerFoco Text1(2)
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


Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
'Dim B As Boolean
'Dim I As Byte
'Dim tots As String
'Dim Sql2 As String
'
'
'    On Error GoTo ECarga
'
'    tots = MontaSQLCarga(Index, enlaza)
'    Sql2 = tots
'
'    B = DataGridAux(Index).Enabled
'    DataGridAux(Index).Enabled = False
'
'    Adoaux(Index).ConnectionString = conn
'    Adoaux(Index).RecordSource = Sql2
'    Adoaux(Index).CursorType = adOpenDynamic
'    Adoaux(Index).LockType = adLockPessimistic
'    DataGridAux(Index).ScrollBars = dbgNone
'    Adoaux(Index).Refresh
'    Set DataGridAux(Index).DataSource = Adoaux(Index)
'
'    DataGridAux(Index).AllowRowSizing = False
'    DataGridAux(Index).RowHeight = 290
'    If PrimeraVez Then
'        DataGridAux(Index).ClearFields
'        DataGridAux(Index).ReBind
'        DataGridAux(Index).Refresh
'    End If
'
'    For I = 0 To DataGridAux(Index).Columns.Count - 1
'        DataGridAux(Index).Columns(I).AllowSizing = False
'    Next I
'
'    CargaGridGnral Me.DataGridAux(Index), Me.Adoaux(Index), tots, PrimeraVez
'
'
'    'DataGridAux(Index).Enabled = b
''    PrimeraVez = False
'
'    Select Case Index
'        Case 1 'clasificacion segun la calidad
'            'si es visible|control|tipo campo|nombre campo|ancho control|
'            tots = "N||||0|;N||||0|;S|txtaux1(1)|T|C�d.|800|;S|cmdAux(1)|B|||;" 'codsocio,codsecci
'            tots = tots & "S|txtAux2(1)|T|Nombre|3870|;"
'            tots = tots & "S|txtaux1(3)|T|Muestra|1200|;"
'            arregla tots, DataGridAux(Index), Me, 350
'
'
'            DataGridAux(Index).Columns(4).Alignment = dbgRight
'
'            B = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
'
'            If (enlaza = True) And (Not Adoaux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
'                If VisualizaClasificacion Then
'                    PonerClasificacionGrafica
'                End If
'            Else
'                For I = 0 To 3
'                    txtAux1(I).Text = ""
'                Next I
'                txtAux2(1).Text = ""
'                Me.MSChart1.visible = False
'            End If
'
'        Case 0 ' coopropietarios
'            tots = "N||||0|;N||||0|;S|txtaux3(2)|T|C�digo|1000|;S|cmdAux(0)|B|||;" 'codsocio,numlinea
'            tots = tots & "S|txtAux2(0)|T|Nombre|4070|;"
'            tots = tots & "S|txtaux3(3)|T|Porcentaje|1400|;"
'            arregla tots, DataGridAux(Index), Me, 350
'
'            DataGridAux(Index).Columns(4).Alignment = dbgRight
'
'            B = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
'
'            If (enlaza = True) And (Not Adoaux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
''                SumaTotalPorcentajes
'            Else
'                For I = 0 To 3
'                    txtAux3(I).Text = ""
'                Next I
'                txtAux2(0).Text = ""
'            End If
'
'
'        Case 2 'parcelas del campo
'            'si es visible|control|tipo campo|nombre campo|ancho control|
'            If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
'                tots = "N||||0|;N||||0|;S|txtaux4(2)|T|Poligono|1225|;" 'codsocio,codsecci
'                tots = tots & "S|txtaux4(3)|T|Parcela|1225|;"
'                tots = tots & "S|txtaux4(4)|T|Subrecinto|1225|;"
'                tots = tots & "S|txtaux4(5)|T|Recinto|1225|;"
'                tots = tots & "N|txtaux4(6)|T|CodSigpac|1000|;"
'                tots = tots & "S|txtaux4(7)|T|Has.Parc.Sig|2100|;"
'                tots = tots & "S|txtaux4(8)|T|Has.Recinto|2100|;"
'                tots = tots & "S|txtaux4(9)|T|Has.Catastro|2100|;"
'                tots = tots & "S|txtaux4(10)|T|Has.Cult.Recinto|2100|;"
'            Else
'                tots = "N||||0|;N||||0|;S|txtaux4(2)|T|Poligono|1225|;" 'codsocio,codsecci
'                tots = tots & "S|txtaux4(3)|T|Parcela|1225|;"
'                tots = tots & "S|txtaux4(4)|T|Subparcela|1225|;"
'                tots = tots & "S|txtaux4(5)|T|Recinto|1225|;"
'                tots = tots & "S|txtaux4(6)|T|CodSigpac|1000|;"
'                tots = tots & "S|txtaux4(7)|T|Has.Sigpac|1850|;"
'                tots = tots & "S|txtaux4(8)|T|Has.Cult.Sigpac|1850|;"
'                tots = tots & "S|txtaux4(9)|T|Has.Catastro|1850|;"
'                tots = tots & "S|txtaux4(10)|T|Has.Cult.Catastro|1850|;"
'            End If
'            arregla tots, DataGridAux(Index), Me, 350
'
'            DataGridAux(Index).Columns(4).Alignment = dbgRight
'
'            B = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
'
'            If (enlaza = True) And (Not Adoaux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
'
'            Else
'                For I = 0 To 3
'                    txtAux4(I).Text = ""
'                Next I
'            End If
'
'            CalcularTotalSuperficie Sql2
'
'        Case 3 'incidencias de seguro
'            'si es visible|control|tipo campo|nombre campo|ancho control|
'            tots = "N||||0|;N||||0|;S|txtaux5(2)|T|Fecha|1400|;S|cmdAux(3)|B|||;" 'codcampo,numlinea
'            tots = tots & "S|txtaux5(3)|T|Incidencia|1100|;S|cmdAux(2)|B|||;"
'            tots = tots & "S|txtAux2(2)|T|Descripcion|3600|;"
'            '[Monica]26/01/2016: nueva columna de kilos aportacion, cambio etiqueta de los kilos a indemnizables
'            tots = tots & "S|txtaux5(4)|T|Kilos Indemnizaci�n|2250|;"
'            tots = tots & "S|txtaux5(7)|T|Kilos Aportacion|1850|;"
'
'            tots = tots & "S|txtaux5(5)|T|Importe|1600|;"
'            tots = tots & "S|txtaux5(6)|T|Fecha Pago|1400|;S|cmdAux(4)|B|||;"
'            tots = tots & "N||||0|;S|chkAux(0)|CB|Sin|360|;"
'
'            arregla tots, DataGridAux(Index), Me, 350
'
'            DataGridAux(Index).Columns(4).Alignment = dbgLeft
'            DataGridAux(Index).Columns(5).Alignment = dbgRight
'            DataGridAux(Index).Columns(6).Alignment = dbgRight
'
'            B = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
'
'            If (enlaza = True) And (Not Adoaux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
'
'            Else
'                For I = 0 To 3
'                    txtAux5(I).Text = ""
'                Next I
'            End If
'
'
'        Case 4 'hco del campo
'            'si es visible|control|tipo campo|nombre campo|ancho control|
'            tots = "N||||0|;N||||0|;" 'codcampo,numlinea
'            tots = tots & "S|txtaux6(2)|T|Socio|1100|;S|cmdAux(8)|B|||;"
'            tots = tots & "S|txtAux2(4)|T|Nombre|3500|;"
'            tots = tots & "S|txtaux6(3)|T|Fecha Alta|1400|;S|cmdAux(6)|B|||;"
'            tots = tots & "S|txtaux6(4)|T|Fecha Baja|1400|;S|cmdAux(5)|B|||;"
'            tots = tots & "S|txtaux6(5)|T|Incidencia|1300|;S|cmdAux(7)|B|||;"
'            tots = tots & "S|txtAux2(3)|T|Descripcion|3400|;"
'
'            arregla tots, DataGridAux(Index), Me, 350
'
'            DataGridAux(Index).Columns(2).Alignment = dbgLeft
'            DataGridAux(Index).Columns(3).Alignment = dbgLeft
'
'            B = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
'
'            If (enlaza = True) And (Not Adoaux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
'
'            Else
'                For I = 0 To 5
'                    txtaux6(I).Text = ""
'                Next I
'            End If
'
'        Case 5 'gastos del campo
'            'si es visible|control|tipo campo|nombre campo|ancho control|
'            tots = "N||||0|;N||||0|;" 'codcampo,numlinea
'            tots = tots & "S|txtAux7(2)|T|C�digo|900|;S|cmdAux(9)|B|||;"
'            tots = tots & "S|txtAux2(5)|T|Concepto|4400|;"
'            tots = tots & "S|txtAux7(3)|T|Fecha|1400|;S|cmdAux(10)|B|||;"
'            tots = tots & "S|txtAux7(4)|T|Importe|1500|;"
'            tots = tots & "N||||0|;S|chkAux(1)|CB|Id|360|;"
'
'            arregla tots, DataGridAux(Index), Me, 350
'
'            DataGridAux(Index).Columns(2).Alignment = dbgLeft
'            DataGridAux(Index).Columns(3).Alignment = dbgLeft
'
'            B = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
'
'            If (enlaza = True) And (Not Adoaux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
'
'            Else
'                For I = 0 To 4
'                    txtaux7(I).Text = ""
'                Next I
'            End If
'
'        Case 6  'ordenes de recoleccion
'            'si es visible|control|tipo campo|nombre campo|ancho control|
'            tots = "N||||0|;" 'codcampo
'            tots = tots & "S|txtAux8(1)|T|Nro.Orden|1100|;"
'            tots = tots & "S|txtAux8(2)|T|Fecha|1100|;S|cmdAux(11)|B|||;"
'
'            arregla tots, DataGridAux(Index), Me, 350
'
'
'            B = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
'
'            If (enlaza = True) And (Not Adoaux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
'
'            Else
'                For I = 0 To 0
'                    txtaux8(I).Text = ""
'                Next I
'            End If
'
'        Case 7 ' revisiones
'            'si es visible|control|tipo campo|nombre campo|ancho control|
'            tots = "N||||0|;N||||0|;" 'codcampo,numlinea
'            tots = tots & "S|txtAux9(2)|T|Fecha|1400|;S|cmdAux(12)|B|||;"
'            tots = tots & "S|txtAux9(3)|T|T�cnico|4500|;"
'            tots = tots & "S|txtAux9(4)|T|Observaciones|6100|;"
'
'            arregla tots, DataGridAux(Index), Me, 350
'
'            DataGridAux(Index).Columns(2).Alignment = dbgLeft
'            DataGridAux(Index).Columns(3).Alignment = dbgLeft
'
'            B = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
'
'            If (enlaza = True) And (Not Adoaux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
'
'            Else
'                For I = 0 To txtAux9.Count - 1
'                    txtAux9(I).Text = ""
'                Next I
'            End If
'
'
'    End Select
'    DataGridAux(Index).ScrollBars = dbgAutomatic
'
'    ' **** si n'hi han ll�nies en grids i camps fora d'estos ****
'    If Not Adoaux(Index).Recordset.EOF Then
'        DataGridAux_RowColChange Index, 1, 1
'    Else
'        LimpiarCamposFrame Index
'    End If
'    ' **********************************************************
'
'ECarga:
'    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub


Private Sub InsertarLinea()
''Inserta registre en les taules de Ll�nies
'Dim nomframe As String
'Dim B As Boolean
'
'    On Error Resume Next
'
'    ' *** posa els noms del frames, tant si son de grid com si no ***
'    Select Case NumTabMto
'        Case 0: nomframe = "FrameAux0" 'coopropietarios
'        Case 1: nomframe = "FrameAux1" 'clasificacion
'        Case 2: nomframe = "FrameAux2" 'parcelas
'        Case 3: nomframe = "FrameAux3" 'agroseguro
'        Case 4: nomframe = "FrameAux4" 'hco
'        Case 5: nomframe = "FrameAux5" 'concepto de gastos
'        Case 6: nomframe = "FrameAux6" 'ordenes de recoleccion
'        Case 7: nomframe = "FrameAux7" 'revision de campos
'    End Select
'    ' ***************************************************************
'
'    If DatosOkLlin(nomframe) Then
'        TerminaBloquear
'        If InsertarDesdeForm2(Me, 2, nomframe) Then
'            ' *** si n'hi ha que fer alguna cosa abas d'insertar
'            ' *************************************************
'            B = BLOQUEADesdeFormulario2(Me, Data1, 1)
'
'            '++monica: en caso de estar insertando seccion y que no existan las
'            'cuentas contables hacemos esto para que las inserte en contabilidad.
''            If NumTabMto = 1 Then
''               txtAux2(4).Text = PonerNombreCuenta(txtaux1(4), 3, Text1(0))
''               txtAux2(5).Text = PonerNombreCuenta(txtaux1(5), 3, Text1(0))
''            End If
'
'            Select Case NumTabMto
'                Case 0, 1, 2, 3, 4, 6, 7 ' *** els index de les llinies en grid (en o sense tab) ***
'                    CargaGrid NumTabMto, True
'                    If B Then BotonAnyadirLinea NumTabMto
'
'                Case 5 ' Caso de gastos de campo, tenemos que insertar un asiento en el diario
'                    Screen.MousePointer = vbHourglass
'
'                    frmListado.OpcionListado = 37
'                    frmListado.NumCod = "rcampos_gastos.codcampo = " & DBSet(txtaux7(0).Text, "N") & " and rcampos_gastos.numlinea = " & DBSet(txtaux7(1).Text, "N")
'                    frmListado.Show vbModal
'                    CargaGrid NumTabMto, True
'
'                    Screen.MousePointer = vbDefault
'
'                    CargaGrid NumTabMto, True
'                    If B Then BotonAnyadirLinea NumTabMto
'
'            End Select
'
'            'SituarTab (NumTabMto)
'        End If
'    End If
End Sub


Private Sub ModificarLinea()
''Modifica registre en les taules de Ll�nies
'Dim nomframe As String
'Dim V As Integer
'Dim cad As String
'    On Error Resume Next
'
'    ' *** posa els noms del frames, tant si son de grid com si no ***
'    Select Case NumTabMto
'        Case 0: nomframe = "FrameAux0" 'coopropietarios
'        Case 1: nomframe = "FrameAux1" 'secciones
'        Case 2: nomframe = "FrameAux2" 'parcelas
'        Case 3: nomframe = "FrameAux3" 'seguros
'        Case 4: nomframe = "FrameAux4" 'hco
'        Case 5: nomframe = "FrameAux5" 'conceptos de gastos
'        Case 7: nomframe = "FrameAux7" 'revisiones de campos
'
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
''            If NumTabMto <> 3 Then
'                V = Adoaux(NumTabMto).Recordset.Fields(1) 'el 2 es el n� de llinia
'                CargaGrid NumTabMto, True
''            End If
'
'            ' *** si n'hi han tabs ***
'            'SituarTab (NumTabMto)
'
'            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
''            If NumTabMto <> 3 Then
'                DataGridAux(NumTabMto).SetFocus
'                Adoaux(NumTabMto).Recordset.Find (Adoaux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
''            End If
'            ' ***********************************************************
'
'            LLamaLineas NumTabMto, 0
'
'        End If
'    End If
'
End Sub


Private Sub Modificar()
'Modifica registre en les taules de Ll�nies
Dim nomframe As String
Dim V As Integer
Dim cad As String
Dim Sql As String
Dim vCadena As String
Dim Produ As Integer

    On Error GoTo eModificar

    conn.BeginTrans

    
    B = True
        
    ' modificamos los datos del campo
    If B Then
        If ModificaDesdeFormulario2(Me, 1) Then
            TerminaBloquear
            
            PosicionarData
        
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
    ' *** canviar-ho per la clau primaria de la cap�alera ***
    vWhere = vWhere & " rcampos_clasif.codcampo=" & Val(Text1(0).Text)
    vWhere = vWhere & " and rcampos_clasif.codvarie = " & Val(Text1(2).Text)
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function

Private Function ObtenerWhereCab2(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la cap�alera ***
    vWhere = vWhere & " rcampos_parcelas.codcampo=" & Val(Text1(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab2 = vWhere
End Function

Private Function ObtenerWhereCab3(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la cap�alera ***
    vWhere = vWhere & " rcampos_seguros.codcampo=" & Val(Text1(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab3 = vWhere
End Function


Private Function ObtenerWhereCab4(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la cap�alera ***
    vWhere = vWhere & " rcampos_hco.codcampo=" & Val(Text1(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab4 = vWhere
End Function


Private Function ObtenerWhereCab5(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la cap�alera ***
    vWhere = vWhere & " rcampos_gastos.codcampo=" & Val(Text1(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab5 = vWhere
End Function


Private Function ObtenerWhereCab6(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la cap�alera ***
    vWhere = vWhere & " rcampos_ordrec.codcampo=" & Val(Text1(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab6 = vWhere
End Function

Private Function ObtenerWhereCab7(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la cap�alera ***
    vWhere = vWhere & " rcampos_revision.codcampo=" & Val(Text1(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab7 = vWhere
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

'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del rat�n.
'Private Sub DataGridAux_GotFocus(Index As Integer)
'  WheelHook DataGridAux(Index)
'End Sub
'Private Sub DataGridAux_LostFocus(Index As Integer)
'  WheelUnHook
'End Sub

' ********* si n'hi han combos a la cap�alera ************
Private Sub CargaCombo()
'Dim Ini As Integer
'Dim Fin As Integer
'Dim I As Integer
'Dim Sql As String
'Dim Rs As ADODB.Recordset
'
'   ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
'    For I = 0 To Combo1.Count - 1
'        Combo1(I).Clear
'    Next I
'
'    'tipo de parcela
'    Combo1(0).AddItem "R�stica"
'    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
'    Combo1(0).AddItem "Urbana"
'    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
'
'    'tipo de recoleccion
'    Combo1(1).AddItem "Cooperativa"
'    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
'    Combo1(1).AddItem "Socio"
'    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
'
'    'tipo de campo
'    Combo1(3).AddItem "Normal"
'    Combo1(3).ItemData(Combo1(3).NewIndex) = 0
'    Combo1(3).AddItem "Comercio"
'    Combo1(3).ItemData(Combo1(3).NewIndex) = 1
'    Combo1(3).AddItem "Industria"
'    Combo1(3).ItemData(Combo1(3).NewIndex) = 2
'
'
'    'TIPO DE SISTEMA DE RIEGO
'    Sql = "select codriego, nomriego from rriego "
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    I = 0
'    While Not Rs.EOF
''        Sql = Replace(Rs.Fields(1).Value, "Factura", "Fac.")
'        Sql = Rs.Fields(1).Value
''        Sql = Rs.Fields(0).Value & " - " & Sql
'        Combo1(2).AddItem Sql 'campo del codigo
'        Combo1(2).ItemData(Combo1(2).NewIndex) = I
'        I = I + 1
'        Rs.MoveNext
'    Wend
'
'    ' Entrega Ficha de Cultivo
'    Sql = "select codtipo, descripcion from rfichculti "
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    I = 0
'    While Not Rs.EOF
'        Sql = Rs.Fields(1).Value
'        Combo1(4).AddItem Sql 'campo del codigo
'        Combo1(4).ItemData(Combo1(4).NewIndex) = I
'        I = I + 1
'        Rs.MoveNext
'    Wend
'
'
End Sub


Private Sub txtAux1_LostFocus(Index As Integer)
'Dim cadMen As String
'Dim Nuevo As Boolean
'
'    If Not PerderFocoGnral(txtAux1(Index), Modo) Then Exit Sub
'
'    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
'    'mostrar mensajes ni hacer nada
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
'
'
'    ' ******* configurar el LostFocus dels camps de ll�nies (dins i fora grid) ********
'    Select Case Index
'        Case 1 ' calidad
'            If PonerFormatoEntero(txtAux1(Index)) Then
'                txtAux2(Index).Text = PonerNombreDeCod(txtAux1(Index), "rcalidad", "nomcalid", "codcalid", "N", , "codvarie", txtAux1(2).Text, "N")
'                If txtAux2(Index).Text = "" Then
'                    cadMen = "No existe la Calidad: " & txtAux1(Index).Text & vbCrLf
'                    cadMen = cadMen & "�Desea crearla?" & vbCrLf
'                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
'                        Set frmCalid = New frmManCalidades
'                        frmCalid.DatosADevolverBusqueda = "0|1|"
'                        frmCalid.NuevoCodigo = txtAux1(Index).Text
'                        txtAux1(Index).Text = ""
'                        TerminaBloquear
'                        frmCalid.Show vbModal
'                        Set frmCalid = Nothing
'                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
'                    Else
'                        txtAux1(Index).Text = ""
'                    End If
'                    PonerFoco txtAux1(Index)
'                End If
'            Else
'                txtAux2(Index).Text = ""
'            End If
'
'        Case 3 ' muestra debe sumar el 100%
'            If PonerFormatoDecimal(txtAux1(Index), 4) Then
'                cmdAceptar.SetFocus
'            End If
'
''        Case 2, 3 'fecha de alta y de baja
''            PonerFormatoFecha txtaux1(Index)
'
''        Case 4, 5 'cta Cliente y Proveedor
''            If txtaux1(Index).Text = "" Then Exit Sub
''
''            If Not vSeccion Is Nothing Then
''                txtAux2(Index).Text = PonerNombreCuenta(txtaux1(Index), Modo)
''                If txtaux1(Index).Text <> "" Then
''                    If Not vSeccion.CtaConRaizCorrecta(txtaux1(Index).Text, Index - 4) Then
''                        MsgBox "La cuenta no tiene la raiz correcta. Revise.", vbExclamation
''                    Else
''                        ' si la cuenta es correcta y no existe la insertamos en contabilidad
''                        txtAux2(Index).Text = PonerNombreCuenta(txtaux1(Index), 3, Text1(0))
''                    End If
''                End If
''            End If
''
''        Case 6 'codigo iva
''            If txtaux1(Index).Text = "" Then Exit Sub
''
''            If Not vSeccion Is Nothing Then
''                  txtAux2(0).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", txtaux1(Index).Text, "N")
''            End If
''            cmdAceptar.SetFocus
'
'    End Select
'
'    ' ******************************************************************************
End Sub

Private Sub txtAux1_GotFocus(Index As Integer)
'   If Not txtAux1(Index).MultiLine Then ConseguirFocoLin txtAux1(Index)
End Sub

Private Sub txtAux1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    If Not txtAux1(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub txtAux1_KeyPress(Index As Integer, KeyAscii As Integer)
'    If Not txtAux1(Index).MultiLine Then
'        If KeyAscii = teclaBuscar Then
'            If Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) Then
'                Select Case Index
'                    Case 1: KEYBusquedaAux KeyAscii, 1 'calidad
'                End Select
'            End If
'        Else
'            KEYpress KeyAscii
'        End If
'    End If
End Sub


Private Sub PonerDatosPartida()
Dim Zona As String
Dim OtroCampo As String
Dim CodPobla As String

    Zona = ""
    Text5(3).Text = ""
    Text4(3).Text = ""
    
    OtroCampo = "codpobla"
    Zona = DevuelveDesdeBDNew(cAgro, "rpartida", "codzonas", "codparti", Text1(2), "N", OtroCampo)
    
    If Zona <> "" Then
        If OtroCampo <> "" Then
            CodPobla = OtroCampo
            Text4(3).Text = CodPobla
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





Private Function VisualizaClasificacion() As Boolean
Dim Sql As String


    If Data1.Recordset.EOF Then
        VisualizaClasificacion = False
        Exit Function
    End If

    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "variedades", "tipoclasifica", "codvarie", Data1.Recordset!Codvarie, "N")
    
    SSTab1.TabEnabled(3) = (Sql = "0")
    SSTab1.TabVisible(3) = (Sql = "0")
    
    VisualizaClasificacion = (Sql = "0")

End Function


Private Sub TxtAux4_GotFocus(Index As Integer)
'    If Not txtAux4(Index).MultiLine Then ConseguirFocoLin txtAux4(Index)
End Sub

Private Sub TxtAux4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    If Not txtAux4(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub TxtAux4_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub TxtAux4_LostFocus(Index As Integer)
'Dim cadMen As String
'Dim Nuevo As Boolean
'
'    If Not PerderFocoGnral(txtAux4(Index), Modo) Then Exit Sub
'
'    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
'    'mostrar mensajes ni hacer nada
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
'
'
'    ' ******* configurar el LostFocus dels camps de ll�nies (dins i fora grid) ********
'    Select Case Index
'        Case 2, 3 'poligono y parcela
'            PonerFormatoEntero txtAux4(Index)
'
'        Case 5 'recinto
'            PonerFormatoEntero txtAux4(Index)
'
'        Case 6 'COD SIGPAC
'            PonerFormatoEntero txtAux4(Index)
'
'        Case 7, 8, 9, 10 'superficies en hectareas
'            If Modo = 1 Then Exit Sub
'            If PonerFormatoDecimal(txtAux4(Index), 7) Then
'                If Index = 10 Then cmdAceptar.SetFocus
'            Else
'                If Index = 10 And txtAux4(Index) = "" Then cmdAceptar.SetFocus
'            End If
'
'
'    End Select
'
'    ' ******************************************************************************
End Sub

'*******************************
Private Sub TxtAux5_GotFocus(Index As Integer)
'    If Not txtAux5(Index).MultiLine Then ConseguirFocoLin txtAux5(Index)
End Sub

Private Sub TxtAux5_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    If Not txtAux5(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub TxtAux5_KeyPress(Index As Integer, KeyAscii As Integer)
'    If Not txtAux5(Index).MultiLine Then
'        If KeyAscii = teclaBuscar Then
'            If Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) Then
'                Select Case Index
'                    Case 2: KEYBusquedaAux KeyAscii, 3 'fecha
'                    Case 3: KEYBusquedaAux KeyAscii, 2 'incidencia
'                    Case 6: KEYBusquedaAux KeyAscii, 4 'fecha de pago
'                End Select
'            End If
'        Else
'            KEYpress KeyAscii
'        End If
'    End If

End Sub

Private Sub TxtAux5_LostFocus(Index As Integer)
'Dim cadMen As String
'Dim Nuevo As Boolean
'
'    If Not PerderFocoGnral(txtAux5(Index), Modo) Then Exit Sub
'
'    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
'    'mostrar mensajes ni hacer nada
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
'
'    ' ******* configurar el LostFocus dels camps de ll�nies (dins i fora grid) ********
'    Select Case Index
'        Case 2, 6 'fecha y fecha de pago
'            PonerFormatoFecha txtAux5(Index)
'
'        Case 3 ' codigo de incidencia
'            If PonerFormatoEntero(txtAux5(Index)) Then
'                txtAux2(2).Text = PonerNombreDeCod(txtAux5(Index), "rincidencia", "nomincid", "codincid", "N")
'                If txtAux2(2).Text = "" Then
'                    cadMen = "No existe la Incidencia: " & txtAux5(Index).Text & vbCrLf
'                    cadMen = cadMen & "�Desea crearla?" & vbCrLf
'                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
'                        Set frmInc = New frmManInciden
'                        frmInc.DatosADevolverBusqueda = "0|1|"
'                        frmInc.NuevoCodigo = txtAux1(Index).Text
'                        txtAux5(Index).Text = ""
'                        TerminaBloquear
'                        frmInc.Show vbModal
'                        Set frmInc = Nothing
'                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
'                    Else
'                        txtAux5(Index).Text = ""
'                    End If
'                    PonerFoco txtAux5(Index)
'                End If
'            Else
'                txtAux2(2).Text = ""
'            End If
'
'        Case 4 'kilos
'            PonerFormatoEntero txtAux5(Index)
'
'        '[Monica]26/01/2016: nueva columna de kilos aportacion
'        Case 7 ' kilos aportacion
'            PonerFormatoEntero txtAux5(Index)
'
'        Case 5 ' importe
'            If Modo = 1 Then Exit Sub
'            PonerFormatoDecimal txtAux5(Index), 1
'
'    End Select
'
'    ' ******************************************************************************
End Sub




'*******************************

Private Sub BotonCambioSocio()
Dim Sql As String
Dim campo As String
Dim NroContadores As Long

    If Text1(5).Text <> "" Then
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
'    If PorHectareas Then
'        Text2(5).TabIndex = 83
'        Text2(6).TabIndex = 84
'        Text2(7).TabIndex = 85
'        Text2(33).TabIndex = 86
'
'        Text1(5).TabIndex = 8
'        Text1(6).TabIndex = 9
'        Text1(7).TabIndex = 10
'        Text1(33).TabIndex = 11
'    Else
'        Text2(5).TabIndex = 8
'        Text2(6).TabIndex = 9
'        Text2(7).TabIndex = 10
'        Text2(33).TabIndex = 11
'
'        Text1(5).TabIndex = 83
'        Text1(6).TabIndex = 84
'        Text1(7).TabIndex = 85
'        Text1(33).TabIndex = 86
'    End If
End Sub



