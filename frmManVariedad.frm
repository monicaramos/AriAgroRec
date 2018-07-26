VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmManVariedad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Variedades"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   17385
   Icon            =   "frmManVariedad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   17385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   240
      TabIndex        =   133
      Top             =   0
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   134
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
      Left            =   3900
      TabIndex        =   131
      Top             =   0
      Width           =   975
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   210
         TabIndex        =   132
         Top             =   180
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Copia Calidades/Calibres"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   4980
      TabIndex        =   129
      Top             =   0
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   130
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
      Index           =   0
      Left            =   14085
      TabIndex        =   128
      Top             =   285
      Width           =   1605
   End
   Begin VB.Frame Frame2 
      Height          =   765
      Index           =   0
      Left            =   240
      TabIndex        =   31
      Top             =   690
      Width           =   16815
      Begin VB.TextBox text1 
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
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "Código de variedad|N|N|0|999999|variedades|codvarie|000000|S|"
         Top             =   240
         Width           =   885
      End
      Begin VB.TextBox text1 
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
         Left            =   4290
         MaxLength       =   20
         TabIndex        =   1
         Tag             =   "Nombre|T|N|||variedades|nomvarie|||"
         Top             =   240
         Width           =   8235
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
         Left            =   3390
         TabIndex        =   33
         Top             =   270
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "Código Variedad"
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
         Left            =   330
         TabIndex        =   32
         Top             =   270
         Width           =   2040
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   270
      TabIndex        =   28
      Top             =   8820
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
         TabIndex        =   29
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
      Left            =   16020
      TabIndex        =   22
      Top             =   8940
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
      Left            =   14850
      TabIndex        =   21
      Top             =   8940
      Width           =   1065
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7290
      Left            =   225
      TabIndex        =   30
      Top             =   1530
      Width           =   16830
      _ExtentX        =   29686
      _ExtentY        =   12859
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
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
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmManVariedad.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(26)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label18"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "imgBuscar(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "imgBuscar(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label20"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "imgBuscar(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "imgBuscar(4)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label26"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "imgBuscar(9)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label27"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label28"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "imgBuscar(10)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "imgAyuda(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(2)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(3)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "imgAyuda(1)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label39"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label40"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "imgFec(0)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "imgFec(1)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "text2(2)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "text1(2)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "text1(4)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "text1(3)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "text2(3)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "text1(9)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "text2(9)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Frame3"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "text2(26)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "text1(26)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "text1(27)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "text2(27)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Frame4"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Frame6"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Combo1(1)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Combo1(2)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "text1(38)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "text1(39)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "text1(40)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).ControlCount=   41
      TabCaption(1)   =   "Calibres"
      TabPicture(1)   =   "frmManVariedad.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameAux0"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Calidades"
      TabPicture(2)   =   "frmManVariedad.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameAux1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Recolección"
      TabPicture(3)   =   "frmManVariedad.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "chkComer(0)"
      Tab(3).Control(1)=   "Frame5"
      Tab(3).Control(2)=   "text1(30)"
      Tab(3).Control(3)=   "Combo1(0)"
      Tab(3).Control(4)=   "text1(20)"
      Tab(3).Control(5)=   "text1(19)"
      Tab(3).Control(6)=   "text1(18)"
      Tab(3).Control(7)=   "text1(17)"
      Tab(3).Control(8)=   "text1(16)"
      Tab(3).Control(9)=   "text1(15)"
      Tab(3).Control(10)=   "text1(14)"
      Tab(3).Control(11)=   "text1(13)"
      Tab(3).Control(12)=   "text1(12)"
      Tab(3).Control(13)=   "text1(8)"
      Tab(3).Control(14)=   "text1(11)"
      Tab(3).Control(15)=   "text1(10)"
      Tab(3).Control(16)=   "text1(7)"
      Tab(3).Control(17)=   "text1(6)"
      Tab(3).Control(18)=   "text1(5)"
      Tab(3).Control(19)=   "Label31"
      Tab(3).Control(20)=   "Label1(19)"
      Tab(3).Control(21)=   "Label17"
      Tab(3).Control(22)=   "Label16"
      Tab(3).Control(23)=   "Label15"
      Tab(3).Control(24)=   "Label14"
      Tab(3).Control(25)=   "Label13"
      Tab(3).Control(26)=   "Label12"
      Tab(3).Control(27)=   "Label11"
      Tab(3).Control(28)=   "Label10"
      Tab(3).Control(29)=   "Label9"
      Tab(3).Control(30)=   "Label3"
      Tab(3).Control(31)=   "Label2"
      Tab(3).Control(32)=   "Label8"
      Tab(3).Control(33)=   "Label19"
      Tab(3).Control(34)=   "Label6"
      Tab(3).Control(35)=   "Label7"
      Tab(3).ControlCount=   36
      TabCaption(4)   =   "Variedades Relacionadas"
      TabPicture(4)   =   "frmManVariedad.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "FrameAux2"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Bonificaciones"
      TabPicture(5)   =   "frmManVariedad.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "FrameAux3"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Calibradores"
      TabPicture(6)   =   "frmManVariedad.frx":00B4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "FrameAux4"
      Tab(6).ControlCount=   1
      Begin VB.Frame FrameAux4 
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   -74880
         TabIndex        =   161
         Top             =   500
         Width           =   11940
         Begin VB.TextBox txtAux4 
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   2
            Left            =   2745
            MaxLength       =   3
            TabIndex        =   171
            Tag             =   "Linea|N|N|||rcalidad_calibrador|numlinea|000|S|"
            Text            =   "lin"
            Top             =   3735
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.TextBox txtAux4 
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   1
            Left            =   630
            MaxLength       =   2
            TabIndex        =   170
            Tag             =   "Codigo Calidad|N|S|0|99|rcalidad_calibrador|codcalid|00|S|"
            Text            =   "Ca"
            Top             =   3690
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.TextBox txtAux4 
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   3
            Left            =   3840
            MaxLength       =   30
            TabIndex        =   172
            Tag             =   "Nom.Calibrador 1|T|N|||rcalidad_calibrador|nomcalibrador1|||"
            Text            =   "nom.calibrador 1"
            Top             =   3720
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.TextBox txtAux4 
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   4
            Left            =   4830
            MaxLength       =   30
            TabIndex        =   173
            Tag             =   "Nom.Calibrador 2|T|N|||rcalidad_calibrador|nomcalibrador2|||"
            Text            =   "nom.calibrador 2"
            Top             =   3720
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.TextBox txtAux2 
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   1
            Left            =   1680
            TabIndex        =   169
            Top             =   3690
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.CommandButton btnBuscar 
            Appearance      =   0  'Flat
            Caption         =   "+"
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
            Height          =   330
            Index           =   2
            Left            =   1410
            MaskColor       =   &H00000000&
            TabIndex        =   175
            ToolTipText     =   "Buscar Calidad"
            Top             =   3690
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtAux4 
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   5
            Left            =   5790
            MaxLength       =   30
            TabIndex        =   174
            Tag             =   "Nom.Calibrador 3|T|S|||rcalidad_calibrador|nomcalibrador3|||"
            Text            =   "nom.calibrador 2"
            Top             =   3720
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.TextBox txtAux4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   0
            Left            =   135
            MaxLength       =   6
            TabIndex        =   162
            Tag             =   "Código Variedad|N|N|1|999999|rcalidad_calibrador|codvarie|000000|S|"
            Text            =   "codvar"
            Top             =   3690
            Visible         =   0   'False
            Width           =   375
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   4
            Left            =   0
            TabIndex        =   163
            Top             =   0
            Width           =   1200
            _ExtentX        =   2117
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
            Height          =   5985
            Index           =   4
            Left            =   0
            TabIndex        =   164
            Top             =   510
            Width           =   10755
            _ExtentX        =   18971
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
         Height          =   6615
         Left            =   -74865
         TabIndex        =   158
         Top             =   500
         Width           =   11940
         Begin VB.TextBox txtAux3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   0
            Left            =   1485
            MaxLength       =   6
            TabIndex        =   167
            Tag             =   "Código Variedad|N|N|1|999999|variedades_rel|codvarie|000000|S|"
            Text            =   "codvar"
            Top             =   4185
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtAux3 
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   1
            Left            =   2025
            MaxLength       =   10
            TabIndex        =   165
            Tag             =   "Fecha Entrada|F|N|||rbonifentradas|fechaent|dd/mm/yyyy|S|"
            Top             =   4140
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton btnBuscar 
            Appearance      =   0  'Flat
            Caption         =   "+"
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
            Height          =   330
            Index           =   1
            Left            =   2790
            MaskColor       =   &H00000000&
            TabIndex        =   168
            ToolTipText     =   "Buscar fecha"
            Top             =   4140
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   2
            Left            =   3105
            MaxLength       =   6
            TabIndex        =   166
            Tag             =   "Porcentaje Bonificación|N|N|0.00|999.99|rbonifentradas|porcbonif|##0.00||"
            Top             =   4140
            Visible         =   0   'False
            Width           =   720
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   3
            Left            =   0
            TabIndex        =   159
            Top             =   0
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   688
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   6
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
                  Style           =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Alta Masiva"
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Baja Masiva"
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
            Height          =   5985
            Index           =   3
            Left            =   0
            TabIndex        =   160
            Top             =   510
            Width           =   4230
            _ExtentX        =   7461
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
      Begin VB.CheckBox chkComer 
         Caption         =   "Comercializada en común"
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
         Left            =   -74640
         TabIndex        =   94
         Tag             =   "Comercializada en Común|N|N|||variedades|comerciocomun|||"
         Top             =   6150
         Width           =   3600
      End
      Begin VB.Frame FrameAux2 
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   -74880
         TabIndex        =   138
         Top             =   500
         Width           =   11940
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
            Height          =   330
            Index           =   0
            Left            =   1800
            MaskColor       =   &H00000000&
            TabIndex        =   145
            ToolTipText     =   "Buscar Variedad"
            Top             =   3720
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtAux2 
            Appearance      =   0  'Flat
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
            Height          =   330
            Index           =   6
            Left            =   2070
            TabIndex        =   144
            Top             =   3720
            Visible         =   0   'False
            Width           =   1815
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
            Height          =   330
            Index           =   6
            Left            =   1320
            MaxLength       =   12
            TabIndex        =   141
            Tag             =   "Variedad Relacionada|N|N|0|999999|variedades_rel|codvarie1|000000||"
            Text            =   "nomcali"
            Top             =   3675
            Visible         =   0   'False
            Width           =   255
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
            Height          =   330
            Index           =   5
            Left            =   720
            MaxLength       =   2
            TabIndex        =   140
            Tag             =   "Numlinea|N|N|1|99|variedades_rel|numlinea|00|S|"
            Text            =   "li"
            Top             =   3675
            Visible         =   0   'False
            Width           =   315
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
            Height          =   330
            Index           =   4
            Left            =   120
            MaxLength       =   6
            TabIndex        =   139
            Tag             =   "Código Variedad|N|N|1|999999|variedades_rel|codvarie|000000|S|"
            Text            =   "codvar"
            Top             =   3600
            Visible         =   0   'False
            Width           =   375
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   2
            Left            =   0
            TabIndex        =   142
            Top             =   0
            Width           =   1200
            _ExtentX        =   2117
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
            Height          =   5985
            Index           =   2
            Left            =   0
            TabIndex        =   143
            Top             =   510
            Width           =   10125
            _ExtentX        =   17859
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
      Begin VB.TextBox text1 
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
         Index           =   40
         Left            =   14700
         MaxLength       =   10
         TabIndex        =   16
         Tag             =   "Fecha Fin Pixat|F|S|||variedades|fecfinpixat|dd/mm/yyyy||"
         Top             =   2100
         Width           =   1350
      End
      Begin VB.TextBox text1 
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
         Index           =   39
         Left            =   10410
         MaxLength       =   10
         TabIndex        =   15
         Tag             =   "Fecha Inicio Pixat|F|S|||variedades|fecinipixat|dd/mm/yyyy||"
         Text            =   "dd/mm/yyyy"
         Top             =   2085
         Width           =   1350
      End
      Begin VB.TextBox text1 
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
         Index           =   38
         Left            =   15105
         MaxLength       =   6
         TabIndex        =   14
         Tag             =   "Variedad Retirada|T|S|||variedades|codvarret|||"
         Top             =   1470
         Width           =   915
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
         Left            =   10410
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Tag             =   "Tipo Variedad|N|N|||variedades|tipovarie2||N|"
         Top             =   1470
         Width           =   1800
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
         Left            =   10410
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Tag             =   "Tipo Mercancia|N|N|||variedades|tipovariedad||N|"
         Top             =   960
         Width           =   1800
      End
      Begin VB.Frame Frame6 
         Caption         =   "Cuenta Contable Comisionistas"
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
         Height          =   1275
         Left            =   8460
         TabIndex        =   116
         Top             =   5580
         Width           =   8220
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
            Index           =   34
            Left            =   2715
            TabIndex        =   117
            Top             =   360
            Width           =   5250
         End
         Begin VB.TextBox text1 
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
            Index           =   34
            Left            =   1350
            MaxLength       =   10
            TabIndex        =   20
            Tag             =   "Cta Comisionista|T|S|||variedades|ctacomisionista|||"
            Top             =   360
            Width           =   1290
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   16
            Left            =   1080
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   390
            Width           =   240
         End
         Begin VB.Label Label35 
            Caption         =   "Cuenta"
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
            TabIndex        =   118
            Top             =   390
            Width           =   840
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Cuentas Contables"
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
         Height          =   3345
         Left            =   -70290
         TabIndex        =   109
         Top             =   3525
         Width           =   10080
         Begin VB.TextBox text1 
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
            Index           =   37
            Left            =   2475
            MaxLength       =   35
            TabIndex        =   108
            Tag             =   "Cta.Acarreo Recolección|T|S|||variedades|ctaacarecol|||"
            Top             =   2700
            Width           =   1410
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
            Index           =   37
            Left            =   3990
            TabIndex        =   123
            Top             =   2700
            Width           =   5520
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
            Index           =   36
            Left            =   3990
            TabIndex        =   121
            Top             =   2265
            Width           =   5520
         End
         Begin VB.TextBox text1 
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
            Index           =   36
            Left            =   2475
            MaxLength       =   35
            TabIndex        =   107
            Tag             =   "Cta.Facturas Transporte|T|S|||variedades|ctatransporte|||"
            Top             =   2265
            Width           =   1410
         End
         Begin VB.TextBox text1 
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
            Index           =   35
            Left            =   2475
            MaxLength       =   35
            TabIndex        =   106
            Tag             =   "Cta.Compras Terceros|T|S|||variedades|ctasiniestros|||"
            Top             =   1800
            Width           =   1410
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
            Index           =   35
            Left            =   3990
            TabIndex        =   119
            Top             =   1800
            Width           =   5520
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
            Index           =   33
            Left            =   3990
            TabIndex        =   115
            Top             =   1335
            Width           =   5520
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
            Index           =   32
            Left            =   3990
            TabIndex        =   114
            Top             =   870
            Width           =   5520
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
            Index           =   31
            Left            =   3990
            TabIndex        =   113
            Top             =   375
            Width           =   5520
         End
         Begin VB.TextBox text1 
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
            Index           =   33
            Left            =   2475
            MaxLength       =   35
            TabIndex        =   105
            Tag             =   "Cta.Compras Terceros|T|S|||variedades|ctacomtercero|||"
            Top             =   1335
            Width           =   1410
         End
         Begin VB.TextBox text1 
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
            Index           =   32
            Left            =   2475
            MaxLength       =   35
            TabIndex        =   104
            Tag             =   "Cta Liquidación|T|S|||variedades|ctaliquidacion|||"
            Top             =   855
            Width           =   1410
         End
         Begin VB.TextBox text1 
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
            Index           =   31
            Left            =   2475
            MaxLength       =   35
            TabIndex        =   103
            Tag             =   "Cuenta Anticipos|T|S|||variedades|ctaanticipo|||"
            Top             =   375
            Width           =   1410
         End
         Begin VB.Label Label38 
            Caption         =   "Acarreo Recolección"
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
            TabIndex        =   124
            Top             =   2760
            Width           =   1470
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   19
            Left            =   2220
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   2730
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   18
            Left            =   2220
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   2295
            Width           =   240
         End
         Begin VB.Label Label37 
            Caption         =   "Facturas Transporte"
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
            TabIndex        =   122
            Top             =   2295
            Width           =   2100
         End
         Begin VB.Label Label36 
            Caption         =   "Siniestros"
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
            TabIndex        =   120
            Top             =   1830
            Width           =   1470
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   17
            Left            =   2220
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   1830
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   15
            Left            =   2220
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   1365
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   14
            Left            =   2220
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   885
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   13
            Left            =   2220
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   405
            Width           =   240
         End
         Begin VB.Label Label34 
            Caption         =   "Compras Terceros"
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
            TabIndex        =   112
            Top             =   1365
            Width           =   1830
         End
         Begin VB.Label Label33 
            Caption         =   "Liquidación"
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
            TabIndex        =   111
            Top             =   885
            Width           =   1440
         End
         Begin VB.Label Label32 
            Caption         =   "Anticipos"
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
            TabIndex        =   110
            Top             =   405
            Width           =   1440
         End
      End
      Begin VB.TextBox text1 
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
         Index           =   30
         Left            =   -72420
         MaxLength       =   35
         TabIndex        =   93
         Tag             =   "Rdto Maximo|N|S|||variedades|rdtomaximo|###,###,##0||"
         Top             =   5565
         Width           =   1365
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
         Left            =   -72420
         TabIndex        =   83
         Tag             =   "Clasificación|N|N|0|1|variedades|tipoclasifica|0|N|"
         Text            =   "Combo1"
         Top             =   1200
         Width           =   1365
      End
      Begin VB.TextBox text1 
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
         Index           =   20
         Left            =   -63330
         MaxLength       =   35
         TabIndex        =   102
         Tag             =   "Euros/kg hanegada|N|S|||variedades|eurhaneg|0.0000||"
         Top             =   1725
         Width           =   1410
      End
      Begin VB.TextBox text1 
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
         Index           =   19
         Left            =   -63330
         MaxLength       =   35
         TabIndex        =   101
         Tag             =   "Euros/kg tria|N|S|||variedades|eurotria|0.0000||"
         Top             =   1170
         Width           =   1410
      End
      Begin VB.TextBox text1 
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
         Left            =   -67830
         MaxLength       =   35
         TabIndex        =   100
         Tag             =   "Euros/kg Seg.Social|N|S|||variedades|eursegsoc|0.0000||"
         Top             =   2835
         Width           =   1410
      End
      Begin VB.TextBox text1 
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
         Index           =   17
         Left            =   -67830
         MaxLength       =   35
         TabIndex        =   99
         Tag             =   "Euros/kg mano obra|N|S|||variedades|eurmanob|0.0000||"
         Top             =   2280
         Width           =   1410
      End
      Begin VB.TextBox text1 
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
         Index           =   16
         Left            =   -67830
         MaxLength       =   35
         TabIndex        =   98
         Tag             =   "Euros/kg recolecion|N|S|||variedades|eurecole|0.0000||"
         Top             =   1725
         Width           =   1410
      End
      Begin VB.TextBox text1 
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
         Left            =   -67830
         MaxLength       =   35
         TabIndex        =   97
         Tag             =   "Euros/kg destajo|N|S|||variedades|eurdesta|0.0000||"
         Top             =   1170
         Width           =   1410
      End
      Begin VB.TextBox text1 
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
         Left            =   -72420
         MaxLength       =   35
         TabIndex        =   92
         Tag             =   "Porc.Destrio|N|S|||variedades|porcdest|##0.00||"
         Top             =   5010
         Width           =   1365
      End
      Begin VB.TextBox text1 
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
         Left            =   -72420
         MaxLength       =   35
         TabIndex        =   91
         Tag             =   "Porc.Mermas|N|S|||variedades|porcmerm|##0.00||"
         Top             =   4470
         Width           =   1365
      End
      Begin VB.TextBox text1 
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
         Left            =   -72420
         MaxLength       =   35
         TabIndex        =   90
         Tag             =   "Porc.Industria|N|S|||variedades|porcindu|##0.00||"
         Top             =   3930
         Width           =   1365
      End
      Begin VB.TextBox text1 
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
         Left            =   -72420
         MaxLength       =   35
         TabIndex        =   89
         Tag             =   "Arroba/Jornal|N|S|0|999.99|variedades|arrobjor|##0.00||"
         Top             =   3315
         Width           =   1365
      End
      Begin VB.TextBox text1 
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
         Left            =   -63330
         MaxLength       =   35
         TabIndex        =   96
         Tag             =   "Factor Cor.Destrio|N|S|0|999.99|variedades|facorrme|##0.00||"
         Top             =   2835
         Width           =   1410
      End
      Begin VB.TextBox text1 
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
         Left            =   -63330
         MaxLength       =   35
         TabIndex        =   95
         Tag             =   "Factor Cor.Destrio|N|S|0|999.99|variedades|facorrde|##0.00||"
         Top             =   2280
         Width           =   1410
      End
      Begin VB.TextBox text1 
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
         Left            =   -72420
         MaxLength       =   35
         TabIndex        =   88
         Tag             =   "Max Kilos Cajon|N|S|0|999.99|variedades|maxkgcaj|##0.00||"
         Top             =   2805
         Width           =   1365
      End
      Begin VB.TextBox text1 
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
         Left            =   -72420
         MaxLength       =   35
         TabIndex        =   86
         Tag             =   "Min Kilos Cajon|N|S|0|999.99|variedades|minkgcaj|##0.00||"
         Top             =   2265
         Width           =   1365
      End
      Begin VB.TextBox text1 
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
         Left            =   -72420
         MaxLength       =   35
         TabIndex        =   85
         Tag             =   "Kilos Cajon|N|S|0|999.99|variedades|kgscajon|##0.00||"
         Top             =   1770
         Width           =   1365
      End
      Begin VB.Frame FrameAux1 
         BorderStyle     =   0  'None
         Height          =   6585
         Left            =   -74910
         TabIndex        =   65
         Top             =   500
         Width           =   16665
         Begin VB.TextBox txtAux1 
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
            Left            =   180
            MaxLength       =   6
            TabIndex        =   146
            Tag             =   "Código Variedad|N|N|0|999999|rcalidad|codvarie|000000|S|"
            Text            =   "Var"
            Top             =   3420
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.TextBox txtAux1 
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
            Left            =   2745
            MaxLength       =   12
            TabIndex        =   148
            Tag             =   "Nombre Calidad|T|N|||rcalidad|nomcalid|||"
            Text            =   "nomcal"
            Top             =   3420
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.TextBox txtAux1 
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
            Left            =   1215
            MaxLength       =   2
            TabIndex        =   147
            Tag             =   "Codigo Calidad|N|S|1|20|rcalidad|codcalid|00|S|"
            Text            =   "Ca"
            Top             =   3420
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.TextBox txtAux1 
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
            Left            =   3735
            MaxLength       =   3
            TabIndex        =   149
            Tag             =   "Calidad Abrev.|T|N|||rcalidad|nomcalab|||"
            Text            =   "abr"
            Top             =   3420
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmManVariedad.frx":00D0
            Left            =   4725
            List            =   "frmManVariedad.frx":00D2
            Style           =   2  'Dropdown List
            TabIndex        =   150
            Tag             =   "Tipo Calidad|N|N|0|4|rcalidad|tipcalid|||"
            Top             =   3420
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmManVariedad.frx":00D4
            Left            =   6165
            List            =   "frmManVariedad.frx":00D6
            Style           =   2  'Dropdown List
            TabIndex        =   151
            Tag             =   "Tipo Calidad 1|N|N|0|2|rcalidad|tipcalid1|||"
            Top             =   3420
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox txtAux1 
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
            Left            =   7740
            MaxLength       =   30
            TabIndex        =   152
            Tag             =   "Nom.Calibrador 1|T|N|||rcalidad|nomcalibrador1|||"
            Text            =   "nom.calibrador 1"
            Top             =   3420
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.CheckBox chkAux 
            BackColor       =   &H80000005&
            Height          =   255
            Index           =   0
            Left            =   10050
            TabIndex        =   154
            Tag             =   "Hay gastos|N|N|0|1|rcalidad|gastosrec|||"
            Top             =   3420
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtAux1 
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
            Left            =   9060
            MaxLength       =   30
            TabIndex        =   153
            Tag             =   "Nom.Calibrador 2|T|S|||rcalidad|nomcalibrador2|||"
            Text            =   "nom.calibrador 2"
            Top             =   3420
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.TextBox txtAux1 
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
            Left            =   10350
            MaxLength       =   6
            TabIndex        =   155
            Tag             =   "Eu.Rec.Soc|N|S|||rcalidad|eurrecsoc|0.0000||"
            Text            =   "E.Rec.S"
            Top             =   3420
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.TextBox txtAux1 
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
            Index           =   7
            Left            =   11340
            MaxLength       =   6
            TabIndex        =   156
            Tag             =   "Eu.Rec.Coop|N|S|||rcalidad|eurreccoop|0.0000||"
            Text            =   "E.Rec."
            Top             =   3420
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.CheckBox chkAux 
            BackColor       =   &H80000005&
            Height          =   255
            Index           =   1
            Left            =   12270
            TabIndex        =   157
            Tag             =   "Se aplica bonificaciones|N|N|0|1|rcalidad|seaplicabonif|||"
            Top             =   3465
            Visible         =   0   'False
            Width           =   255
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   1
            Left            =   0
            TabIndex        =   66
            Top             =   0
            Width           =   1530
            _ExtentX        =   2699
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
                  Object.ToolTipText     =   "Imprimir"
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc AdoAux 
            Height          =   375
            Index           =   1
            Left            =   5580
            Top             =   90
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
            Height          =   5985
            Index           =   1
            Left            =   0
            TabIndex        =   67
            Top             =   495
            Width           =   16335
            _ExtentX        =   28813
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
      Begin VB.Frame Frame4 
         Caption         =   "Cuentas Contables Transporte"
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
         Height          =   1560
         Left            =   8460
         TabIndex        =   60
         Top             =   3840
         Width           =   8220
         Begin VB.TextBox text1 
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
            Left            =   1350
            MaxLength       =   10
            TabIndex        =   19
            Tag             =   "Cta Transp.Export.|T|S|||variedades|ctatraexporta|||"
            Top             =   990
            Width           =   1290
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
            Index           =   29
            Left            =   2715
            TabIndex        =   62
            Top             =   990
            Width           =   5250
         End
         Begin VB.TextBox text1 
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
            Left            =   1350
            MaxLength       =   10
            TabIndex        =   18
            Tag             =   "Cta Transp.Int.|T|S|||variedades|ctatrainterior|||"
            Text            =   "0000000011"
            Top             =   450
            Width           =   1290
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
            Index           =   28
            Left            =   2715
            TabIndex        =   61
            Top             =   435
            Width           =   5250
         End
         Begin VB.Label Label30 
            Caption         =   "Interior"
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
            TabIndex        =   64
            Top             =   480
            Width           =   840
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   12
            Left            =   1080
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   990
            Width           =   240
         End
         Begin VB.Label Label29 
            Caption         =   "Export."
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
            TabIndex        =   63
            Top             =   975
            Width           =   885
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   11
            Left            =   1080
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   480
            Width           =   240
         End
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
         Index           =   27
         Left            =   11175
         TabIndex        =   58
         Top             =   2670
         Width           =   5310
      End
      Begin VB.TextBox text1 
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
         Index           =   27
         Left            =   10410
         MaxLength       =   4
         TabIndex        =   17
         Tag             =   "Centro Coste|T|S|||variedades|codccost|||"
         Top             =   2670
         Width           =   690
      End
      Begin VB.TextBox text1 
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
         Left            =   1890
         MaxLength       =   2
         TabIndex        =   5
         Tag             =   "Codigo IVA|N|N|0|99|variedades|codigiva|00||"
         Top             =   2505
         Width           =   690
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
         Left            =   2610
         TabIndex        =   56
         Top             =   2520
         Width           =   5520
      End
      Begin VB.Frame Frame3 
         Caption         =   "Cuentas Contables Ventas"
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
         Height          =   3045
         Left            =   270
         TabIndex        =   44
         Top             =   3825
         Width           =   8055
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
            Index           =   25
            Left            =   2655
            TabIndex        =   53
            Top             =   2445
            Width           =   5250
         End
         Begin VB.TextBox text1 
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
            Index           =   25
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   11
            Tag             =   "Cta Vtas Otros|T|S|||variedades|ctavtasotros|||"
            Top             =   2430
            Width           =   1320
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
            Index           =   24
            Left            =   2655
            TabIndex        =   51
            Top             =   1935
            Width           =   5250
         End
         Begin VB.TextBox text1 
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
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   10
            Tag             =   "Cta Vtas Retirada|T|S|||variedades|ctavtasretirada|||"
            Top             =   1935
            Width           =   1320
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
            Index           =   23
            Left            =   2655
            TabIndex        =   49
            Top             =   1470
            Width           =   5250
         End
         Begin VB.TextBox text1 
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
            Index           =   23
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   9
            Tag             =   "Cta Vtas Industria|T|S|||variedades|ctavtasindustria|||"
            Top             =   1455
            Width           =   1320
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
            Index           =   22
            Left            =   2655
            TabIndex        =   47
            Top             =   990
            Width           =   5250
         End
         Begin VB.TextBox text1 
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
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   8
            Tag             =   "Cta Vtas Exportación|T|S|||variedades|ctavtasexportacion|||"
            Top             =   975
            Width           =   1320
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
            Left            =   2655
            TabIndex        =   45
            Top             =   510
            Width           =   5250
         End
         Begin VB.TextBox text1 
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
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "Cta Vtas Interior|T|S|||variedades|ctavtasinterior|||"
            Text            =   "1234567890"
            Top             =   495
            Width           =   1320
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   8
            Left            =   1050
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   2475
            Width           =   240
         End
         Begin VB.Label Label25 
            Caption         =   "Otros"
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
            Left            =   135
            TabIndex        =   54
            Top             =   2460
            Width           =   765
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   7
            Left            =   1050
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   1965
            Width           =   240
         End
         Begin VB.Label Label24 
            Caption         =   "Retirada"
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
            Left            =   135
            TabIndex        =   52
            Top             =   1965
            Width           =   915
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   6
            Left            =   1050
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   1485
            Width           =   240
         End
         Begin VB.Label Label23 
            Caption         =   "Industria"
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
            Left            =   135
            TabIndex        =   50
            Top             =   1485
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   1050
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   1005
            Width           =   240
         End
         Begin VB.Label Label22 
            Caption         =   "Export."
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
            Left            =   135
            TabIndex        =   48
            Top             =   1005
            Width           =   735
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   1050
            ToolTipText     =   "Buscar Cta.Contable"
            Top             =   525
            Width           =   240
         End
         Begin VB.Label Label21 
            Caption         =   "Interior"
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
            Left            =   135
            TabIndex        =   46
            Top             =   525
            Width           =   855
         End
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
         Index           =   9
         Left            =   2610
         TabIndex        =   43
         Top             =   2010
         Width           =   5520
      End
      Begin VB.TextBox text1 
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
         Left            =   1890
         MaxLength       =   3
         TabIndex        =   4
         Tag             =   "Tipo Unidad|N|S|||variedades|codunida|00||"
         Top             =   1995
         Width           =   690
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
         Index           =   3
         Left            =   2610
         TabIndex        =   40
         Top             =   1485
         Width           =   5520
      End
      Begin VB.TextBox text1 
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
         MaxLength       =   3
         TabIndex        =   3
         Tag             =   "Clase|N|N|0|999|variedades|codclase|000||"
         Top             =   1485
         Width           =   675
      End
      Begin VB.Frame FrameAux0 
         BorderStyle     =   0  'None
         Height          =   6525
         Left            =   -74910
         TabIndex        =   37
         Top             =   500
         Width           =   13515
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
            Height          =   330
            Index           =   3
            Left            =   1920
            MaxLength       =   3
            TabIndex        =   26
            Tag             =   "Nombre Calibre Abr|T|N|||calibres|nomcalab|||"
            Text            =   "cab"
            Top             =   3555
            Visible         =   0   'False
            Width           =   2295
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
            Height          =   330
            Index           =   0
            Left            =   -120
            MaxLength       =   6
            TabIndex        =   23
            Tag             =   "Código Variedad|N|N|1|999999|calibres|codvarie|000000|S|"
            Text            =   "codvar"
            Top             =   3555
            Visible         =   0   'False
            Width           =   375
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
            Height          =   330
            Index           =   1
            Left            =   360
            MaxLength       =   2
            TabIndex        =   24
            Tag             =   "Codigo Calibre|N|N|1|99|calibres|codcalib|00|S|"
            Text            =   "ca"
            Top             =   3555
            Visible         =   0   'False
            Width           =   315
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
            Height          =   330
            Index           =   2
            Left            =   720
            MaxLength       =   12
            TabIndex        =   25
            Tag             =   "Nombre Calibre|T|N|||calibres|nomcalib|||"
            Text            =   "nomcali"
            Top             =   3555
            Visible         =   0   'False
            Width           =   1020
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   0
            Left            =   0
            TabIndex        =   38
            Top             =   0
            Width           =   1650
            _ExtentX        =   2910
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
                  Object.ToolTipText     =   "Imprimir"
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
            Height          =   5985
            Index           =   0
            Left            =   0
            TabIndex        =   39
            Top             =   450
            Width           =   5700
            _ExtentX        =   10054
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
      Begin VB.TextBox text1 
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
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   6
         Tag             =   "C.Conselleria|N|N|||variedades|codconse|||"
         Top             =   3075
         Width           =   735
      End
      Begin VB.TextBox text1 
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
         Left            =   1890
         MaxLength       =   3
         TabIndex        =   2
         Tag             =   "Producto|N|N|0|999|variedades|codprodu|000||"
         Top             =   975
         Width           =   675
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
         Index           =   2
         Left            =   2610
         TabIndex        =   27
         Top             =   975
         Width           =   5520
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   14430
         ToolTipText     =   "Buscar fecha"
         Top             =   2130
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   10140
         ToolTipText     =   "Buscar fecha"
         Top             =   2100
         Width           =   240
      End
      Begin VB.Label Label40 
         Caption         =   "F.Fin Pixat"
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
         Left            =   13320
         TabIndex        =   137
         Top             =   2115
         Width           =   1095
      End
      Begin VB.Label Label39 
         Caption         =   "F.Inicio Pixat"
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
         Left            =   8640
         TabIndex        =   136
         Top             =   2100
         Width           =   1455
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   1
         Left            =   16050
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   1470
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Variedad Retirada"
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
         Left            =   13305
         TabIndex        =   127
         Top             =   1500
         Width           =   1920
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Variedad"
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
         Left            =   8625
         TabIndex        =   126
         Top             =   1500
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Mercancia"
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
         Left            =   8625
         TabIndex        =   125
         Top             =   990
         Width           =   1920
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   0
         Left            =   12240
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   990
         Width           =   240
      End
      Begin VB.Label Label31 
         Caption         =   "Rdto Máximo Hda."
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
         Left            =   -74640
         TabIndex        =   87
         Top             =   5595
         Width           =   1920
      End
      Begin VB.Label Label1 
         Caption         =   "Clasificación"
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
         Index           =   19
         Left            =   -74640
         TabIndex        =   84
         Top             =   1245
         Width           =   1350
      End
      Begin VB.Label Label17 
         Caption         =   "Euros/kg Hanegada"
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
         Left            =   -66195
         TabIndex        =   82
         Top             =   1755
         Width           =   2370
      End
      Begin VB.Label Label16 
         Caption         =   "Euros/kg Tria"
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
         Left            =   -66195
         TabIndex        =   81
         Top             =   1200
         Width           =   1770
      End
      Begin VB.Label Label15 
         Caption         =   "Euros/kg Seg.Social"
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
         Left            =   -70155
         TabIndex        =   80
         Top             =   2895
         Width           =   2370
      End
      Begin VB.Label Label14 
         Caption         =   "Euros/kg Mano Obra"
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
         Left            =   -70155
         TabIndex        =   79
         Top             =   2310
         Width           =   2520
      End
      Begin VB.Label Label13 
         Caption         =   "Euros/kg Recolección"
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
         Left            =   -70155
         TabIndex        =   78
         Top             =   1755
         Width           =   2160
      End
      Begin VB.Label Label12 
         Caption         =   "Euros/kg Destajo"
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
         Left            =   -70155
         TabIndex        =   77
         Top             =   1200
         Width           =   2040
      End
      Begin VB.Label Label11 
         Caption         =   "Porcentaje Destrio"
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
         Left            =   -74640
         TabIndex        =   76
         Top             =   5055
         Width           =   2070
      End
      Begin VB.Label Label10 
         Caption         =   "Porcentaje Mermas"
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
         Left            =   -74640
         TabIndex        =   75
         Top             =   4515
         Width           =   1980
      End
      Begin VB.Label Label9 
         Caption         =   "Porcentaje Industria"
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
         Left            =   -74640
         TabIndex        =   74
         Top             =   4005
         Width           =   2250
      End
      Begin VB.Label Label3 
         Caption         =   "Arroba/Jornal"
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
         Left            =   -74640
         TabIndex        =   73
         Top             =   3360
         Width           =   1410
      End
      Begin VB.Label Label2 
         Caption         =   "Factor Corrección Mermas"
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
         Left            =   -66195
         TabIndex        =   72
         Top             =   2865
         Width           =   2715
      End
      Begin VB.Label Label8 
         Caption         =   "Factor Correción Destrio"
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
         Left            =   -66195
         TabIndex        =   71
         Top             =   2310
         Width           =   3330
      End
      Begin VB.Label Label19 
         Caption         =   "Max.Kilos / Cajon"
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
         Left            =   -74640
         TabIndex        =   70
         Top             =   2850
         Width           =   1920
      End
      Begin VB.Label Label6 
         Caption         =   "Min.Kilos / Cajon"
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
         Left            =   -74640
         TabIndex        =   69
         Top             =   2310
         Width           =   1650
      End
      Begin VB.Label Label7 
         Caption         =   "Kilos / Cajon"
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
         Left            =   -74640
         TabIndex        =   68
         Top             =   1770
         Width           =   1515
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   10140
         ToolTipText     =   "Buscar Centro Coste"
         Top             =   2700
         Width           =   240
      End
      Begin VB.Label Label28 
         Caption         =   "Centro Coste"
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
         Left            =   8625
         TabIndex        =   59
         Top             =   2700
         Width           =   1335
      End
      Begin VB.Label Label27 
         Caption         =   "Código IVA"
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
         Left            =   315
         TabIndex        =   57
         Top             =   2535
         Width           =   1125
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   1620
         ToolTipText     =   "Buscar Cta.Contable"
         Top             =   2535
         Width           =   240
      End
      Begin VB.Label Label26 
         Caption         =   "Código EAN"
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
         Left            =   14400
         TabIndex        =   55
         Top             =   990
         Width           =   1185
      End
      Begin VB.Image imgBuscar 
         Height          =   330
         Index           =   4
         Left            =   15630
         ToolTipText     =   "Códigos EAN asociados"
         Top             =   945
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1620
         ToolTipText     =   "Buscar T.Unidad"
         Top             =   2025
         Width           =   240
      End
      Begin VB.Label Label20 
         Caption         =   "Tipo Unidad"
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
         Left            =   315
         TabIndex        =   42
         Top             =   2025
         Width           =   1275
      End
      Begin VB.Label Label5 
         Caption         =   "Clase"
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
         Left            =   315
         TabIndex        =   41
         Top             =   1470
         Width           =   735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1620
         ToolTipText     =   "Buscar Clase"
         Top             =   1530
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1620
         ToolTipText     =   "Buscar Producto"
         Top             =   1005
         Width           =   240
      End
      Begin VB.Label Label18 
         Caption         =   "Producto"
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
         Left            =   315
         TabIndex        =   35
         Top             =   960
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Cód.Consellería"
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
         Left            =   315
         TabIndex        =   34
         Top             =   3105
         Width           =   1710
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4245
      Top             =   6105
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
      Left            =   16035
      TabIndex        =   36
      Top             =   8940
      Visible         =   0   'False
      Width           =   1065
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   16635
      TabIndex        =   135
      Top             =   165
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
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnCopiaCalibres 
         Caption         =   "Copia Calibres/Calidades"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
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
Attribute VB_Name = "frmManVariedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MANOLO                   -+-+
' +-+- Menú: CLIENTES                  -+-+
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

Private Const IdPrograma = 2028


'Dim T1 As Single

Public DatosADevolverBusqueda As String    'Tindrà el nº de text que vol que torne, empipat
Public Event DatoSeleccionado(CadenaSeleccion As String)
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1

Private WithEvents frmIva As frmTipIVAConta
Attribute frmIva.VB_VarHelpID = -1

Private WithEvents frmProd As frmBasico2   'Ayuda Productos de comercial
Attribute frmProd.VB_VarHelpID = -1
Private WithEvents frmCla As frmBasico2   'Clase
Attribute frmCla.VB_VarHelpID = -1
Private WithEvents frmTra As frmComercial  'frmTraCal  'Traer calibres
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmTUn As frmBasico2   'Tipos de unidad
Attribute frmTUn.VB_VarHelpID = -1

Private WithEvents frmCtas As frmCtasConta 'cuentas contables
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmCCos As frmCCosConta 'centros de coste
Attribute frmCCos.VB_VarHelpID = -1
Private WithEvents frmVar As frmBasico2 ' vista previa de variedades
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmVarAux As frmBasico2 ' vista previa de variedades
Attribute frmVarAux.VB_VarHelpID = -1

Private WithEvents frmC1 As frmCal 'calendario fecha
Attribute frmC1.VB_VarHelpID = -1

Private WithEvents frmCal As frmManCalidades 'calidades
Attribute frmCal.VB_VarHelpID = -1

Dim indCodigo As Integer 'indice para txtCodigo


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
Dim Indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim CadB As String

Dim vSeccion As CSeccion

Private BuscaChekc As String



Private Sub cmbAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    
    Select Case Index
        Case 0 'variedades de comercial
            
            Indice = Index + 6
            
            Set frmVarAux = New frmBasico2
            
            AyudaVariedad frmVarAux, , CadB
            
            Set frmVarAux = Nothing
            
            PonerFoco txtAux(Indice)
    
    
        Case 1
           Screen.MousePointer = vbHourglass
           
           Dim esq As Long
           Dim dalt As Long
           Dim menu As Long
           Dim obj As Object
        
           Set frmC1 = New frmCal
            
           esq = btnBuscar(Index).Left
           dalt = btnBuscar(Index).Top
            
           Set obj = btnBuscar(Index).Container
        
           While btnBuscar(Index).Parent.Name <> obj.Name
                esq = esq + obj.Left
                dalt = dalt + obj.Top
                Set obj = obj.Container
           Wend
            
           menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
        
           frmC1.Left = esq + btnBuscar(Index).Parent.Left + 30
           frmC1.Top = dalt + btnBuscar(Index).Parent.Top + btnBuscar(Index).Height + menu - 40
        
           
           frmC1.NovaData = Now
           Select Case Index
                Case 1
                    Indice = Index
           End Select
           
           Me.btnBuscar(1).Tag = Indice
           
           PonerFormatoFecha txtAux3(Indice)
           If txtAux3(Indice).Text <> "" Then frmC1.NovaData = CDate(txtAux3(Indice).Text)
        
           Screen.MousePointer = vbDefault
           frmC1.Show vbModal
           Set frmC1 = Nothing
           PonerFoco txtAux3(Indice)
        
        
        Case 2
            Indice = Index
            Set frmCal = New frmManCalidades
            If txtAux(0).Text <> "" Then frmCal.ParamVariedad = txtAux(0).Text
            frmCal.DatosADevolverBusqueda = "0|2|3|"
            frmCal.CodigoActual = txtAux(Indice).Text
            frmCal.Show vbModal
            Set frmCal = Nothing
            PonerFoco txtAux(Indice)
        
    
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.Adoaux(2), 1

End Sub

Private Sub chkComer_GotFocus(Index As Integer)
    PonerFocoChk Me.chkComer(Index)
End Sub

Private Sub chkComer_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkComer(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "chkComer(" & Index & ")|"
    End If
End Sub

Private Sub chkComer_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub cmdAceptar_Click()
Dim Produ As Long
Dim vCadena As String

    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BÚSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm2(Me, 1) Then
                    
                    CargarUnaVariedad CLng(Text1(0).Text), "I"
                    
                    Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
                    TerminaBloquear
                    PosicionarData
                End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario2(Me, 1) Then
                    '[Monica]18/09/2013: si estamos actualizando variedad en Picassent el claveant en 'PP&VVVV'
                    Produ = DevuelveValor("select codprodu from variedades where codvarie = " & DBSet(Text1(0).Text, "N"))
                    vCadena = CLng(Produ) & "&" & CLng(Text1(0).Text)
                    
                    CargarUnaVariedad CLng(Text1(0).Text), "U", vCadena
                    
                    TerminaBloquear
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
                    If ModificarLinea Then
                        PosicionarData
                    Else
                        PonerFoco txtAux(12)
                    End If
            End Select
        ' **************************
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

' *** si n'hi han combos a la capçalera ***
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
'    If PrimeraVez Then PrimeraVez = False
    If PrimeraVez Then
        PrimeraVez = False
        If DatosADevolverBusqueda = "" Then
            PonerModo 0
        Else
            If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                BotonAnyadir
            Else
                PonerModo 1 'búsqueda
                ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
                Text1(0).BackColor = vbLightBlue 'codvarie
                ' ****************************************************************************
            End If
        End If
    End If
    
'[Monica]21/10/2013: quito esto con respecto a comercial
'    SSTab1.TabEnabled(2) = ExisteTabla("rcalidad")
'    SSTab1.TabVisible(2) = ExisteTabla("rcalidad")
'    SSTab1.TabEnabled(3) = ExisteTabla("rcalidad")
'    SSTab1.TabVisible(3) = ExisteTabla("rcalidad")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
    If Modo = 4 Then TerminaBloquear
    Screen.MousePointer = vbDefault
    
    If Not vSeccion Is Nothing Then
        vSeccion.CerrarConta
        Set vSeccion = Nothing
    End If
    
    
End Sub

Private Sub Form_Load()
Dim i As Integer
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    
    ' ICONETS DE LA BARRA
    btnPrimero = 16 'index del botó "primero"
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
        'el 10 i el 11 son separadors
        .Buttons(8).Image = 10  'Imprimir
    End With
    
    
    With Me.Toolbar2
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 22  'Copiar Calidades y Calibres
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
            If i <= 1 Then
                .Buttons(4).Image = 10 ' Imprimir
            End If
            
            If i = 3 Then
                .Buttons(5).Image = 33 ' alta masiva
                .Buttons(6).Image = 20 ' baja masiva
            End If
        End With
    Next i
    ' ***********************************
    
    'cargar IMAGES de busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    Me.imgBuscar(4).Picture = frmPpal.imgListComun.ListImages(21).Picture
    
    For i = 0 To imgAyuda.Count - 1
        imgAyuda(i).Picture = frmPpal.ImageListB.ListImages(10).Picture
    Next i

    CargaCombo
    
    Set vSeccion = New CSeccion
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If vSeccion.AbrirConta Then
        End If
    End If
    
    ' *** si n'hi han tabs, per a que per defecte sempre es pose al 1r***
    Me.SSTab1.Tab = 0
    ' *******************************************************************
    
    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han llínies *******
    DataGridAux(0).ClearFields
'    DataGridAux(1).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "variedades"
    Ordenacion = " ORDER BY codvarie"
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    Data1.ConnectionString = conn
    '***** cambiar el nombre de la PK de la cabecera *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where codvarie=-1"
    Data1.Refresh
    
    CargaGrid 0, False
    'If ExisteTabla("rcalidad") Then CargaGrid 1, False
    CargaGrid 1, False
    
    '[Monica]21/08/2017: nueva solapa
    CargaGrid 2, False
    
    '[Monica]23/07/2018: bonificaciones y calibradores
    CargaGrid 3, False
    CargaGrid 4, False
    
    
    ModoLineas = 0
       
    ' Para el chivato
    Set dbAriagro = New BaseDatos
    dbAriagro.abrir_MYSQL vConfig.SERVER, vUsu.CadenaConexion, vConfig.User, vConfig.password
    
    
'    If DatosADevolverBusqueda = "" Then
'        PonerModo 0
'    Else
'        PonerModo 1 'búsqueda
'        ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
'        text1(0).BackColor = vbLightBlue 'codclien
'        ' ****************************************************************************
'    End If

    '[Monica]01/03/2016: Si es Abn ponemos los gastos molturacion y envasado en la variedad
    If vParamAplic.Cooperativa = 1 Then
        Label14.Caption = "Euros/kg Gtos.Molturación"
        Text1(17).Tag = "Euros/kg Gtos.Molturación|N|S|||variedades|eurmanob|0.0000||"
        Label15.Caption = "Euros/litro Gtos.Envasado"
        Text1(18).Tag = "Euros/kg Gtos.Envasado|N|S|||variedades|eursegsoc|0.0000||"
        Label12.Caption = "Precio Venta"
        Text1(15).Tag = "Precio Venta|N|S|||variedades|eurdesta|0.0000||"
        Label13.Caption = "Precio Excedido"
        Text1(16).Tag = "Precio Excedido|N|S|||variedades|eurecole|0.0000||"
    End If


End Sub

Private Sub LimpiarCampos()
    On Error Resume Next
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    Me.Combo1(0).ListIndex = -1
    Me.Combo1(1).ListIndex = -1
    Me.Combo1(2).ListIndex = -1
    lblIndicador.Caption = ""
    Me.chkComer(0).Value = 0
    
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
Dim B As Boolean

    On Error GoTo EPonerModo
 
    'Actualisa Iconos Insertar,Modificar,Eliminar
    'ActualizarToolbar Modo, Kmodo
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo, ModoLineas
       
    'Modo 2. N'hi han datos i estam visualisant-los
    '=========================================
    'Posem visible, si es formulari de búsqueda, el botó "Regresar" quan n'hi han datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    
    '=======================================
    B = (Modo = 2)
    'Posar Fleches de desplasament visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Només es per a saber que n'hi ha + d'1 registre
    End If
    DesplazamientoVisible B And Data1.Recordset.RecordCount > 1
    
    '---------------------------------------------
    B = Modo <> 0 And Modo <> 2
    CmdCancelar.visible = B
    cmdAceptar.visible = B
       
    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a més neteja els camps Text1
    BloquearText1 Me, Modo
    
    '++monica: si el modo es insertar damos el siguiente pero dejamos modificar
    If Modo = 3 Then Text1(0).Locked = False
    
    ' ***** bloquejar tots els controls visibles de la clau primaria de la capçalera ***
'    If Modo = 4 Then _
'        BloquearTxt Text1(0), True 'si estic en  modificar, bloqueja la clau primaria
    ' **********************************************************************************
    
    ' **** si n'hi han imagens de buscar en la capçalera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    ' ********************************************************
    For i = 0 To imgFec.Count - 1
        BloquearImgFec Me, i, Modo
    Next i
    
    
    Me.imgBuscar(4).Enabled = (Modo = 2)
    Me.imgBuscar(4).visible = (Modo = 2)
    Me.Label26.visible = (Modo = 2)
    
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    BloquearChk Me.chkComer(0), (Modo = 0 Or Modo = 2 Or Modo = 5 Or ((Modo = 3 Or Modo = 4) And vUsu.Login <> "root"))
    
    BloquearCombo Me, Modo
    PonerLongCampos

    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 0, False
        '[Monica]21/10/2013: quito esto cuando
        'If ExisteTabla("rcalidad") Then
        CargaGrid 1, False
        
        '[Monica]24/08/2017: nueva solapa
        CargaGrid 2, False
        
        '[Monica]23/07/2018: bonificaciones y calibradores
        CargaGrid 3, False
        CargaGrid 4, False
        
    End If
    
    B = (Modo = 4) Or (Modo = 2)
    DataGridAux(0).Enabled = B
    DataGridAux(1).Enabled = B
    
    '[Monica]23/07/2018: bonificaciones y calibradores
    DataGridAux(2).Enabled = B
    DataGridAux(3).Enabled = B
    DataGridAux(4).Enabled = B
      
    ' calidades
    B = (Modo = 5) And (NumTabMto = 1) 'And (ModoLineas <> 3)
    For i = 1 To txtAux1.Count - 1
        BloquearTxt txtAux1(i), Not B
    Next i
    BloquearCmb Me.Combo1(3), Not B
    BloquearCmb Me.Combo1(4), Not B
    BloquearChk Me.chkAux(0), Not B
    BloquearChk Me.chkAux(1), Not B
    
    ' bonificaciones
    B = (Modo = 5) And (NumTabMto = 3)
    For i = 1 To txtAux3.Count - 1
        BloquearTxt txtAux3(i), Not B
    Next i
    B = (Modo = 5) And (NumTabMto = 3) And ModoLineas = 1
    BloquearTxt txtAux3(1), Not B
    BloquearBtn btnBuscar(1), Not B
    
    ' calibradores
    B = (Modo = 5) And (NumTabMto = 4)
    For i = 1 To txtAux4.Count - 1
        BloquearTxt txtAux4(i), Not B
    Next i
    B = (Modo = 5) And (NumTabMto = 4) And ModoLineas = 1
    BloquearTxt txtAux4(1), Not B
    BloquearTxt txtAux4(2), Not B
    
    BloquearBtn btnBuscar(2), Not B
    
    
'    B = (Modo = 5) And (NumTabMto = 1) And ModoLineas = 2
'    BloquearTxt txtAux1(1), B
'    BloquearBtn cmdAux(4), B
      
      
      
    PonerModoOpcionesMenu (Modo) 'Activar opcions menú según modo
    PonerOpcionesMenu   'Activar opcions de menú según nivell
                        'de permisos de l'usuari

EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub frmC1_Selec(vFecha As Date)
Dim Indice As Byte
    Indice = CByte(Me.btnBuscar(1).Tag)
    txtAux3(Indice).Text = Format(vFecha, "dd/mm/yyyy")
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
    PonerOpcionesMenuGeneralNew Me
End Sub

Private Sub PonerModoOpcionesMenu(Modo)
'Actives unes Opcions de Menú i Toolbar según el modo en que estem
Dim B As Boolean, bAux As Boolean
Dim i As Byte
    
    'Barra de CAPÇALERA
    '------------------------------------------
    'b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    B = (Modo = 2 Or Modo = 0)
    'Buscar
    Toolbar1.Buttons(5).Enabled = B
    Me.mnBuscar.Enabled = B
    'Vore Tots
    Toolbar1.Buttons(6).Enabled = B
    Me.mnVerTodos.Enabled = B
    
    'Insertar
    Toolbar1.Buttons(1).Enabled = B And Not DeConsulta
    Me.mnNuevo.Enabled = B And Not DeConsulta
    
    B = (Modo = 2 And Data1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    Me.mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(3).Enabled = B
    Me.mnEliminar.Enabled = B
    'copiar calibres y calidades
    Toolbar2.Buttons(1).Enabled = B
    Me.mnCopiaCalibres.Enabled = B
    
    'Imprimir
    'Toolbar1.Buttons(12).Enabled = (b Or Modo = 0)
    Toolbar1.Buttons(8).Enabled = True And Not DeConsulta
       
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
    B = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    For i = 0 To 4
        ToolAux(i).Buttons(1).Enabled = B
        If B Then bAux = (B And Me.Adoaux(i).Recordset.RecordCount > 0)
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
        
        If i = 3 Then
            ToolAux(i).Buttons(5).Enabled = True
            ToolAux(i).Buttons(6).Enabled = True
        End If
    Next i
    
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botons de Desplaçament; per a desplaçar-se pels registres de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index, True
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
Dim SQL As String
Dim Tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
               
        Case 0 'CALIBRES
            SQL = "SELECT codvarie, codcalib, nomcalib, nomcalab"
            SQL = SQL & " FROM calibres "
            If enlaza Then
                SQL = SQL & ObtenerWhereCab(True)
            Else
                SQL = SQL & " WHERE calibres.codvarie = -1"
            End If
            SQL = SQL & " ORDER BY calibres.codcalib"
               
        Case 1 'CALIDADES
            SQL = "SELECT rcalidad.codvarie,codcalid,nomcalid, nomcalab, tipcalid, CASE rcalidad.tipcalid WHEN 0 THEN ""Normal"" WHEN 1 THEN ""Destrio"" WHEN 2 THEN ""Venta Campo"" WHEN 3 THEN ""Mermas"" WHEN 4 THEN ""Pequeño"" WHEN 5 THEN ""Pixat"" WHEN 6 THEN ""Destr.Comer."" END,   "
            SQL = SQL & "rcalidad.tipcalid1, "
            SQL = SQL & " CASE rcalidad.tipcalid1 WHEN 0 THEN ""Comercial"" WHEN 1 THEN ""No Comercial"" WHEN 2 THEN ""Retirada"" END, "
            SQL = SQL & " nomcalibrador1, nomcalibrador2, gastosrec, IF(gastosrec=1,'*','') as dgastorec, "
            SQL = SQL & " rcalidad.eurrecsoc, rcalidad.eurreccoop, "
            SQL = SQL & " seaplicabonif, IF(seaplicabonif=1,'*','') as dseaplicabonif"
            SQL = SQL & " FROM rcalidad"
            If enlaza Then
                SQL = SQL & ObtenerWhereCab(True)
            Else
                SQL = SQL & " where codvarie is null "
            End If
            
            SQL = SQL & " ORDER BY codcalid"
            
            
        Case 2 ' variedades relacionadas
            SQL = "SELECT variedades_rel.codvarie, variedades_rel.numlinea, variedades_rel.codvarie1, dd.nomvarie"
            SQL = SQL & " FROM variedades_rel inner join variedades dd on variedades_rel.codvarie1 = dd.codvarie "
            If enlaza Then
                SQL = SQL & " WHERE variedades_rel.codvarie=" & Val(Text1(0).Text)
            Else
                SQL = SQL & " WHERE variedades_rel.codvarie is null "
            End If
            SQL = SQL & " ORDER BY variedades_rel.numlinea"
            
            
        Case 3 ' bonificaciones
            SQL = "SELECT rbonifentradas.codvarie,  "
            SQL = SQL & "rbonifentradas.fechaent, rbonifentradas.porcbonif"
            SQL = SQL & " FROM rbonifentradas "
            If enlaza Then
                SQL = SQL & ObtenerWhereCab(True)
            Else
                SQL = SQL & " WHERE rbonifentradas.codvarie is null "
            End If
            
        Case 4 ' calibradores
            SQL = "SELECT rcalidad_calibrador.codvarie, rcalidad_calibrador.codcalid, rcalidad.nomcalid, "
            SQL = SQL & " rcalidad_calibrador.numlinea, rcalidad_calibrador.nomcalibrador1, "
            SQL = SQL & " rcalidad_calibrador.nomcalibrador2, rcalidad_calibrador.nomcalibrador3 "
            SQL = SQL & " FROM rcalidad, rcalidad_calibrador "
            SQL = SQL & " WHERE rcalidad.codvarie = rcalidad_calibrador.codvarie and "
            SQL = SQL & " rcalidad.codcalid = rcalidad_calibrador.codcalid "
            If enlaza Then
                SQL = SQL & " and rcalidad_calibrador.codvarie=" & Val(Text1(0).Text)
            Else
                SQL = SQL & " and rcalidad_calibrador.codvarie is null "
            End If
        
    End Select
    
    MontaSQLCarga = SQL
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
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub


Private Sub frmC_Selec(vFecha As Date)
Dim Indice As Byte
    Indice = CByte(Me.imgFec(0).Tag)
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCCos_DatoSeleccionado(CadenaSeleccion As String)
'Centro de Coste de la contabilidad
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta
End Sub

Private Sub frmIva_DatoSeleccionado(CadenaSeleccion As String)
'Tipo de iva de la contabilidad
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigiva
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Porceiva
End Sub

Private Sub frmTra_Actualizar(vValor As String)
    On Error GoTo EEPonerBusq
    
    LimpiarCampos
    Text1(0).Text = vValor 'codvarie
    
    FormateaCampo Text1(0)
    
    Screen.MousePointer = vbHourglass
    
    If vValor = "" Then vValor = " codvarie = -1"
    Data1.RecordSource = "select * from variedades where " & vValor
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
    
'        Modo = 1
'        cmdAceptar_Click
End Sub

Private Sub frmProd_DatoSeleccionado(CadenaSeleccion As String)
    Text1(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    Text2(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTUn_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de tipos de unidad
    Text1(9).Text = RecuperaValor(CadenaSeleccion, 1) 'tipos de unidad
    FormateaCampo Text1(9)
    Text2(9).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre de unidad
End Sub


Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
    CadB = "codvarie = " & RecuperaValor(CadenaSeleccion, 1)
    
    'Se muestran en el mismo form
    CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
    PonerCadenaBusqueda
    Screen.MousePointer = vbDefault

End Sub


Private Sub frmVarAux_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        txtAux(6).Text = RecuperaValor(CadenaSeleccion, 1)
        txtAux2(6).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub


Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "Indica si la variedad es mercancia que es suministrada por el cliente, o si por el contrario es mercancia de la cooperativa." & vbCrLf & vbCrLf & _
                      "Se utiliza para restringir las variedades a mostrar en los informes. " & vbCrLf & _
                      vbCrLf
                      
        Case 1
           ' "____________________________________________________________"
            vCadena = "Indica la variedad que viene en el fichero de traspaso de albaranes de retirada de almazara." & vbCrLf & vbCrLf & _
                      "En los albaranes de retirada se grabará el código de variedad asociada a la variedad de retirada." & vbCrLf & _
                      vbCrLf
                      
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
    
End Sub

Private Sub imgFec_Click(Index As Integer)
       
       Screen.MousePointer = vbHourglass
       
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
    
       
       frmC.NovaData = Now
       Select Case Index
            Case 0, 1
                Indice = Index + 39
       End Select
       
       Me.imgFec(0).Tag = Indice
       
       PonerFormatoFecha Text1(Indice)
       If Text1(Indice).Text <> "" Then frmC.NovaData = CDate(Text1(Indice).Text)
    
       Screen.MousePointer = vbDefault
       frmC.Show vbModal
       Set frmC = Nothing
       PonerFoco Text1(Indice)
      
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnBuscarCalibre_Click()
'    Set frmTra = New frmTraerCalib
'    frmTra.DatosADevolverBusqueda = "0|1|"
'    frmTra.CodigoActual = text1(0).Text
'    frmTra.Show vbModal
'    Set frmTra = Nothing
'    PonerFoco text1(0)
End Sub

Private Sub mnCopiaCalibres_Click()

    If Data1.Recordset.EOF Then Exit Sub
    
    frmCopiaCalibCalid.NumCod = Data1.Recordset!codvarie
    frmCopiaCalibCalid.Show vbModal

End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    AbrirListado (122)
End Sub

Private Sub mnModificar_Click()
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
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

Private Sub SSTab1_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
        PonerFocoGrid DataGridAux(0)
    End If
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
        Case 8 'Imprimir
            mnImprimir_Click
            
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
        Text1(0).BackColor = vbLightBlue ' <===
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
        ' *** foco al 1r camp visible de la capçalera que siga clau primaria ***
        PonerFoco Text1(0)
        ' **********************************************************************
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
'    Dim Cad As String
'
'    'Cridem al form
'    ' **************** arreglar-ho per a vore lo que es desije ****************
'    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
'    Cad = ""
'    Cad = Cad & ParaGrid(Text1(0), 15, "Cód.")
'    Cad = Cad & ParaGrid(Text1(1), 60, "Nombre")
'    Cad = Cad & ParaGrid(Text1(2), 25, "Producto")
'    If Cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = Cad
'        frmB.vtabla = NombreTabla
'        frmB.vSQL = CadB
'        HaDevueltoDatos = False
'        frmB.vDevuelve = "0|1|2|" '*** els camps que volen que torne ***
'        frmB.vTitulo = "Variedades" ' ***** repasa açò: títol de BuscaGrid *****
'        frmB.vSelElem = 1
'
'        frmB.Show vbModal
'        Set frmB = Nothing
'        'Si ha posat valors i tenim que es formulari de búsqueda llavors
'        'tindrem que tancar el form llançant l'event
'        If HaDevueltoDatos Then
'            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                cmdRegresar_Click
'        Else   'de ha retornat datos, es a decir NO ha retornat datos
'            PonerFoco Text1(kCampo)
'        End If
'    End If


    Set frmVar = New frmBasico2
    
    AyudaVariedadPrevio frmVar, , CadB
    
    Set frmVar = Nothing


End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim Aux As String
Dim i As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    cad = ""
    i = 0
    Do
        J = i + 1
        i = InStr(J, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, i - J)
            J = Val(Aux)
            cad = cad & Text1(J).Text & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(cad)
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
    
    If chkVistaPrevia(0).Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String

    LimpiarCampos 'Huida els TextBox
    PonerModo 3
    
    '******************** canviar taula i camp **************************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = SugerirCodigoSiguienteStr("variedades", "codvarie")
    End If
    '********************************************************************
    
    
    ' codEmpre i quins camps tenen la PK de la capçalera *******
    Text1(0).Text = NumF
    FormateaCampo Text1(0)
       
    PonerFoco Text1(0) '*** 1r camp visible que siga PK ***
    
    ' *** si n'hi han camps de descripció a la capçalera ***
    'PosarDescripcions

    ' *** si n'hi han tabs, em posicione al 1r ***
    Me.SSTab1.Tab = 0
End Sub

Private Sub BotonModificar()

    PonerModo 4

    ' *** bloquejar els camps visibles de la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(1)
End Sub

Private Sub BotonEliminar()
Dim cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    ' ***************************************************************************

    ' *************** canviar la pregunta ****************
    cad = "¿Seguro que desea eliminar la Variedad?"
    cad = cad & vbCrLf & "Código: " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
    cad = cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(2)
    
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Variedad", Err.Description
End Sub

Private Sub PonerCampos()
Dim i As Integer
Dim CodPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la capçalera
    
    ' *** si n'hi han llínies en datagrids ***
    'For i = 0 To DataGridAux.Count - 1
    '[Monica]23/07/2018: bonificaciones y calibradores
    For i = 0 To 4 ' antes 2
            '[Monica]21/10/2013: quito esta condicion
            'If Not ExisteTabla("rcalidad") Then Exit For
            
            CargaGrid i, True
            If Not Adoaux(i).Recordset.EOF Then _
                PonerCamposForma2 Me, Adoaux(i), 2, "FrameAux" & i
    Next i

    If vParamAplic.NumeroConta <> 0 Then
        Text2(26).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(26), "N")
    
    ' ************* configurar els camps de les descripcions de la capçalera *************
'    Text2(21).Text = NombreCuentaCorrecta(Text1(21).Text)
'    Text2(22).Text = NombreCuentaCorrecta(Text1(22).Text)
'    Text2(23).Text = NombreCuentaCorrecta(Text1(23).Text)
'    Text2(24).Text = NombreCuentaCorrecta(Text1(24).Text)
'    Text2(25).Text = NombreCuentaCorrecta(Text1(25).Text)
        Text2(21).Text = PonerNombreCuenta(Text1(21), Modo)
        Text2(22).Text = PonerNombreCuenta(Text1(22), Modo)
        Text2(23).Text = PonerNombreCuenta(Text1(23), Modo)
        Text2(24).Text = PonerNombreCuenta(Text1(24), Modo)
        Text2(25).Text = PonerNombreCuenta(Text1(25), Modo)
        Text2(27).Text = PonerNombreDeCod(Text1(27), "cabccost", "nomccost", "codccost", "T", cConta)
    
        Text2(28).Text = PonerNombreCuenta(Text1(28), Modo)
        Text2(29).Text = PonerNombreCuenta(Text1(29), Modo)
    
        Text2(34).Text = PonerNombreCuenta(Text1(34), Modo)
        
        '[Monica]18/07/2012: las ejecuto solo en el caso de que coindicida la contabilidad de horto con la de comercial
        '                    las vuelvo a poner
            '[Monica]23/03/2012: las cuentas de recoleccion puede que no sean de la conta de parametros
            '                    seran de la seccion que corresponda de recoleccion
            '                    quito las 6 instrucciones siguientes
        If vParamAplic.NumeroConta = DevuelveValor("select empresa_conta from rseccion, rparam where rseccion.codsecci = rparam.seccionhorto") Then
            Text2(31).Text = PonerNombreCuenta(Text1(31), Modo)
            Text2(32).Text = PonerNombreCuenta(Text1(32), Modo)
            Text2(33).Text = PonerNombreCuenta(Text1(33), Modo)
            Text2(35).Text = PonerNombreCuenta(Text1(35), Modo)
            Text2(36).Text = PonerNombreCuenta(Text1(36), Modo)
            Text2(37).Text = PonerNombreCuenta(Text1(37), Modo)
        End If
    Else
        Text2(26).Text = DevuelveDesdeBDNew(cAgro, "tiposiva", "nombriva", "nombriva", Text1(26), "N")
    End If
    Text2(2).Text = PonerNombreDeCod(Text1(2), "productos", "nomprodu", "codprodu", "N")
    Text2(3).Text = PonerNombreDeCod(Text1(3), "clases", "nomclase", "codclase", "N")
    Text2(9).Text = PonerNombreDeCod(Text1(9), "sunida", "nomunida", "codunida", "N")
    ' ********************************************************************************
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu
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

        Case 4  'Modificar
                TerminaBloquear
                PonerModo 2
                PonerCampos
                ' *** primer camp visible de la capçalera ***
                PonerFoco Text1(0)
        
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                    ModoLineas = 0
                    ' *** les llínies que tenen datagrid (en o sense tab) ***
                    If NumTabMto = 0 Or NumTabMto = 1 Or NumTabMto = 2 Or NumTabMto = 3 Or NumTabMto = 4 Then
                        DataGridAux(NumTabMto).AllowAddNew = False
                        ' **** repasar si es diu Data1 l'adodc de la capçalera ***
                        'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
                        LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                        DataGridAux(NumTabMto).Enabled = True
                        DataGridAux(NumTabMto).SetFocus

                        ' *** si n'hi han camps de descripció dins del grid, els neteje ***
                        'txtAux2(2).text = ""

                    End If
                    
                    ' *** si n'hi han tabs ***
                    SituarTab (NumTabMto + 1)
                    
                    If Not Adoaux(NumTabMto).Recordset.EOF Then
                        Adoaux(NumTabMto).Recordset.MoveFirst
                    End If

                Case 2 'modificar llínies
                    ModoLineas = 0
                    
                    ' *** si n'hi han tabs ***
                    SituarTab (NumTabMto + 1)
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                    PonerModo 4
                    If Not Adoaux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de llínies ***
                        V = Adoaux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
                        Adoaux(NumTabMto).Recordset.Find (Adoaux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
                        ' ***************************************************************
                    End If

            End Select
            
            PosicionarData
            
            ' *** si n'hi han llínies en grids i camps fora d'estos ***
            If Not Adoaux(NumTabMto).Recordset.EOF Then
                DataGridAux_RowColChange NumTabMto, 1, 1
            Else
                LimpiarCamposFrame NumTabMto
            End If
    End Select
End Sub

Private Function DatosOk() As Boolean
Dim B As Boolean
Dim SQL As String
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOk = False
    B = CompForm2(Me, 1)
    If Not B Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
        'comprobar si existe ya el cod. del campo clave primaria
        If ExisteCP(Text1(0)) Then B = False
    End If
    
    If B Then
        If vEmpresa.TieneAnalitica Then 'hay contab. analitica
             SQL = DevuelveDesdeBDNew(cConta, "cabccost", "codccost", "codccost", Text1(27), "T")
             If SQL = "" Then
                MsgBox "No existe el Centro de Coste. Reintroduzca.", vbExclamation
                PonerFoco Text1(27)
                B = False
             End If
        End If
    End If

    ' ************************************************************************************
    
    DatosOk = B
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub PosicionarData()
Dim cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    cad = "(codvarie=" & Text1(0).Text & ")"
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarData(Data1, cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub

Private Function Eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar

    conn.BeginTrans
    ' ***** canviar el nom de la PK de la capçalera, repasar codEmpre *******
    vWhere = " WHERE codvarie=" & Data1.Recordset!codvarie
        
    ' ***** elimina les llínies ****
    conn.Execute "DELETE FROM variane " & vWhere
        
    conn.Execute "DELETE FROM calibres " & vWhere
    
'[Monica]21/10/2013: quito estas condicion
'    If ExisteTabla("rcalidad") Then
        conn.Execute "DELETE FROM rcalidad " & vWhere
'    End If
        
    CargarUnaVariedad CLng(Data1.Recordset!codvarie), "D"
        
    'Eliminar la CAPÇALERA
    vWhere = " WHERE codvarie=" & Data1.Recordset!codvarie
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
Dim NumDigit As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    ' ***************** configurar els LostFocus dels camps de la capçalera *****************
    Select Case Index
        Case 0 'cod variedad
            PonerFormatoEntero Text1(0)

        Case 1 'NOMBRE
            Text1(Index).Text = UCase(Text1(Index).Text)
        
        Case 2 'Producto
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "productos", "nomprodu")
                If Text2(Index).Text = "" Then
                    MsgBox "No existe el Producto. Reintroduzca.", vbExclamation
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 3 'clase
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "clases", "nomclase", "codclase", "N")
                If Text2(Index).Text = "" Then
                    MsgBox "No existe la Clase. Reintroduzca.", vbExclamation
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
                
        Case 9 'Tipo de Unidad
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "sunida", "nomunida")
                If Text2(Index).Text = "" Then
                    MsgBox "No existe el Tipo de Unidad. Reintroduzca.", vbExclamation
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 5, 6, 7, 8, 10, 11, 12, 13, 14
            If Modo = 1 Then Exit Sub
            cadMen = TransformaPuntosComas(Text1(Index).Text)
            Text1(Index).Text = Format(cadMen, "##0.00")
            
        Case 15, 16, 17, 18, 19, 20
            If Modo = 1 Then Exit Sub
            cadMen = TransformaPuntosComas(Text1(Index).Text)
            Text1(Index).Text = Format(cadMen, "0.0000")
        
        Case 21, 22, 23, 24, 25, 28, 29, 31, 32, 33, 34, 35, 36, 37 'cta contable de ventas
            If Text1(Index).Text = "" Then
                Text2(Index) = ""
                Exit Sub
            End If
            
'            If Modo <> 1 Then
'                NumDigit = DevuelveDesdeBDNew(cConta, "empresa", "numdigi3", "codempre", vParamAplic.NumeroConta, "N")
'                If Len(Text1(21).Text) <> CCur(NumDigit) Then
'                    MsgBox "La longitud de la cuenta no se corresponde con el nivel 3.", vbExclamation
'                End If
'            End If
            
'            Text2(Index).Text = NombreCuentaCorrecta(Text1(Index).Text)
            Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo)
            
            If Index = 37 Then cmdAceptar.SetFocus
    
    
        Case 26 ' tipo de iva de contabilidad
            'conConta: BD Contabilidad
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "tiposiva", "nombriva", , , cConta)
            Else
                Text2(Index).Text = ""
            End If
            
        Case 27 ' centro de coste
            If Text1(Index).Text <> "" Then
                If vParamAplic.NumeroConta <> 0 Then
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), "cabccost", "nomccost", "codccost", "T", cConta)
                End If
            End If
        
        Case 30
            If Text1(Index).Text <> "" Then PonerFormatoEntero Text1(Index)
    
        '[Monica]25/01/2017: se pone inicio y fin del pixat
        Case 39, 40
            PonerFormatoFecha Text1(Index), False
            
    End Select
        ' ***************************************************************************
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 2: KEYBusqueda KeyAscii, 0 'producto
                Case 3: KEYBusqueda KeyAscii, 1 'clase
                
                Case 39: KEYFecha KeyAscii, 0 ' fechainicio de pixat
                Case 40: KEYFecha KeyAscii, 1 ' fechafin de pixat
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
            BotonImprimir Index
        Case 5 ' alta masiva en bonificaciones
            mnAltaMasiva_Click
        Case 6 ' baja masiva en bonificaciones
            mnBajaMasiva_Click
        Case Else
    End Select
    'End If
End Sub

Private Sub BotonImprimir(Index As Integer)
    Select Case Index
        Case 0
            AbrirListado (121)
        Case 1
            AbrirListado (12)
    End Select

End Sub

Private Sub mnAltaMasiva_Click()
    BotonAltaMasiva
End Sub

Private Sub mnBajaMasiva_Click()
    BotonBajaMasiva
End Sub

Private Sub BotonAltaMasiva()
    AbrirListado (28)
    CargaGrid 3, True
End Sub

Private Sub BotonBajaMasiva()
    AbrirListado (29)
    CargaGrid 3, True
End Sub

Private Sub BotonEliminarLinea(Index As Integer)
Dim SQL As String
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
        Case 0 'calibres
            SQL = "¿Seguro que desea eliminar el Calibre?"
            SQL = SQL & vbCrLf & "Calibre: " & Adoaux(Index).Recordset!codcalib
            SQL = SQL & vbCrLf & "Nombre: " & Adoaux(Index).Recordset!nomcalib
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                SQL = "DELETE FROM calibres"
                SQL = SQL & vWhere & " AND codcalib= " & Adoaux(Index).Recordset!codcalib
            End If
            
        Case 1 'calidades
            SQL = "¿Seguro que desea eliminar la Calidad?"
            SQL = SQL & vbCrLf & "Código: " & Adoaux(Index).Recordset!codcalid
            SQL = SQL & vbCrLf & "Nombre: " & Adoaux(Index).Recordset!nomcalid
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                SQL = "DELETE FROM rcalidad "
                SQL = SQL & vWhere & " AND codcalid= " & Adoaux(Index).Recordset!codcalid
            End If
        
        Case 2 'variedades realcionadas
            SQL = "¿Seguro que desea eliminar la Variedad Relacionada?"
            SQL = SQL & vbCrLf & "Código: " & Adoaux(Index).Recordset!codvarie1
            SQL = SQL & vbCrLf & "Nombre: " & Adoaux(Index).Recordset!nomvarie
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                SQL = "DELETE FROM variedades_rel "
                SQL = SQL & " where codvarie = " & DBSet(Text1(0).Text, "N") & " AND numlinea= " & Adoaux(Index).Recordset!numlinea
            End If
            
        '[Monica]23/07/2018: bonificaciones y calibradores
        Case 3 ' bonificaciones
            SQL = "¿Seguro que desea eliminar la Bonificación?"
            SQL = SQL & vbCrLf & "Fecha: " & Adoaux(Index).Recordset!FechaEnt
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                SQL = "DELETE FROM rbonifentradas "
                SQL = SQL & " where codvarie = " & DBSet(Text1(0).Text, "N") & " AND fechaent= " & DBSet(Adoaux(Index).Recordset!FechaEnt, "F")
            End If
        
        Case 4 ' calibradores
            SQL = "¿Seguro que desea eliminar datos del calibrador?"
            SQL = SQL & vbCrLf & "Calidad: " & Adoaux(Index).Recordset!codcalid
            SQL = SQL & vbCrLf & "Línea: " & Adoaux(Index).Recordset!numlinea
            
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                SQL = "DELETE FROM rcalidad_calibrador "
                SQL = SQL & " where codvarie = " & DBSet(Text1(0).Text, "N") & " AND codcalid= " & Adoaux(Index).Recordset!codcalid
                SQL = SQL & " and numlinea = " & Adoaux(Index).Recordset!numlinea
            End If
        
            
    End Select

    If Eliminar Then
        NumRegElim = Adoaux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        conn.Execute SQL
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        CargaGrid Index, True
        If Not SituarDataTrasEliminar(Adoaux(Index), NumRegElim, True) Then
'            PonerCampos
            
        End If
        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
        ' *** si n'hi han tabs ***
        SituarTab (NumTabMto + 1)
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

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 0: vtabla = "calibres"
        Case 1: vtabla = "variane"
        Case 2: vtabla = "variedades_rel"
        
        '[Monica]23/07/2018: bonificaciones y calibradores
        Case 3: vtabla = "rbonifentradas"
        Case 4: vtabla = "rcalidad_calibrador"
    End Select
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0, 1, 2, 3, 4 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            If Index <> 3 And Index <> 4 Then
                If Index = 0 Then
                    NumF = SugerirCodigoSiguienteStr(vtabla, "codcalib", vWhere)
                Else
                    NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)
                End If
            End If
            
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
                Case 0 'calibres
                    txtAux(0).Text = Text1(0).Text 'codvarie
'                    txtAux(3).Text = text1(1).Text 'nomcalibre
                    txtAux(1).Text = NumF 'codcalib
                    txtAux(2).Text = ""
                    txtAux(3).Text = ""
                    txtAux(4).Text = ""
                    PonerFoco txtAux(1)
                    
                Case 1 'calidades
                    txtAux1(0).Text = Text1(0).Text 'codvarie
                    For i = 1 To txtAux1.Count - 1
                        txtAux1(i).Text = ""
                    Next i
                    Combo1(3).ListIndex = -1
                    Combo1(4).ListIndex = -1
                    chkAux(0).Value = 0
                    chkAux(1).Value = 0
                    
                    PonerFoco txtAux1(1)
                    
                Case 2 ' variedades relacionadas
                    txtAux(4).Text = Text1(0).Text 'codvarie
                    txtAux(5).Text = NumF 'numlinea
                    txtAux(6).Text = ""
                    txtAux2(6).Text = ""
                    
                    PonerFoco txtAux(6)
                    
                Case 3 ' bonificaciones
                    txtAux3(0).Text = Text1(0).Text 'codvarie
                    txtAux3(1).Text = "" 'fecha
                    txtAux3(2).Text = ""
                    
                    PonerFoco txtAux3(1)
                
                Case 4 ' calibradores
                    txtAux4(0).Text = Text1(0).Text 'codvarie
                    txtAux4(1).Text = "" 'codcalid
                    txtAux4(2).Text = "" ' numlinea
                    txtAux2(1).Text = ""
                    txtAux4(3).Text = ""
                    txtAux4(4).Text = ""
                    txtAux4(5).Text = ""
                    
                    PonerFoco txtAux4(1)
                
            End Select
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
        Case 0 'calibres
        
            For J = 0 To 3
                txtAux(J).Text = DataGridAux(Index).Columns(J).Text
            Next J
            
            For i = 0 To 1
                BloquearTxt txtAux(i), False
            Next i
            
        Case 1 'calidades
'            For J = 0 To 4
'                txtAux1(J).Text = DataGridAux(Index).Columns(J).Text
'            Next J
'                txtAux1(J).Text = DataGridAux(Index).Columns(J).Text
            txtAux1(0).Text = DataGridAux(Index).Columns(0).Text
            txtAux1(1).Text = DataGridAux(Index).Columns(1).Text
            txtAux1(2).Text = DataGridAux(Index).Columns(2).Text
            txtAux1(3).Text = DataGridAux(Index).Columns(3).Text
            txtAux1(4).Text = DataGridAux(Index).Columns(8).Text
            txtAux1(5).Text = DataGridAux(Index).Columns(9).Text
            ' ***** canviar-ho pel nom del camp del combo *********
            i = Adoaux(Index).Recordset!tipcalid
            ' *****************************************************
            PosicionarCombo Me.Combo1(3), i
            
            i = Adoaux(Index).Recordset!tipcalid1
            ' *****************************************************
            PosicionarCombo Me.Combo1(4), i
            
            Me.chkAux(0).Value = Me.Adoaux(Index).Recordset!gastosrec
            
            '[Monica]11/12/2013
            txtAux1(6).Text = DataGridAux(Index).Columns(12).Text
            txtAux1(7).Text = DataGridAux(Index).Columns(13).Text
            
            Me.chkAux(1).Value = Me.Adoaux(Index).Recordset!seaplicabonif
            
            
            For i = 0 To 0
                BloquearTxt txtAux1(i), False
            Next i
            
        Case 2 'variedades relacionadas
            For J = 4 To 5
                txtAux(J).Text = DataGridAux(Index).Columns(J - 4).Text
            Next J
            
            For J = 6 To 6
                txtAux(J).Text = DataGridAux(Index).Columns(J - 4).Text
                txtAux2(J).Text = DataGridAux(Index).Columns(J - 3).Text
            Next J
            
            For i = 6 To 6
                BloquearTxt txtAux(i), False
            Next i
            BloquearBtn btnBuscar(0), False
            
        Case 3 ' bonificaciones
            For J = 0 To 2
                txtAux3(J).Text = DataGridAux(Index).Columns(J).Text
            Next J
            BloquearBtn btnBuscar(1), False
            
        Case 4 ' calibradores
            For J = 0 To 6
                If J = 2 Then
                    txtAux2(1).Text = DataGridAux(Index).Columns(J).Text
                Else
                    If J < 2 Then
                        txtAux4(J).Text = DataGridAux(Index).Columns(J).Text
                    Else
                        txtAux4(J - 1).Text = DataGridAux(Index).Columns(J).Text
                    End If
                End If
            Next J
            BloquearBtn btnBuscar(2), False
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'cuentas bancarias
            PonerFoco txtAux(2)
        Case 1 'departamentos
            PonerFoco txtAux(10)
        Case 2
            PonerFoco txtAux(6)
        Case 3
            PonerFoco txtAux3(2)
        Case 4
            PonerFoco txtAux4(1)
    End Select
    ' ***************************************************************************************
End Sub

Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim B As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    DeseleccionaGrid DataGridAux(Index)
       
    B = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 0 'calibres
            For jj = 1 To 3
                txtAux(jj).visible = B
                txtAux(jj).Top = alto
            Next jj
            
        Case 1 'calidades
            For jj = 1 To 7
                txtAux1(jj).visible = B
                txtAux1(jj).Top = alto
            Next jj
            Combo1(3).visible = B
            Combo1(3).Top = alto
            Combo1(4).visible = B
            Combo1(4).Top = alto
            chkAux(0).visible = B
            chkAux(0).Top = alto
            chkAux(1).visible = B
            chkAux(1).Top = alto
            
        Case 2 ' variedades relacionadas
            For jj = 6 To 6
                txtAux(jj).visible = B
                txtAux(jj).Top = alto
            Next jj
            
            txtAux2(6).visible = B
            txtAux2(6).Top = alto
            
            Me.btnBuscar(0).visible = B
            Me.btnBuscar(0).Top = alto
            
        '[Monica]23/07/2018: bonificaciones y calibradores
        Case 3 ' bonificaciones
            For jj = 1 To 2
                txtAux3(jj).visible = B
                txtAux3(jj).Top = alto
            Next jj
            
            Me.btnBuscar(1).visible = (xModo = 1)
            Me.btnBuscar(1).Top = alto
        
        Case 4 ' calibradores
            For jj = 1 To 5
                txtAux4(jj).visible = B
                txtAux4(jj).Top = alto
            Next jj
            
            txtAux2(1).visible = B
            txtAux2(1).Top = alto
            
            Me.btnBuscar(2).visible = B
            Me.btnBuscar(2).Top = alto
        
    End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'copiar calibres
            mnCopiaCalibres_Click
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
            
        Case 2, 3 ' Nombre de calibre y calibre abreviado
            txtAux(Index).Text = UCase(txtAux(Index).Text)
            
        Case 6 ' variedad relacionada
            If txtAux(Index).Text = "" Then Exit Sub
            txtAux2(6).Text = PonerNombreDeCod(txtAux(Index), "variedades", "nomvarie", "codvarie", "N")
            
            If txtAux2(6).Text <> "" Then Me.cmdAceptar.SetFocus
            
        Case 10 ' variedad anecoop
            txtAux(Index).Text = UCase(txtAux(Index).Text)
            
    End Select
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
   If Not txtAux(Index).MultiLine Then ConseguirFocoLin txtAux(Index)
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not txtAux(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not txtAux(Index).MultiLine Then
        If KeyAscii = teclaBuscar Then
            If Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) Then
            End If
        Else
            KEYpress KeyAscii
        End If
    End If
End Sub

Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim B As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte

    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False
        
    B = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not B Then Exit Function
    
    
    If nomframe = "FrameAux3" Then
        If (ModoLineas = 1) Then 'insertar
            'comprobar si existe ya el cod. del campo clave primaria
            SQL = "select count(*) from rbonifentradas where codvarie = " & DBSet(txtAux3(0).Text, "N") & " and fechaent = " & DBSet(txtAux3(1).Text, "F")
            If TotalRegistros(SQL) <> 0 Then
                MsgBox "Ya existe el registro para la variedad fecha. Revise.", vbExclamation
                PonerFoco txtAux3(1)
                B = False
            End If
        End If
    End If
    
    If nomframe = "FrameAux4" Then
        If ModoLineas = 1 Then   'Estamos insertando
            SQL = DevuelveDesdeBDNew(cAgro, "rcalidad_calibrador", "codcalid", "codvarie", txtAux4(0).Text, "N", , "codcalid", txtAux4(1).Text, "N", "numlinea", txtAux4(2).Text, "N")
            If SQL <> "" Then
                MsgBox "Linea de calibrador existente para esta calidad. Reintroduzca.", vbExclamation
                PonerFoco txtAux4(1)
                B = False
            End If
        End If
        
        If B And (ModoLineas = 1 Or ModoLineas = 2) Then
            SQL = "select count(*) from rcalidad_calibrador where codvarie = " & DBSet(txtAux4(0).Text, "N")
            SQL = SQL & " and codcalid <> " & DBSet(txtAux4(1).Text, "N")
            SQL = SQL & " and nomcalibrador1 = " & DBSet(txtAux4(3).Text, "T")
        
            If TotalRegistros(SQL) <> 0 Then
                MsgBox "El nombre que utiliza el Calibrador 1 de esta calidad está asignada a otra. Revise.", vbExclamation
                PonerFoco txtAux4(3)
                B = False
            End If
        End If
        
        If B And (ModoLineas = 1 Or ModoLineas = 2) Then
            SQL = "select count(*) from rcalidad_calibrador where codvarie = " & DBSet(txtAux4(0).Text, "N")
            SQL = SQL & " and codcalid <> " & DBSet(txtAux4(1).Text, "N")
            SQL = SQL & " and nomcalibrador2 = " & DBSet(txtAux4(4).Text, "T")
        
            If TotalRegistros(SQL) <> 0 Then
                MsgBox "El nombre que utiliza el Calibrador 2 de esta calidad está asignada a otra. Revise.", vbExclamation
                PonerFoco txtAux4(4)
                B = False
            End If
        End If
        
        If B And (ModoLineas = 3 Or ModoLineas = 4) Then
            If txtAux(5).Text <> "" Then
                SQL = "select count(*) from rcalidad_calibrador where codvarie = " & DBSet(txtAux4(0).Text, "N")
                SQL = SQL & " and codcalid <> " & DBSet(txtAux4(1).Text, "N")
                SQL = SQL & " and nomcalibrador3 = " & DBSet(txtAux4(5).Text, "T")
            
                If TotalRegistros(SQL) <> 0 Then
                    MsgBox "El nombre que utiliza el Calibrador 3 de esta calidad está asignada a otra. Revise.", vbExclamation
                    PonerFoco txtAux4(5)
                    B = False
                End If
            End If
        End If
    End If
    
    
    
    
    ' ******************************************************************************
    DatosOkLlin = B
    
EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function SepuedeBorrar(ByRef Index As Integer) As Boolean
    SepuedeBorrar = False
    
    ' *** si cal comprovar alguna cosa abans de borrar ***
'    Select Case Index
'        Case 0 'cuentas bancarias
'            If AdoAux(Index).Recordset!ctaprpal = 1 Then
'                MsgBox "No puede borrar una Cuenta Principal. Seleccione antes otra cuenta como Principal", vbExclamation
'                Exit Function
'            End If
'    End Select
    ' ****************************************************
    
    SepuedeBorrar = True
End Function

Private Sub imgBuscar_Click(Index As Integer)
Dim numnivel As Byte

    TerminaBloquear
     Select Case Index
        Case 0 'productos
            AbrirFrmProducto (2)

'[Monica]21/10/2013: cambiado pq viene de comercial
'            Set frmPro = New frmManProductos
'            frmPro.DatosADevolverBusqueda = "0|1|"
'            frmPro.CodigoActual = text1(2).Text
'            frmPro.Show vbModal
'            Set frmPro = Nothing
'            PonerFoco text1(2)
            
        Case 1 'clases
            AbrirFrmClase (3)
        
'[Monica]21/10/2013: cambiado pq viene de comercial
'            Set frmCla = New frmManClases
'            frmCla.DatosADevolverBusqueda = "0|1|"
'            frmCla.CodigoActual = text1(3).Text
'            frmCla.Show vbModal
'            Set frmCla = Nothing
'            PonerFoco text1(3)
        
        Case 2 'tipos de unidad
            AbrirFrmTUnidad (2)
            
'[Monica]21/10/2013: cambiado pq viene de comercial
'            Set frmTUn = New frmManTipUnid
'            frmTUn.DatosADevolverBusqueda = "0|1|"
'            frmTUn.CodigoActual = Text1(3).Text
'            frmTUn.Show vbModal
'            Set frmTUn = Nothing
'            PonerFoco Text1(9)
        
        Case 3, 5, 6, 7, 8 'cuenta contable de venta
            If vParamAplic.NumeroConta = 0 Then Exit Sub
        
            If Index = 3 Then
                Indice = Index + 18
            Else
                Indice = Index + 17
            End If
            
            Set frmCtas = New frmCtasConta
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(Indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(Indice)
            
       Case 13, 14, 15 ' cuentas contables de recoleccion
            If vParamAplic.NumeroConta = 0 Then Exit Sub
        
            Indice = Index + 18
            
            Set frmCtas = New frmCtasConta
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(Indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(Indice)
       
                   
       Case 9  'Porcentaje iva
            If vParamAplic.NumeroConta = 0 Then Exit Sub
            Indice = 26
            Set frmIva = New frmTipIVAConta
            frmIva.DeConsulta = True
            frmIva.DatosADevolverBusqueda = "0|1|2|"
            frmIva.CodigoActual = Text1(Indice).Text
            frmIva.Show vbModal
            Set frmIva = Nothing
            PonerFoco Text1(Indice)
       
       Case 10 'Centro de Coste
            If vParamAplic.NumeroConta = 0 Then Exit Sub
            Indice = 27
            Set frmCCos = New frmCCosConta
            frmCCos.DatosADevolverBusqueda = "0|1|"
            frmCCos.CodigoActual = Text1(Indice).Text
            frmCCos.Show vbModal
            Set frmCCos = Nothing
            PonerFoco Text1(Indice)
    
        Case 11, 12 'cuentas contables de transporte
            If vParamAplic.NumeroConta = 0 Then Exit Sub
        
            Indice = Index + 17
            
            Set frmCtas = New frmCtasConta
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(Indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(Indice)
    
        Case 16, 17, 18, 19
            If vParamAplic.NumeroConta = 0 Then Exit Sub
        
            Indice = Index + 18
            
            Set frmCtas = New frmCtasConta
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(Indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(Indice)
    
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

Private Sub frmCla_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de Clases
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codclase
    FormateaCampo Text1(3)
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'nomclase
End Sub

Private Sub frmPro_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento Porductos
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codprodu
    FormateaCampo Text1(2)
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2) 'nomprodu
End Sub

Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte

    If ModoLineas <> 1 Then
        Select Case Index
            Case 0 'cuentas bancarias
                If DataGridAux(Index).Columns.Count > 2 Then
'                    txtAux(11).Text = DataGridAux(Index).Columns("direccio").Text
'                    txtAux(12).Text = DataGridAux(Index).Columns("observac").Text
                End If
                
            Case 1 'departamentos
                If DataGridAux(Index).Columns.Count > 2 Then
'                    txtAux(21).Text = DataGridAux(Index).Columns(5).Text
'                    txtAux(22).Text = DataGridAux(Index).Columns(6).Text
'                    txtAux(23).Text = DataGridAux(Index).Columns(8).Text
'                    txtAux(24).Text = DataGridAux(Index).Columns(15).Text
'                    txtAux2(22).Text = DataGridAux(Index).Columns(7).Text
                End If
                
        End Select
        
    Else 'vamos a Insertar
        Select Case Index
            Case 0 'cuentas bancarias
'                txtAux(11).Text = ""
'                txtAux(12).Text = ""
            Case 1 'departamentos
                For i = 21 To 24
'                   txtAux(i).Text = ""
                Next i
'               txtAux2(22).Text = ""
            Case 2 'Tarjetas
'               txtAux(50).Text = ""
'               txtAux(51).Text = ""
        End Select
    End If
End Sub

' ***** si n'hi han varios nivells de tabs *****
Private Sub SituarTab(numTab As Integer)
    On Error Resume Next
    
    If numTab >= 3 Then numTab = numTab + 1
    
    
    SSTab1.Tab = numTab
    
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
    Else
        ' *** si n'hi han tabs sense datagrids, li pose els valors als camps ***
        NetejaFrameAux "FrameAux3" 'neteja només lo que te TAG
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

Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
Dim B As Boolean
Dim i As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    CargaGridGnral Me.DataGridAux(Index), Me.Adoaux(Index), tots, PrimeraVez
    
    Select Case Index
        Case 0 'calibres
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;S|txtAux(1)|T|Cod|600|;" 'codvarie,codcalib
            tots = tots & "S|txtAux(2)|T|Nombre|3500|;" ' nombre del calibre
            tots = tots & "S|txtAux(3)|T|Abrev.|1000|;" ' nombre de calibre abreviado
            
            arregla tots, DataGridAux(Index), Me, 350
        
            B = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
        Case 1 'Calidades
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;S|txtAux1(1)|T|Código|850|;" 'codvarie,codcalid
            tots = tots & "S|txtAux1(2)|T|Nombre|3760|;"
            tots = tots & "S|txtAux1(3)|T|Abrev.|1100|;"
            tots = tots & "N||||0|;S|Combo1(3)|C|Tipo|2020|;"
            tots = tots & "N||||0|;S|Combo1(4)|C|Tipo|1920|;"
            tots = tots & "S|txtAux1(4)|T|Calibrador 1|1500|;"
            tots = tots & "S|txtAux1(5)|T|Calibrador 2|1500|;"
            tots = tots & "N||||0|;S|chkAux(0)|CB|GR|360|;"
            tots = tots & "S|txtAux1(6)|T|Eu.R.Soc|1200|;S|txtAux1(7)|T|Eu.R.Coop|1200|;"
            tots = tots & "N||||0|;S|chkAux(1)|CB|Bo|360|;"
            
            arregla tots, DataGridAux(Index), Me, 350

            B = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
            
        Case 2 ' Variedades relacionadas
            tots = "N||||0|;N||||0|;S|txtAux(6)|T|Codigo|1000|;" 'codvarie,codcalib
            tots = tots & "S|btnBuscar(0)|B|||;S|txtAux2(6)|T|Variedad|4500|;" ' nombre del calibre
            
            arregla tots, DataGridAux(Index), Me, 350
        
            B = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
        
        
        Case 3 ' Bonificaciones
            tots = "N||||0|;" 'codvarie
            tots = tots & "S|txtAux3(1)|T|Fecha Entrada|2000|;"
            tots = tots & "S|btnBuscar(1)|B|||;S|txtAux3(2)|T|Porcentaje|1650|;"
            
            arregla tots, DataGridAux(Index), Me, 350
        
            B = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
        
        Case 4 ' Calibradores
            tots = "N||||0|;"
            tots = tots & "S|txtAux4(1)|T|Calidad|900|;S|btnBuscar(2)|B|||;S|txtAux2(1)|T|Denominación|2380|;"
            tots = tots & "S|txtAux4(2)|T|Linea|900|;"
            ' este programa solo lo ven catadau(0) y alzira(4)
            '[Monica]29/02/2012: Natural era la cooperativa 0 junto con Catadau ahora es la 9
            If vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 9 Or vParamAplic.Cooperativa = 18 Then
                tots = tots & "S|txtAux4(3)|T|Calibrador 1|2000|;S|txtAux4(4)|T|Calibrador 2|2000|;"
                tots = tots & "S|txtAux4(5)|T|Calibrador 3|2000|;"
            Else
                tots = tots & "S|txtAux4(3)|T|Precalibradora|2000|;S|txtAux4(4)|T|Escandalladora|2000|;"
                tots = tots & "S|txtAux4(5)|T|Calibrador Kaki|2000|;"
            End If
            
            arregla tots, DataGridAux(Index), Me, 350
            
    End Select
    
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
    ' **** si n'hi han llínies en grids i camps fora d'estos ****
    If Not Adoaux(Index).Recordset.EOF Then
        DataGridAux_RowColChange Index, 1, 1
    Else
'        LimpiarCamposFrame Index
    End If
      
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub

Private Sub InsertarLinea()
'Inserta registre en les taules de Llínies
Dim nomframe As String
Dim B As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'calibres
        Case 1: nomframe = "FrameAux1" 'variedades anecoop
        Case 2: nomframe = "FrameAux2" ' variedades relacionadas
        Case 3: nomframe = "FrameAux3" ' bonificaciones
        Case 4: nomframe = "FrameAux4" ' calibradores
    End Select
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomframe) Then
            B = BLOQUEADesdeFormulario2(Me, Data1, 1)
            Select Case NumTabMto
                Case 0, 1, 2, 3, 4 ' *** els index de les llinies en grid (en o sense tab) ***
                     CargaGrid NumTabMto, True
                    If B Then BotonAnyadirLinea NumTabMto
            End Select
           
'            SituarTab (NumTabMto + 1)
        End If
    End If
End Sub

Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Integer
    
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'calibres
        Case 1: nomframe = "FrameAux1" 'variedades anecoop
        Case 2: nomframe = "FrameAux2" 'variedades relacionadas
        Case 3: nomframe = "FrameAux3" 'bonificaciones
        Case 4: nomframe = "FrameAux4" 'calibradores
    End Select
    
    ModificarLinea = False
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
            ModoLineas = 0
            
            V = Adoaux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
            CargaGrid NumTabMto, True
            
            ' *** si n'hi han tabs ***
            SituarTab (NumTabMto + 1)

            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            PonerFocoGrid Me.DataGridAux(NumTabMto)
            Adoaux(NumTabMto).Recordset.Find (Adoaux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
            
            LLamaLineas NumTabMto, 0
            ModificarLinea = True
        End If
    End If
End Function

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " codvarie=" & Val(Text1(0).Text)
    
    ObtenerWhereCab = vWhere
End Function

'' *** neteja els camps dels tabs de grid que
''estan fora d'este, i els camps de descripció ***
Private Sub LimpiarCamposFrame(Index As Integer)
    On Error Resume Next
 
'    Select Case Index
'        Case 0 'Cuentas Bancarias
'            txtAux(11).Text = ""
'            txtAux(12).Text = ""
'        Case 1 'Departamentos
'            txtAux(21).Text = ""
'            txtAux(22).Text = ""
'            txtAux2(22).Text = ""
'            txtAux(23).Text = ""
'            txtAux(24).Text = ""
'        Case 2 'Tarjetas
'            txtAux(50).Text = ""
'            txtAux(51).Text = ""
'        Case 4 'comisiones
'            txtAux2(2).Text = ""
'    End Select
'
    If Err.Number <> 0 Then Err.Clear
End Sub

' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGridAux_GotFocus(Index As Integer)
'  WheelHook DataGridAux(Index)
'End Sub
'Private Sub DataGridAux_LostFocus(Index As Integer)
'  WheelUnHook
'End Sub

' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo()
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 0 To Combo1.Count - 1
        Combo1(i).Clear
    Next i
    
    Combo1(0).AddItem "Campo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Almacén"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    

    Combo1(1).AddItem "Cooperativa"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Ajena"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    
    Combo1(2).AddItem "Convencional"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 0
    Combo1(2).AddItem "Biológica"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 1


    'Tipo de Calidad
    Combo1(3).AddItem "Normal"
    Combo1(3).ItemData(Combo1(3).NewIndex) = 0
    Combo1(3).AddItem "Destrio"
    Combo1(3).ItemData(Combo1(3).NewIndex) = 1
    Combo1(3).AddItem "Venta Campo"
    Combo1(3).ItemData(Combo1(3).NewIndex) = 2
    Combo1(3).AddItem "Merma"
    Combo1(3).ItemData(Combo1(3).NewIndex) = 3
    Combo1(3).AddItem "Pequeño"
    Combo1(3).ItemData(Combo1(3).NewIndex) = 4
    Combo1(3).AddItem "Pixat"
    Combo1(3).ItemData(Combo1(3).NewIndex) = 5
    Combo1(3).AddItem "Destrio Comercial"
    Combo1(3).ItemData(Combo1(3).NewIndex) = 6
    
    
    ' Tipo de calidad 2
    Combo1(4).AddItem "Comercial"
    Combo1(4).ItemData(Combo1(4).NewIndex) = 0
    Combo1(4).AddItem "No Comercial"
    Combo1(4).ItemData(Combo1(4).NewIndex) = 1
    Combo1(4).AddItem "Retirada"
    Combo1(4).ItemData(Combo1(4).NewIndex) = 2

End Sub

Private Sub AbrirFrmClase(Indice As Integer)

    indCodigo = 3
    
    Set frmCla = New frmBasico2
    
    AyudaClasesCom frmCla, Text1(indCodigo).Text
        
    Set frmCla = Nothing

End Sub

Private Sub AbrirFrmProducto(Indice As Integer)
    
    indCodigo = 2
    
    Set frmProd = New frmBasico2
    
    AyudaProductosCom frmProd, Text1(indCodigo).Text
    
    Set frmProd = Nothing
    
End Sub

Private Sub AbrirFrmTUnidad(Indice As Integer)

    indCodigo = 3
    
    Set frmTUn = New frmBasico2
    
    AyudaTUnidadesCom frmTUn, Text1(indCodigo).Text
        
    Set frmCla = Nothing

End Sub

Private Sub txtAux1_GotFocus(Index As Integer)
    ConseguirFoco txtAux1(Index), Modo
End Sub

Private Sub txtAux1_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux1(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 1 'codigo de calibre
            PonerFormatoEntero txtAux1(Index)
            
        Case 2, 3 'nombre de calibre y de calibre abreviado
            txtAux1(Index).Text = UCase(txtAux1(Index).Text)
            
        Case 6, 7 'precios
            PonerFormatoDecimal txtAux1(Index), 6
    End Select
    
End Sub

Private Sub txtAux1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 0: KEYBusqueda KeyAscii, 0 'variedad
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

' bonificaciones
Private Sub TxtAux3_GotFocus(Index As Integer)
    ConseguirFoco txtAux3(Index), Modo
End Sub


Private Sub TxtAux3_LostFocus(Index As Integer)
Dim cadMen As String

    If Not PerderFocoGnral(txtAux3(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 1 'fecha de entrada
            PonerFormatoFecha txtAux3(Index), True
        
        Case 2 ' porcentaje
            If txtAux3(Index).Text <> "" Then
                If PonerFormatoDecimal(txtAux3(Index), 10) Then PonerFocoBtn cmdAceptar
            Else
                PonerFocoBtn cmdAceptar
            End If
            
    End Select
End Sub

Private Sub TxtAux3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


' calibradores
Private Sub TxtAux4_GotFocus(Index As Integer)
    ConseguirFoco txtAux4(Index), Modo
End Sub


Private Sub TxtAux4_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux4(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 1 'codigo de calibre
            '[Monica]05/11/2012: dejar buscar por calibre
            If Modo = 1 Then Exit Sub
        
            If Not PonerFormatoEntero(txtAux4(Index)) Then Exit Sub
            txtAux2(1).Text = DevuelveDesdeBDNew(cAgro, "rcalidad", "nomcalid", "codvarie", txtAux4(0).Text, "N", , "codcalid", txtAux4(1).Text, "N")
            If txtAux2(1).Text = "" Then
                MsgBox "Calidad no existe. Reintroduzca.", vbExclamation
                PonerFoco txtAux4(Index)
            End If
            
        Case 2 'numlinea
            PonerFormatoEntero txtAux4(Index)
        
        Case 5
            PonerFocoBtn cmdAceptar

        
    End Select
    
End Sub

Private Sub TxtAux4_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 5 Then
            Select Case Index
                Case 1: KEYBusqueda KeyAscii, 1 'calidad
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub TxtAux4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

