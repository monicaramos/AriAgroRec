VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPOZHidrantesIndefa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hidrantes Indefa"
   ClientHeight    =   10455
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   16560
   Icon            =   "frmPOZHidrantesIndefa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10455
   ScaleWidth      =   16560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Cuenta Principal|N|N|0|1|cltebanc|ctaprpal|0||"
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   225
      TabIndex        =   228
      Top             =   135
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   229
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
      Left            =   3885
      TabIndex        =   226
      Top             =   135
      Width           =   1245
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   227
         Top             =   180
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Buscar Diferencias"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Actualizar Registros"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   5205
      TabIndex        =   224
      Top             =   135
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   225
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
      Left            =   13950
      TabIndex        =   223
      Top             =   180
      Width           =   1605
   End
   Begin VB.Frame Frame2 
      Height          =   2925
      Left            =   255
      TabIndex        =   23
      Top             =   930
      Width           =   16185
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
         Index           =   10
         Left            =   14445
         MaxLength       =   10
         TabIndex        =   17
         Tag             =   "Fecha Lectura Actual|F|S|||rpozos|fech_act|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   945
         Width           =   1305
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
         Left            =   11145
         MaxLength       =   7
         TabIndex        =   16
         Tag             =   "Contador Actual|N|S|||rpozos|lect_act|######0||"
         Text            =   "1234567"
         Top             =   915
         Width           =   1035
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
         Left            =   11145
         MaxLength       =   7
         TabIndex        =   14
         Tag             =   "Lectura Anterior|N|S|||rpozos|lect_ant|######0||"
         Text            =   "1234567"
         Top             =   540
         Width           =   1035
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
         Left            =   14445
         MaxLength       =   10
         TabIndex        =   15
         Tag             =   "Fecha lectura anterior|F|S|||rpozos|fech_ant|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   540
         Width           =   1305
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
         Height          =   525
         Index           =   11
         Left            =   10155
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Tag             =   "Observaciones|T|S|||rpozos|observac|||"
         Top             =   2130
         Width           =   5745
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
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   8
         Tag             =   "Parcelas|T|S|||rpozos|parcelas||N|"
         Text            =   "1234567890123456789012345"
         Top             =   2475
         Width           =   3120
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
         Index           =   4
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Polígono|T|S|||rpozos|poligono||N|"
         Text            =   "1234567890"
         Top             =   2115
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
         Index           =   1
         Left            =   3705
         MaxLength       =   40
         TabIndex        =   62
         Top             =   2115
         Width           =   3405
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
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   6
         Tag             =   "Partida|N|N|1|9999|rpozos|codparti|0000||"
         Top             =   1755
         Width           =   855
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
         Left            =   2220
         MaxLength       =   40
         TabIndex        =   60
         Top             =   1755
         Width           =   4890
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
         Index           =   13
         Left            =   2220
         MaxLength       =   30
         TabIndex        =   58
         Top             =   1380
         Width           =   4890
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
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "Pozo|N|N|0|999|rpozos|codpozo|000||"
         Top             =   1380
         Width           =   840
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
         Index           =   18
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   4
         Tag             =   "Campo|N|S|1|99999999|rpozos|codcampo|00000000||"
         Top             =   1020
         Width           =   1155
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
         Left            =   8730
         MaxLength       =   10
         TabIndex        =   11
         Tag             =   "Hanegadas|N|S|0|9999.99|rpozos|hanegada|###0.0000||"
         Text            =   "1234567890"
         Top             =   1335
         Width           =   1305
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
         Index           =   14
         Left            =   8730
         MaxLength       =   4
         TabIndex        =   12
         Tag             =   "Calibre|N|S|||rpozos|calibre|###0|N|"
         Text            =   "1234"
         Top             =   1710
         Width           =   1305
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
         Left            =   8730
         MaxLength       =   7
         TabIndex        =   13
         Tag             =   "Acciones|N|S|||rpozos|nroacciones|#,###,##0|N|"
         Text            =   "1234567"
         Top             =   2085
         Width           =   1305
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
         Index           =   16
         Left            =   8730
         MaxLength       =   10
         TabIndex        =   9
         Tag             =   "Fecha Alta|F|N|||rpozos|fechaalta|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   390
         Width           =   1350
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
         Index           =   17
         Left            =   8730
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "Fecha Alta|F|S|||rpozos|fechabaja|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   795
         Width           =   1350
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
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Socio|N|N|1|999999|rpozos|codsocio|000000||"
         Top             =   660
         Width           =   840
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
         Left            =   2220
         MaxLength       =   30
         TabIndex        =   50
         Top             =   660
         Width           =   4890
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
         Left            =   6810
         MaxLength       =   1
         TabIndex        =   2
         Tag             =   "Digito Control|T|N|||rpozos|digcontrol|||"
         Top             =   240
         Width           =   300
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
         Left            =   3630
         MaxLength       =   6
         TabIndex        =   1
         Tag             =   "Toma|N|S|0|999999|rpozos|nroorden|000000||"
         Top             =   240
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
         Index           =   0
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Contador TCH|T|N|||rpozos|hidrante||S|"
         Top             =   240
         Width           =   1455
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
         Index           =   19
         Left            =   11145
         MaxLength       =   9
         TabIndex        =   109
         Tag             =   "Consumo|N|S|||rpozos|consumo|########0||"
         Text            =   "1234567"
         Top             =   1395
         Width           =   1035
      End
      Begin VB.Line Line1 
         X1              =   11025
         X2              =   12345
         Y1              =   1335
         Y2              =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Consumo"
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
         Left            =   10185
         TabIndex        =   108
         Top             =   1455
         Width           =   1035
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   14145
         Picture         =   "frmPOZHidrantesIndefa.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   945
         Width           =   240
      End
      Begin VB.Label Label12 
         Caption         =   "Fecha Lectura"
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
         Left            =   12645
         TabIndex        =   107
         Top             =   945
         Width           =   1605
      End
      Begin VB.Label Label9 
         Caption         =   "Actual"
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
         Left            =   10185
         TabIndex        =   106
         Top             =   945
         Width           =   1035
      End
      Begin VB.Label Label23 
         Caption         =   "Anterior"
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
         Left            =   10185
         TabIndex        =   105
         Top             =   570
         Width           =   1125
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   14145
         Picture         =   "frmPOZHidrantesIndefa.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   540
         Width           =   240
      End
      Begin VB.Label Label18 
         Caption         =   "Fecha Lectura"
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
         Left            =   12645
         TabIndex        =   104
         Top             =   570
         Width           =   1605
      End
      Begin VB.Label Label180 
         Caption         =   "Lecturas"
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
         Height          =   195
         Left            =   10185
         TabIndex        =   103
         Top             =   180
         Width           =   945
      End
      Begin VB.Line Line3 
         X1              =   10185
         X2              =   15825
         Y1              =   1770
         Y2              =   1770
      End
      Begin VB.Line Line2 
         X1              =   10185
         X2              =   15825
         Y1              =   420
         Y2              =   420
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
         Left            =   10155
         TabIndex        =   66
         Top             =   1860
         Width           =   1545
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   11715
         Tag             =   "-1"
         ToolTipText     =   "Zoom descripción"
         Top             =   1860
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Parcelas"
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
         TabIndex        =   65
         Top             =   2475
         Width           =   810
      End
      Begin VB.Label Label3 
         Caption         =   "Polígono"
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
         TabIndex        =   64
         Top             =   2100
         Width           =   810
      End
      Begin VB.Label Label15 
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
         Left            =   2730
         TabIndex        =   63
         Top             =   2115
         Width           =   990
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1020
         ToolTipText     =   "Buscar Partida"
         Top             =   1755
         Width           =   240
      End
      Begin VB.Label Label5 
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
         Left            =   210
         TabIndex        =   61
         Top             =   1755
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Pozo"
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
         TabIndex        =   59
         Top             =   1395
         Width           =   690
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1020
         ToolTipText     =   "Buscar Pozo"
         Top             =   1395
         Width           =   240
      End
      Begin VB.Label Label13 
         Caption         =   "Campo"
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
         TabIndex        =   57
         Top             =   1005
         Width           =   690
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1020
         ToolTipText     =   "Buscar Campo"
         Top             =   1035
         Width           =   240
      End
      Begin VB.Label Label41 
         Caption         =   "Hanegadas"
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
         Left            =   7260
         TabIndex        =   56
         Top             =   1365
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Calibre"
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
         Left            =   7260
         TabIndex        =   55
         Top             =   1755
         Width           =   810
      End
      Begin VB.Label Label8 
         Caption         =   "Acciones"
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
         Left            =   7260
         TabIndex        =   54
         Top             =   2085
         Width           =   930
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   8430
         Picture         =   "frmPOZHidrantesIndefa.frx":0122
         ToolTipText     =   "Buscar fecha"
         Top             =   390
         Width           =   240
      End
      Begin VB.Label Label10 
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
         Left            =   7260
         TabIndex        =   53
         Top             =   420
         Width           =   1275
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   8430
         Picture         =   "frmPOZHidrantesIndefa.frx":01AD
         ToolTipText     =   "Buscar fecha"
         Top             =   795
         Width           =   240
      End
      Begin VB.Label Label11 
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
         Left            =   7260
         TabIndex        =   52
         Top             =   825
         Width           =   1185
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1020
         ToolTipText     =   "Buscar Socio"
         Top             =   675
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
         Left            =   210
         TabIndex        =   51
         Top             =   645
         Width           =   690
      End
      Begin VB.Label Label16 
         Caption         =   "Dígito Control"
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
         Left            =   5460
         TabIndex        =   37
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label Label26 
         Caption         =   "Toma"
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
         Left            =   3015
         TabIndex        =   27
         Top             =   270
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Contador TCH"
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
         TabIndex        =   24
         Top             =   270
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   270
      TabIndex        =   21
      Top             =   9765
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
         TabIndex        =   22
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
      Left            =   15360
      TabIndex        =   20
      Top             =   9945
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
      Left            =   14160
      TabIndex        =   19
      Top             =   9945
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
      Left            =   15315
      TabIndex        =   25
      Top             =   9945
      Visible         =   0   'False
      Width           =   1065
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5700
      Left            =   270
      TabIndex        =   26
      Top             =   4005
      Width           =   16155
      _ExtentX        =   28496
      _ExtentY        =   10054
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   6
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
      TabCaption(0)   =   "Datos Indefa"
      TabPicture(0)   =   "frmPOZHidrantesIndefa.frx":0238
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSTab2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Coopropietarios"
      TabPicture(1)   =   "frmPOZHidrantesIndefa.frx":0254
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameAux0"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Campos"
      TabPicture(2)   =   "frmPOZHidrantesIndefa.frx":0270
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameAux1"
      Tab(2).ControlCount=   1
      Begin TabDlg.SSTab SSTab2 
         Height          =   5205
         Left            =   120
         TabIndex        =   68
         Top             =   420
         Width           =   15780
         _ExtentX        =   27834
         _ExtentY        =   9181
         _Version        =   393216
         Tabs            =   5
         Tab             =   4
         TabsPerRow      =   5
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Contadores"
         TabPicture(0)   =   "frmPOZHidrantesIndefa.frx":028C
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "FrameAux2"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Hidrantes"
         TabPicture(1)   =   "frmPOZHidrantesIndefa.frx":02A8
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "FrameAux4"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Válvulas"
         TabPicture(2)   =   "frmPOZHidrantesIndefa.frx":02C4
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "FrameAux5"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Ventosas"
         TabPicture(3)   =   "frmPOZHidrantesIndefa.frx":02E0
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "FrameAux6"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Desagües"
         TabPicture(4)   =   "frmPOZHidrantesIndefa.frx":02FC
         Tab(4).ControlEnabled=   -1  'True
         Tab(4).Control(0)=   "FrameAux3"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).ControlCount=   1
         Begin VB.Frame FrameAux3 
            BorderStyle     =   0  'None
            Height          =   4695
            Left            =   120
            TabIndex        =   200
            Top             =   420
            Width           =   15405
            Begin VB.TextBox txtaux8 
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
               Height          =   555
               Index           =   8
               Left            =   2460
               MaxLength       =   250
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   211
               Tag             =   "Hanegadas|N|S|||rae_visitas_hidtomas|hanegadas||N|"
               Top             =   4080
               Width           =   12720
            End
            Begin VB.TextBox txtaux8 
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
               Index           =   7
               Left            =   2460
               MaxLength       =   250
               ScrollBars      =   2  'Vertical
               TabIndex        =   210
               Tag             =   "the_geom|T|S|||rae_visitas_hidtomas|the_geom||N|"
               Top             =   3180
               Width           =   2865
            End
            Begin VB.TextBox txtaux8 
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
               Index           =   5
               Left            =   2460
               MaxLength       =   50
               TabIndex        =   209
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|observaciones||N|"
               Top             =   2250
               Width           =   2865
            End
            Begin VB.TextBox txtaux8 
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
               Index           =   4
               Left            =   2460
               MaxLength       =   250
               TabIndex        =   208
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|prueba_emisor_impulsos||N|"
               Top             =   1785
               Width           =   2865
            End
            Begin VB.TextBox txtaux8 
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
               Left            =   2460
               MaxLength       =   250
               TabIndex        =   207
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|prueba_contador||N|"
               Top             =   1350
               Width           =   2865
            End
            Begin VB.TextBox txtaux8 
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
               Index           =   2
               Left            =   2460
               MaxLength       =   250
               TabIndex        =   206
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|prueba_vavulas_3vias||N|"
               Top             =   915
               Width           =   2865
            End
            Begin VB.TextBox txtaux8 
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
               Index           =   1
               Left            =   2460
               MaxLength       =   10
               TabIndex        =   205
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|dn_valvula_esfera||N|"
               Top             =   555
               Width           =   2865
            End
            Begin VB.TextBox txtaux8 
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
               Left            =   2460
               MaxLength       =   10
               TabIndex        =   204
               Tag             =   "dn_contador|T|S|||rae_visitas_hidtomas|dn_contador||N|"
               Top             =   210
               Width           =   2865
            End
            Begin VB.TextBox txtaux8 
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
               Index           =   6
               Left            =   2460
               MaxLength       =   250
               TabIndex        =   203
               Tag             =   "Hanegadas|N|S|||rae_visitas_hidtomas|hanegadas||N|"
               Top             =   2715
               Width           =   2865
            End
            Begin MSAdodcLib.Adodc AdoAux 
               Height          =   375
               Index           =   3
               Left            =   2550
               Top             =   4110
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
            Begin MSComctlLib.Toolbar Toolbar3 
               Height          =   330
               Index           =   2
               Left            =   6165
               TabIndex        =   201
               Top             =   270
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               Style           =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   1
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Imagen"
                  EndProperty
               EndProperty
            End
            Begin MSComctlLib.Toolbar Toolbar3 
               Height          =   330
               Index           =   3
               Left            =   10740
               TabIndex        =   202
               Top             =   270
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               Style           =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   1
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Imagen"
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label51 
               Caption         =   "Opertividad de la vávula"
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
               Left            =   30
               TabIndex        =   222
               Top             =   2295
               Width           =   2625
            End
            Begin VB.Label Label49 
               Caption         =   "Tapa arqueta"
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
               Left            =   30
               TabIndex        =   221
               Top             =   3195
               Width           =   1815
            End
            Begin VB.Label Label48 
               Caption         =   "Punto de desagüe"
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
               Left            =   30
               TabIndex        =   220
               Top             =   1845
               Width           =   1815
            End
            Begin VB.Label Label47 
               Caption         =   "Tipo de vávula"
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
               Left            =   30
               TabIndex        =   219
               Top             =   1395
               Width           =   1815
            End
            Begin VB.Label Label46 
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
               Height          =   315
               Left            =   30
               TabIndex        =   218
               Top             =   4170
               Width           =   1635
            End
            Begin VB.Label Label44 
               Caption         =   "Tipo de arqueta"
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
               Left            =   30
               TabIndex        =   217
               Top             =   2745
               Width           =   2175
            End
            Begin VB.Image Image9 
               BorderStyle     =   1  'Fixed Single
               Height          =   4035
               Left            =   11730
               Stretch         =   -1  'True
               Top             =   30
               Width           =   3405
            End
            Begin VB.Image Image8 
               BorderStyle     =   1  'Fixed Single
               Height          =   4035
               Left            =   7200
               Stretch         =   -1  'True
               Top             =   30
               Width           =   3405
            End
            Begin VB.Label Label39 
               Caption         =   "INTERIOR"
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
               Left            =   10740
               TabIndex        =   216
               Top             =   60
               Width           =   1035
            End
            Begin VB.Label Label37 
               Caption         =   "EXTERIOR"
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
               Left            =   6165
               TabIndex        =   215
               Top             =   60
               Width           =   1005
            End
            Begin VB.Label Label34 
               Caption         =   "NUDO +"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00972E0B&
               Height          =   315
               Left            =   30
               TabIndex        =   214
               Top             =   915
               Width           =   1875
            End
            Begin VB.Label Label33 
               Caption         =   "SECTOR"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00972E0B&
               Height          =   315
               Left            =   30
               TabIndex        =   213
               Top             =   240
               Width           =   1875
            End
            Begin VB.Label Label32 
               Caption         =   "NUDO -"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00972E0B&
               Height          =   315
               Left            =   30
               TabIndex        =   212
               Top             =   555
               Width           =   1875
            End
         End
         Begin VB.Frame FrameAux6 
            BorderStyle     =   0  'None
            Height          =   4710
            Left            =   -74880
            TabIndex        =   173
            Top             =   420
            Width           =   15450
            Begin VB.TextBox txtaux7 
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
               Height          =   600
               Index           =   10
               Left            =   2640
               MaxLength       =   250
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   184
               Tag             =   "Hanegadas|N|S|||rae_visitas_hidtomas|hanegadas||N|"
               Top             =   4080
               Width           =   12720
            End
            Begin VB.TextBox txtaux7 
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
               Index           =   7
               Left            =   2640
               MaxLength       =   250
               ScrollBars      =   2  'Vertical
               TabIndex        =   183
               Tag             =   "the_geom|T|S|||rae_visitas_hidtomas|the_geom||N|"
               Top             =   2880
               Width           =   2865
            End
            Begin VB.TextBox txtaux7 
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
               Index           =   5
               Left            =   2640
               MaxLength       =   50
               TabIndex        =   182
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|observaciones||N|"
               Top             =   2130
               Width           =   2865
            End
            Begin VB.TextBox txtaux7 
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
               Index           =   4
               Left            =   2640
               MaxLength       =   250
               TabIndex        =   181
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|prueba_emisor_impulsos||N|"
               Top             =   1725
               Width           =   2865
            End
            Begin VB.TextBox txtaux7 
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
               Left            =   2640
               MaxLength       =   250
               TabIndex        =   180
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|prueba_contador||N|"
               Top             =   1350
               Width           =   2865
            End
            Begin VB.TextBox txtaux7 
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
               Index           =   2
               Left            =   2640
               MaxLength       =   250
               TabIndex        =   179
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|prueba_vavulas_3vias||N|"
               Top             =   915
               Width           =   2865
            End
            Begin VB.TextBox txtaux7 
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
               Index           =   1
               Left            =   2640
               MaxLength       =   10
               TabIndex        =   178
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|dn_valvula_esfera||N|"
               Top             =   555
               Width           =   2865
            End
            Begin VB.TextBox txtaux7 
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
               Left            =   2640
               MaxLength       =   10
               TabIndex        =   177
               Tag             =   "dn_contador|T|S|||rae_visitas_hidtomas|dn_contador||N|"
               Top             =   210
               Width           =   2865
            End
            Begin VB.TextBox txtaux7 
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
               Left            =   2640
               MaxLength       =   250
               TabIndex        =   176
               Tag             =   "Hanegadas|N|S|||rae_visitas_hidtomas|hanegadas||N|"
               Top             =   3270
               Width           =   2865
            End
            Begin VB.TextBox txtaux7 
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
               Index           =   6
               Left            =   2640
               MaxLength       =   250
               TabIndex        =   175
               Tag             =   "Hanegadas|N|S|||rae_visitas_hidtomas|hanegadas||N|"
               Top             =   2505
               Width           =   2865
            End
            Begin VB.TextBox txtaux7 
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
               Left            =   2640
               MaxLength       =   250
               TabIndex        =   174
               Tag             =   "Hanegadas|N|S|||rae_visitas_hidtomas|hanegadas||N|"
               Top             =   3705
               Width           =   2865
            End
            Begin MSAdodcLib.Adodc AdoAux 
               Height          =   375
               Index           =   6
               Left            =   2340
               Top             =   4140
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
            Begin MSComctlLib.Toolbar Toolbar3 
               Height          =   330
               Index           =   6
               Left            =   10965
               TabIndex        =   185
               Top             =   270
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               Style           =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   1
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Imagen"
                  EndProperty
               EndProperty
            End
            Begin MSComctlLib.Toolbar Toolbar3 
               Height          =   330
               Index           =   7
               Left            =   6435
               TabIndex        =   186
               Top             =   270
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               Style           =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   1
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Imagen"
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label164 
               Caption         =   "Válvula de aislamiento"
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
               Left            =   30
               TabIndex        =   199
               Top             =   2175
               Width           =   2295
            End
            Begin VB.Label Label165 
               Caption         =   "Tipo Arqueta"
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
               Left            =   30
               TabIndex        =   198
               Top             =   2955
               Width           =   1815
            End
            Begin VB.Label Label168 
               Caption         =   "DN tubería"
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
               Left            =   30
               TabIndex        =   197
               Top             =   1785
               Width           =   1815
            End
            Begin VB.Label Label170 
               Caption         =   "DN Ventosa"
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
               Left            =   30
               TabIndex        =   196
               Top             =   1395
               Width           =   1815
            End
            Begin VB.Label Label176 
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
               Height          =   315
               Left            =   30
               TabIndex        =   195
               Top             =   4170
               Width           =   1635
            End
            Begin VB.Label Label177 
               Caption         =   "Tapa Arqueta"
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
               Left            =   30
               TabIndex        =   194
               Top             =   3360
               Width           =   1635
            End
            Begin VB.Label Label178 
               Caption         =   "Operatividad de la válvula"
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
               Left            =   30
               TabIndex        =   193
               Top             =   2565
               Width           =   2760
            End
            Begin VB.Image Image6 
               BorderStyle     =   1  'Fixed Single
               Height          =   4035
               Left            =   11955
               Stretch         =   -1  'True
               Top             =   30
               Width           =   3405
            End
            Begin VB.Image Image7 
               BorderStyle     =   1  'Fixed Single
               Height          =   4035
               Left            =   7425
               Stretch         =   -1  'True
               Top             =   30
               Width           =   3405
            End
            Begin VB.Label Label24 
               Caption         =   "INTERIOR"
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
               Left            =   10965
               TabIndex        =   192
               Top             =   60
               Width           =   990
            End
            Begin VB.Label Label25 
               Caption         =   "EXTERIOR"
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
               Left            =   6390
               TabIndex        =   191
               Top             =   60
               Width           =   1005
            End
            Begin VB.Label Label27 
               Caption         =   "NUDO +"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00972E0B&
               Height          =   315
               Left            =   30
               TabIndex        =   190
               Top             =   915
               Width           =   1875
            End
            Begin VB.Label Label28 
               Caption         =   "SECTOR"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00972E0B&
               Height          =   315
               Left            =   30
               TabIndex        =   189
               Top             =   240
               Width           =   1875
            End
            Begin VB.Label Label30 
               Caption         =   "NUDO -"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00972E0B&
               Height          =   315
               Left            =   30
               TabIndex        =   188
               Top             =   555
               Width           =   1875
            End
            Begin VB.Label Label31 
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
               Height          =   315
               Left            =   30
               TabIndex        =   187
               Top             =   3750
               Width           =   1635
            End
         End
         Begin VB.Frame FrameAux5 
            BorderStyle     =   0  'None
            Height          =   4680
            Left            =   -74880
            TabIndex        =   144
            Top             =   420
            Width           =   15450
            Begin VB.TextBox txtaux6 
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
               Index           =   10
               Left            =   2730
               MaxLength       =   250
               TabIndex        =   156
               Tag             =   "x|T|S|||rae_visitas_hidtomas|x||N|"
               Top             =   3720
               Width           =   2865
            End
            Begin VB.TextBox txtaux6 
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
               Index           =   7
               Left            =   2730
               MaxLength       =   250
               TabIndex        =   155
               Tag             =   "x|T|S|||rae_visitas_hidtomas|x||N|"
               Top             =   2670
               Width           =   2865
            End
            Begin VB.TextBox txtaux6 
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
               Height          =   555
               Index           =   11
               Left            =   2730
               MaxLength       =   250
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   154
               Tag             =   "Hanegadas|N|S|||rae_visitas_hidtomas|hanegadas||N|"
               Top             =   4080
               Width           =   12405
            End
            Begin VB.TextBox txtaux6 
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
               Left            =   2730
               MaxLength       =   250
               TabIndex        =   153
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|prueba_solenoide||N|"
               Top             =   210
               Width           =   2865
            End
            Begin VB.TextBox txtaux6 
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
               Index           =   1
               Left            =   2730
               MaxLength       =   250
               TabIndex        =   152
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|prueba_contador||N|"
               Top             =   555
               Width           =   2865
            End
            Begin VB.TextBox txtaux6 
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
               Index           =   2
               Left            =   2730
               MaxLength       =   250
               TabIndex        =   151
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|prueba_emisor_impulsos||N|"
               Top             =   915
               Width           =   2865
            End
            Begin VB.TextBox txtaux6 
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
               Left            =   2730
               MaxLength       =   50
               TabIndex        =   150
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|observaciones||N|"
               Top             =   1260
               Width           =   2865
            End
            Begin VB.TextBox txtaux6 
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
               Index           =   4
               Left            =   2730
               MaxLength       =   250
               TabIndex        =   149
               Tag             =   "x|T|S|||rae_visitas_hidtomas|x||N|"
               Top             =   1620
               Width           =   2865
            End
            Begin VB.TextBox txtaux6 
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
               Index           =   5
               Left            =   2730
               MaxLength       =   250
               ScrollBars      =   2  'Vertical
               TabIndex        =   148
               Tag             =   "the_geom|T|S|||rae_visitas_hidtomas|the_geom||N|"
               Top             =   1965
               Width           =   2865
            End
            Begin VB.TextBox txtaux6 
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
               Index           =   6
               Left            =   2730
               MaxLength       =   250
               TabIndex        =   147
               Tag             =   "Parcelas|T|S|||rae_visitas_hidtomas|parcelas||N|"
               Top             =   2310
               Width           =   2865
            End
            Begin VB.TextBox txtaux6 
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
               Left            =   2730
               MaxLength       =   250
               TabIndex        =   146
               Tag             =   "Superficie|N|S|||rae_visitas_hidtomas|superficie||N|"
               Top             =   3375
               Width           =   2865
            End
            Begin VB.TextBox txtaux6 
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
               Left            =   2730
               MaxLength       =   255
               TabIndex        =   145
               Tag             =   "Poligono|N|S|||rae_visitas_hidtomas|poligono||N|"
               Top             =   3015
               Width           =   2865
            End
            Begin MSAdodcLib.Adodc AdoAux 
               Height          =   375
               Index           =   5
               Left            =   2610
               Top             =   120
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
            Begin MSComctlLib.Toolbar Toolbar3 
               Height          =   330
               Index           =   4
               Left            =   10740
               TabIndex        =   157
               Top             =   270
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               Style           =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   1
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Imagen"
                  EndProperty
               EndProperty
            End
            Begin MSComctlLib.Toolbar Toolbar3 
               Height          =   330
               Index           =   5
               Left            =   6165
               TabIndex        =   158
               Top             =   270
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               Style           =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   1
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Imagen"
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label135 
               Caption         =   "Situacion Arqueta"
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
               Left            =   30
               TabIndex        =   172
               Top             =   3720
               Width           =   1875
            End
            Begin VB.Label Label136 
               Caption         =   "EXTERIOR"
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
               Left            =   6165
               TabIndex        =   171
               Top             =   60
               Width           =   975
            End
            Begin VB.Label Label137 
               Caption         =   "Operatividad de la válvula"
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
               Left            =   30
               TabIndex        =   170
               Top             =   2670
               Width           =   2595
            End
            Begin VB.Label Label140 
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
               Height          =   315
               Left            =   30
               TabIndex        =   169
               Top             =   4080
               Width           =   1635
            End
            Begin VB.Label Label147 
               Caption         =   "NUDO -"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00972E0B&
               Height          =   315
               Left            =   30
               TabIndex        =   168
               Top             =   555
               Width           =   1875
            End
            Begin VB.Label Label148 
               Caption         =   "SECTOR"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00972E0B&
               Height          =   315
               Left            =   30
               TabIndex        =   167
               Top             =   240
               Width           =   1875
            End
            Begin VB.Label Label149 
               Caption         =   "NUDO +"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00972E0B&
               Height          =   315
               Left            =   30
               TabIndex        =   166
               Top             =   915
               Width           =   1875
            End
            Begin VB.Label Label150 
               Caption         =   "INTERIOR"
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
               Left            =   10740
               TabIndex        =   165
               Top             =   60
               Width           =   945
            End
            Begin VB.Label Label152 
               Caption         =   "Conexiones"
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
               Left            =   30
               TabIndex        =   164
               Top             =   1965
               Width           =   1875
            End
            Begin VB.Label Label153 
               Caption         =   "DN Válvula"
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
               Left            =   30
               TabIndex        =   163
               Top             =   1260
               Width           =   1875
            End
            Begin VB.Label Label154 
               Caption         =   "Tipo de Válvula"
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
               Left            =   30
               TabIndex        =   162
               Top             =   1620
               Width           =   1875
            End
            Begin VB.Label Label158 
               Caption         =   "Tipologia Arqueta"
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
               Left            =   30
               TabIndex        =   161
               Top             =   3015
               Width           =   1875
            End
            Begin VB.Label Label159 
               Caption         =   "Tapa arqueta"
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
               Left            =   30
               TabIndex        =   160
               Top             =   3375
               Width           =   1875
            End
            Begin VB.Label Label160 
               Caption         =   "Posicion del eje (mariposa)"
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
               Left            =   30
               TabIndex        =   159
               Top             =   2310
               Width           =   2730
            End
            Begin VB.Image Image4 
               BorderStyle     =   1  'Fixed Single
               Height          =   4035
               Left            =   7200
               Stretch         =   -1  'True
               Top             =   30
               Width           =   3405
            End
            Begin VB.Image Image5 
               BorderStyle     =   1  'Fixed Single
               Height          =   4035
               Left            =   11730
               Stretch         =   -1  'True
               Top             =   30
               Width           =   3405
            End
         End
         Begin VB.Frame FrameAux4 
            BorderStyle     =   0  'None
            Height          =   4560
            Left            =   -74880
            TabIndex        =   119
            Top             =   420
            Width           =   15360
            Begin VB.TextBox txtaux5 
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
               Index           =   7
               Left            =   2490
               MaxLength       =   250
               TabIndex        =   129
               Tag             =   "Hanegadas|N|S|||rae_visitas_hidtomas|hanegadas||N|"
               Top             =   2835
               Width           =   3915
            End
            Begin VB.TextBox txtaux5 
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
               Index           =   5
               Left            =   2490
               MaxLength       =   250
               TabIndex        =   128
               Tag             =   "Instaladora|T|S|||rae_visitas_hidtomas|instaladora||N|"
               Top             =   2070
               Width           =   3915
            End
            Begin VB.TextBox txtaux5 
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
               Index           =   4
               Left            =   2490
               MaxLength       =   250
               TabIndex        =   127
               Tag             =   "x|T|S|||rae_visitas_hidtomas|x||N|"
               Top             =   1710
               Width           =   3915
            End
            Begin VB.TextBox txtaux5 
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
               Left            =   2490
               MaxLength       =   250
               TabIndex        =   126
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|prueba_estanqueidad||N|"
               Top             =   1350
               Width           =   3915
            End
            Begin VB.TextBox txtaux5 
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
               Index           =   6
               Left            =   2490
               MaxLength       =   250
               TabIndex        =   125
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|prueba_solenoide||N|"
               Top             =   2445
               Width           =   3915
            End
            Begin VB.TextBox txtaux5 
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
               Left            =   2490
               MaxLength       =   10
               TabIndex        =   124
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|dn_toma||N|"
               Top             =   150
               Width           =   1350
            End
            Begin VB.TextBox txtaux5 
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
               Index           =   2
               Left            =   2490
               MaxLength       =   10
               TabIndex        =   123
               Tag             =   "dn_contador|T|S|||rae_visitas_hidtomas|dn_contador||N|"
               Top             =   990
               Width           =   1350
            End
            Begin VB.TextBox txtaux5 
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
               Index           =   1
               Left            =   2490
               MaxLength       =   40
               TabIndex        =   122
               Tag             =   "Falta Bypass|T|N|||rae_visitas_hidtomas|falta_bypass||S|"
               Top             =   630
               Width           =   1350
            End
            Begin VB.TextBox txtaux5 
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
               Left            =   2490
               MaxLength       =   250
               TabIndex        =   121
               Tag             =   "Hanegadas|N|S|||rae_visitas_hidtomas|hanegadas||N|"
               Top             =   3195
               Width           =   3915
            End
            Begin VB.TextBox txtaux5 
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
               Height          =   1005
               Index           =   9
               Left            =   2490
               MaxLength       =   250
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   120
               Tag             =   "x|T|S|||rae_visitas_hidtomas|x||N|"
               Top             =   3645
               Width           =   3945
            End
            Begin MSAdodcLib.Adodc AdoAux 
               Height          =   375
               Index           =   4
               Left            =   2610
               Top             =   120
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
            Begin MSComctlLib.Toolbar Toolbar3 
               Height          =   330
               Index           =   0
               Left            =   8430
               TabIndex        =   130
               Top             =   30
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               Style           =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   1
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Imagen"
                  EndProperty
               EndProperty
            End
            Begin MSComctlLib.Toolbar Toolbar3 
               Height          =   330
               Index           =   1
               Left            =   12555
               TabIndex        =   131
               Top             =   30
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               Style           =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   1
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Imagen"
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label87 
               Caption         =   "Tch Fijacion Colector"
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
               Left            =   30
               TabIndex        =   143
               Top             =   2475
               Width           =   2265
            End
            Begin VB.Label Label88 
               Caption         =   "Caja Empalmes Correcta"
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
               Left            =   30
               TabIndex        =   142
               Top             =   2865
               Width           =   2520
            End
            Begin VB.Label Label89 
               Caption         =   "Tch Tipo"
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
               Left            =   30
               TabIndex        =   141
               Top             =   2130
               Width           =   1755
            End
            Begin VB.Label Label95 
               Caption         =   "INTERIOR"
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
               Left            =   11505
               TabIndex        =   140
               Top             =   180
               Width           =   1125
            End
            Begin VB.Label Label96 
               Caption         =   "EXTERIOR"
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
               Left            =   7290
               TabIndex        =   139
               Top             =   150
               Width           =   1185
            End
            Begin VB.Label Label99 
               Caption         =   "Estado Colector"
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
               Left            =   30
               TabIndex        =   138
               Top             =   1770
               Width           =   1680
            End
            Begin VB.Label Label100 
               Caption         =   "HIDRANTE"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00972E0B&
               Height          =   255
               Left            =   30
               TabIndex        =   137
               Top             =   240
               Width           =   1035
            End
            Begin VB.Label Label101 
               Caption         =   "Válvula Compuerta"
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
               Left            =   30
               TabIndex        =   136
               Top             =   1380
               Width           =   2625
            End
            Begin VB.Label Label102 
               Caption         =   "Fecha 1ª revisión"
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
               Left            =   30
               TabIndex        =   135
               Top             =   990
               Width           =   1860
            End
            Begin VB.Label Label103 
               Caption         =   "Constructora"
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
               Left            =   30
               TabIndex        =   134
               Top             =   630
               Width           =   1575
            End
            Begin VB.Label Label105 
               Caption         =   "Nivelacion Arquetar"
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
               Left            =   30
               TabIndex        =   133
               Top             =   3255
               Width           =   2175
            End
            Begin VB.Label Label119 
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
               Height          =   315
               Left            =   30
               TabIndex        =   132
               Top             =   3675
               Width           =   1575
            End
            Begin VB.Image Image2 
               BorderStyle     =   1  'Fixed Single
               Height          =   4035
               Left            =   7275
               Stretch         =   -1  'True
               Top             =   480
               Width           =   3405
            End
            Begin VB.Image Image3 
               BorderStyle     =   1  'Fixed Single
               Height          =   4035
               Left            =   11475
               Stretch         =   -1  'True
               Top             =   480
               Width           =   3405
            End
         End
         Begin VB.Frame FrameAux2 
            BorderStyle     =   0  'None
            Height          =   4410
            Left            =   -74880
            TabIndex        =   69
            Top             =   420
            Width           =   15405
            Begin VB.TextBox txtaux1 
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
               Index           =   2
               Left            =   1995
               MaxLength       =   255
               TabIndex        =   118
               Top             =   1230
               Width           =   3375
            End
            Begin VB.TextBox txtaux1 
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
               Left            =   1995
               MaxLength       =   5
               TabIndex        =   115
               Tag             =   "Poligono|N|S|||rae_visitas_hidtomas|poligono||N|"
               Top             =   1680
               Width           =   3375
            End
            Begin VB.TextBox txtaux1 
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
               Index           =   11
               Left            =   8520
               MaxLength       =   20
               TabIndex        =   114
               Tag             =   "Fecha Entrada|F|S|||rae_visitas_hidtomas|fecha_entrada||N|"
               Top             =   1313
               Width           =   1605
            End
            Begin VB.TextBox txtaux1 
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
               Left            =   8520
               MaxLength       =   10
               TabIndex        =   111
               Tag             =   "Toma|T|S|||rae_visitas_hidtomas|toma||N|"
               Top             =   531
               Width           =   1575
            End
            Begin MSComctlLib.Toolbar Toolbar2 
               Height          =   330
               Left            =   13470
               TabIndex        =   102
               Top             =   30
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               Style           =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   1
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Sigpac"
                  EndProperty
               EndProperty
            End
            Begin VB.TextBox txtaux1 
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
               Left            =   1995
               MaxLength       =   40
               TabIndex        =   85
               Tag             =   "Contador|T|N|||rae_visitas_hidtomas|contador||S|"
               Top             =   270
               Width           =   1245
            End
            Begin VB.TextBox txtaux1 
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
               Index           =   6
               Left            =   1995
               MaxLength       =   10
               TabIndex        =   84
               Tag             =   "dn_contador|T|S|||rae_visitas_hidtomas|dn_contador||N|"
               Top             =   2880
               Width           =   3345
            End
            Begin VB.TextBox txtaux1 
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
               Index           =   7
               Left            =   1995
               MaxLength       =   125
               TabIndex        =   83
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|dn_valvula_esfera||N|"
               Top             =   3270
               Width           =   3315
            End
            Begin VB.TextBox txtaux1 
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
               Left            =   1995
               MaxLength       =   125
               TabIndex        =   82
               Tag             =   "dn_toma|T|S|||rae_visitas_hidtomas|dn_toma||N|"
               Top             =   3660
               Width           =   3285
            End
            Begin VB.TextBox txtaux1 
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
               Index           =   4
               Left            =   1995
               MaxLength       =   125
               TabIndex        =   81
               Tag             =   "Parcelas|T|S|||rae_visitas_hidtomas|parcelas||N|"
               Top             =   2070
               Width           =   3375
            End
            Begin VB.TextBox txtaux1 
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
               Index           =   12
               Left            =   8520
               MaxLength       =   20
               TabIndex        =   80
               Tag             =   "Instaladora|T|S|||rae_visitas_hidtomas|instaladora||N|"
               Top             =   1704
               Width           =   1605
            End
            Begin VB.TextBox txtaux1 
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
               Index           =   5
               Left            =   1995
               MaxLength       =   20
               TabIndex        =   79
               Tag             =   "Hanegadas|N|S|||rae_visitas_hidtomas|hanegadas||N|"
               Top             =   2430
               Width           =   1275
            End
            Begin VB.TextBox txtaux1 
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
               Index           =   10
               Left            =   8520
               MaxLength       =   20
               TabIndex        =   78
               Tag             =   "Recibido|T|S|||rae_visitas_hidtomas|Recibido||N|"
               Top             =   922
               Width           =   1605
            End
            Begin VB.TextBox txtaux1 
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
               Left            =   8520
               MaxLength       =   20
               TabIndex        =   77
               Tag             =   "Fecha Revision|F|S|||rae_visitas_hidtomas|fecha_revision||N|"
               Top             =   2095
               Width           =   1605
            End
            Begin VB.TextBox txtaux1 
               BackColor       =   &H80000018&
               Height          =   1275
               Index           =   18
               Left            =   10260
               MaxLength       =   255
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   76
               Tag             =   "dn_valvula|T|S|||rae_visitas_hidtomas|observaciones_RAE||N|"
               Top             =   510
               Width           =   4665
            End
            Begin VB.TextBox txtaux1 
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
               Left            =   8520
               MaxLength       =   255
               TabIndex        =   75
               Tag             =   "Fecha Turno|T|S|||rae_visitas_hidtomas|fecha_turno||N|"
               Top             =   2486
               Width           =   1635
            End
            Begin VB.TextBox txtaux1 
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
               Left            =   8520
               MaxLength       =   255
               TabIndex        =   74
               Tag             =   "Verificacion|T|S|||rae_visitas_hidtomas|verificacion||N|"
               Top             =   3660
               Width           =   1635
            End
            Begin VB.TextBox txtaux1 
               BackColor       =   &H80000018&
               Height          =   1275
               Index           =   19
               Left            =   10290
               MaxLength       =   255
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   73
               Tag             =   "Comentarios Indefa|T|S|||rae_visitas_hidtomas|Comentarios_INDEFA||N|"
               Top             =   2610
               Width           =   4665
            End
            Begin VB.TextBox txtaux1 
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
               Left            =   8520
               MaxLength       =   255
               TabIndex        =   72
               Tag             =   "Turno|T|S|||rae_visitas_hidtomas|turno||N|"
               Top             =   2877
               Width           =   1635
            End
            Begin VB.TextBox txtaux1 
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
               Left            =   8520
               MaxLength       =   255
               TabIndex        =   71
               Tag             =   "q_instantaneo|T|S|||rae_visitas_hidtomas|q_instantaneo||N|"
               Top             =   3268
               Width           =   1635
            End
            Begin VB.TextBox txtaux1 
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
               Index           =   1
               Left            =   1995
               MaxLength       =   255
               TabIndex        =   70
               Tag             =   "Nro socio|N|S|||rae_visitas_hidtomas|socio_revisado||N|"
               Top             =   825
               Width           =   1245
            End
            Begin MSAdodcLib.Adodc AdoAux 
               Height          =   375
               Index           =   2
               Left            =   6360
               Top             =   -90
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
            Begin VB.Label Label65 
               Caption         =   "Polígono"
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
               Left            =   30
               TabIndex        =   117
               Top             =   1665
               Width           =   1995
            End
            Begin VB.Label Label40 
               Caption         =   "F.entrada ficha"
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
               Left            =   5925
               TabIndex        =   116
               Top             =   1305
               Width           =   1665
            End
            Begin VB.Label Label186 
               Caption         =   "DN Contador"
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
               Left            =   30
               TabIndex        =   113
               Top             =   2865
               Width           =   1530
            End
            Begin VB.Label Label185 
               Caption         =   "Toma"
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
               Left            =   5925
               TabIndex        =   112
               Top             =   540
               Width           =   1035
            End
            Begin VB.Label Label17 
               Caption         =   "CONTADOR"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00972E0B&
               Height          =   255
               Left            =   30
               TabIndex        =   101
               Top             =   240
               Width           =   1035
            End
            Begin VB.Label Label19 
               Caption         =   "Nombre Socio"
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
               Left            =   30
               TabIndex        =   100
               Top             =   1215
               Width           =   1440
            End
            Begin VB.Label Label20 
               Caption         =   "DN Valvula"
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
               Left            =   30
               TabIndex        =   99
               Top             =   3270
               Width           =   1395
            End
            Begin VB.Label Label21 
               Caption         =   "DN Toma"
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
               Left            =   30
               TabIndex        =   98
               Top             =   3660
               Width           =   1485
            End
            Begin VB.Label Label22 
               Caption         =   "Nº Socio"
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
               Left            =   30
               TabIndex        =   97
               Top             =   840
               Width           =   1815
            End
            Begin VB.Label Label35 
               Caption         =   "Parcela/s"
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
               Left            =   30
               TabIndex        =   96
               Top             =   2055
               Width           =   1155
            End
            Begin VB.Label Label36 
               Caption         =   "Instaladora"
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
               Left            =   5925
               TabIndex        =   95
               Top             =   1695
               Width           =   2100
            End
            Begin VB.Label Label38 
               Caption         =   "Superficie total(hg)"
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
               Left            =   30
               TabIndex        =   94
               Top             =   2445
               Width           =   1995
            End
            Begin VB.Label Label42 
               Caption         =   "Alta Recibida"
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
               Left            =   5955
               TabIndex        =   93
               Top             =   915
               Width           =   1440
            End
            Begin VB.Label Label43 
               Caption         =   "Fecha revisión instalación"
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
               Left            =   5925
               TabIndex        =   92
               Top             =   2085
               Width           =   2700
            End
            Begin VB.Label Label45 
               Caption         =   "Acciones Requeridas RAE"
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
               Left            =   10260
               TabIndex        =   91
               Top             =   210
               Width           =   5115
            End
            Begin VB.Label Label52 
               Caption         =   "Observaciones INDEFA"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   10305
               TabIndex        =   90
               Top             =   2295
               Width           =   2490
            End
            Begin VB.Label Label54 
               Caption         =   "Turno Asignado"
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
               Left            =   5925
               TabIndex        =   89
               Top             =   2865
               Width           =   1815
            End
            Begin VB.Label Label57 
               Caption         =   "Caudal instantáneo"
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
               Left            =   5925
               TabIndex        =   88
               Top             =   3270
               Width           =   2085
            End
            Begin VB.Label Label63 
               Caption         =   "Fecha puesta en turno"
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
               Left            =   5925
               TabIndex        =   87
               Top             =   2475
               Width           =   2535
            End
            Begin VB.Label Label64 
               Caption         =   "Tipo verificación (C/T)"
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
               Left            =   5925
               TabIndex        =   86
               Top             =   3660
               Width           =   2355
            End
         End
      End
      Begin VB.Frame FrameAux1 
         BorderStyle     =   0  'None
         Height          =   4350
         Left            =   -74910
         TabIndex        =   38
         Top             =   450
         Width           =   7780
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
            Height          =   290
            Index           =   5
            Left            =   6570
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   49
            Text            =   "Par"
            Top             =   2940
            Visible         =   0   'False
            Width           =   435
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
            Height          =   290
            Index           =   4
            Left            =   6180
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   48
            Text            =   "Pol"
            Top             =   2940
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.TextBox txtAux2 
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
            Height          =   290
            Index           =   3
            Left            =   5700
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   47
            Text            =   "Hdas"
            Top             =   2940
            Visible         =   0   'False
            Width           =   435
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
            Height          =   290
            Index           =   2
            Left            =   4350
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   46
            Text            =   "Poblacion"
            Top             =   2940
            Visible         =   0   'False
            Width           =   1305
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
            Height          =   290
            Index           =   1
            Left            =   2565
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   43
            Text            =   "Partida"
            Top             =   2925
            Visible         =   0   'False
            Width           =   1695
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
            Height          =   315
            Index           =   1
            Left            =   2385
            TabIndex        =   42
            ToolTipText     =   "Buscar campo"
            Top             =   2925
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtaux4 
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
            Left            =   1710
            MaxLength       =   8
            TabIndex        =   41
            Tag             =   "Campo|N|N|||rpozos_campos|codcampo|00000000|N|"
            Text            =   "campo"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux4 
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
            Left            =   225
            MaxLength       =   10
            TabIndex        =   40
            Tag             =   "Hidrante|T|N|||rpozos_campos|hidrante||S|"
            Text            =   "Hidrante"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux4 
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
            Left            =   945
            MaxLength       =   6
            TabIndex        =   39
            Tag             =   "Linea|N|N|||rpozos_campos|numlinea|000|N|"
            Text            =   "linea"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   1
            Left            =   45
            TabIndex        =   44
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
            Left            =   4590
            Top             =   180
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
            Bindings        =   "frmPOZHidrantesIndefa.frx":0318
            Height          =   3810
            Index           =   1
            Left            =   30
            TabIndex        =   45
            Top             =   480
            Width           =   7660
            _ExtentX        =   13520
            _ExtentY        =   6720
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
      Begin VB.Frame FrameAux0 
         BorderStyle     =   0  'None
         Height          =   3930
         Left            =   -74910
         TabIndex        =   28
         Top             =   450
         Width           =   7780
         Begin VB.TextBox txtaux3 
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
            Index           =   3
            Left            =   5940
            MaxLength       =   6
            TabIndex        =   34
            Tag             =   "Porcentaje|N|N|0|100|rpozos_cooprop|porcentaje|##0.00||"
            Text            =   "porc"
            Top             =   2940
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux3 
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
            Left            =   945
            MaxLength       =   6
            TabIndex        =   33
            Tag             =   "Linea|N|N|||rpozos_cooprop|numlinea|000|S|"
            Text            =   "linea"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux3 
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
            Left            =   225
            MaxLength       =   10
            TabIndex        =   32
            Tag             =   "Hidrante|T|N|||rpozos_cooprop|hidrante||S|"
            Text            =   "Hidrante"
            Top             =   2925
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux3 
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
            Left            =   1710
            MaxLength       =   6
            TabIndex        =   31
            Tag             =   "Socio|N|N|||rpozos_cooprop|codsocio|000000|N|"
            Text            =   "socio"
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
            Height          =   315
            Index           =   0
            Left            =   2385
            TabIndex        =   30
            ToolTipText     =   "Buscar socio"
            Top             =   2925
            Visible         =   0   'False
            Width           =   195
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
            Height          =   315
            Index           =   0
            Left            =   2565
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   29
            Text            =   "Nombre socio"
            Top             =   2925
            Visible         =   0   'False
            Width           =   3285
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   0
            Left            =   45
            TabIndex        =   35
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
            Left            =   4590
            Top             =   180
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
            Bindings        =   "frmPOZHidrantesIndefa.frx":0330
            Height          =   3195
            Index           =   0
            Left            =   30
            TabIndex        =   36
            Top             =   450
            Width           =   7450
            _ExtentX        =   13150
            _ExtentY        =   5636
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
      Height          =   285
      Index           =   0
      Left            =   1710
      MaxLength       =   40
      TabIndex        =   110
      Top             =   990
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   15990
      TabIndex        =   230
      Top             =   120
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
   Begin VB.Image Image1 
      Height          =   525
      Left            =   7170
      Top             =   4770
      Width           =   1245
   End
   Begin VB.Label Label50 
      Caption         =   "Buscando diferencias"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   3510
      TabIndex        =   67
      Top             =   9885
      Visible         =   0   'False
      Width           =   4215
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
      Begin VB.Menu mnDiferencias 
         Caption         =   "Buscar &Diferencias"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnActualizar 
         Caption         =   "Actualizar Registro"
         Shortcut        =   ^A
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
Attribute VB_Name = "frmPOZHidrantesIndefa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MONICA                   -+-+
' +-+- Menú: Hidrantes de Pozos        -+-+
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
Private WithEvents frmCam As frmManCampos 'campos
Attribute frmCam.VB_VarHelpID = -1
Private WithEvents frmCam1 As frmManCampos 'campos
Attribute frmCam1.VB_VarHelpID = -1
Private WithEvents frmMen2 As frmMensajes ' orden de printnou
Attribute frmMen2.VB_VarHelpID = -1
Private WithEvents frmMen3 As frmMensajes ' busqueda de diferencias
Attribute frmMen3.VB_VarHelpID = -1
Private frmMensImg As frmMensajes ' visualizacion de la imagen

Private WithEvents frmHidPrev As frmBasico2 ' contadores vista previa
Attribute frmHidPrev.VB_VarHelpID = -1

' *****************************************************
Dim CodTipoMov As String

Dim Orden As String

Dim ConexionIndefa As Boolean
Dim Continuar As Boolean

Dim SocioAnt As String

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

Dim ListOrigen As Collection
Dim ListDestino As Collection



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


Dim MostradoAviso As Boolean

Private Sub cmdAceptar_Click()
Dim Diferencias As String
Dim i As Integer

    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BÚSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm2(Me, 2, "Frame2") Then
                
                    '------------------------------------------------------------------------------
                    '  LOG de acciones
                    Set LOG = New cLOG
                    LOG.Insertar 10, vUsu, "Nuevo contador: " & Text1(0).Text
                    Set LOG = Nothing
                    '-----------------------------------------------------------------------------
                
                    ImprimirComunicacionIndefa True

                    ' *** canviar o llevar el WHERE, repasar codEmpre ****
                    Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
                    'Data1.RecordSource = "Select * from " & NombreTabla & " where codempre = " & codEmpre & Ordenacion
                    ' ***************************************************************
                    PosicionarData
                End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario2(Me, 2, "Frame2") Then
                
                    Diferencias = ""
                    For i = 0 To Text1.Count - 1
                        If Text1(i).Text <> ListOrigen.item(i + 1) Then
                            Diferencias = Diferencias & Mid(Text1(i).Tag, 1, 8) & ":" & ListOrigen.item(i + 1) & "-" & Text1(i).Text & "·"
                        End If
                    Next i
                    Set ListOrigen = Nothing
                    
                    '------------------------------------------------------------------------------
                    '  LOG de acciones
                    Set LOG = New cLOG
                    LOG.Insertar 11, vUsu, "Contador: " & Text1(0).Text & " " & Diferencias
                    Set LOG = Nothing
                    '-----------------------------------------------------------------------------
                    
                    ImprimirComunicacionIndefa False
                
                
                    TerminaBloquear
                    PosicionarData
                End If
            Else
                ModoLineas = 0
            End If
        
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                    InsertarLinea
                Case 2 'modificar llínies
                    ModificarLinea
                    PosicionarData
                    SumaTotalPorcentajes NumTabMto
            End Select
        
    End Select
    Screen.MousePointer = vbDefault
    
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub ImprimirComunicacionIndefa(esAlta As Boolean)
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Contador As String
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal

    If Text1(0).Text = "" Or Len(Text1(0).Text) <> 6 Then Exit Sub
    
    Contador = Text1(0).Text
    
    If esAlta Then
        If MsgBox(" Se ha dado de alta un nuevo contador." & vbCrLf & vbCrLf & "¿ Desea imprimir un informe de comunicación a Indefa ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            
            indRPT = 88 ' informe de comunicacion de cambios a indefa
            
            If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                       
            InicializarVbles
            
            'Añadir el parametro de Empresa
            cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
            numParam = numParam + 1
            cadParam = cadParam & "pContador=""" & Text1(0).Text & """|"
            numParam = numParam + 1
            cadParam = cadParam & "pHanegadas=" & DBSet(Text1(6).Text, "N") & "|"
            numParam = numParam + 1
            cadParam = cadParam & "pPoligono=" & DBSet(Text1(4).Text, "T") & "|"
            numParam = numParam + 1
            cadParam = cadParam & "pParcela=" & DBSet(Text1(5).Text, "T") & "|"
            numParam = numParam + 1
            cadParam = cadParam & "pSocio=" & Text1(2).Text & "|"
            numParam = numParam + 1
            cadParam = cadParam & "pToma=" & CLng(ComprobarCero(Text1(1).Text)) Mod 100 & "|"
            numParam = numParam + 1
            

            cadTitulo = "Carta de Comunicación a Indefa"
            cadNombreRPT = nomDocu
            cadFormula = "{rpozos.hidrante} = " & DBSet(Contador, "T")
            
            LlamarImprimir
        
        End If
        
        Exit Sub
    End If
    
    
    
    If Not ConexionIndefa Then Exit Sub
    
    
    Sql = "select poligono, parcelas, hanegadas, toma, socio_revisado from rae_visitas_hidtomas where sector = " & DBSet(CInt(Mid(Contador, 1, 2)), "T")
    Sql = Sql & " and hidrante = " & DBSet(CInt(Mid(Contador, 3, 2)), "T")
    '[Monica]18/07/2013:cambio
                                    '[Monica]27/01/2014: lo cambio a numerico
    Sql = Sql & " and salida_tch = " & DBSet(CInt(Mid(Contador, 5, 2)), "N")

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, ConnIndefa, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
      '[Monica]19/11/2013: si no hay socio revisado no hacemos nada
      If DBLet(Rs!socio_revisado, "N") <> 0 Then
        
        If (DBLet(Rs!Poligono, "T") <> Text1(4).Text) Or (Mid(DBLet(Rs!parcelas, "T"), 1, 25) <> Mid(Text1(5).Text, 1, 25)) Or (Int(ComprobarCero(DBLet(Rs!Hanegadas, "N"))) <> Int(ComprobarCero(Text1(6).Text))) Or CInt(DBLet(Rs!socio_revisado, "N") <> Text1(2).Text And DBLet(Rs!socio_revisado, "N") <> 0) Or _
           (CLng(ComprobarCero(DBLet(Rs!toma, "N"))) <> CLng(ComprobarCero(Text1(1).Text)) Mod 100) Then
            If MsgBox("Se han producido diferencias con los datos de Indefa." & vbCrLf & vbCrLf & " ¿ Desea imprimir un informe de comunicación ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                
                indRPT = 88 ' informe de comunicacion de cambios a indefa
                
                If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                           
                InicializarVbles
                
                'Añadir el parametro de Empresa
                cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
                numParam = numParam + 1
                cadParam = cadParam & "pIndPol=""" & DBLet(Rs!Poligono, "T") & """|"
                numParam = numParam + 1
                cadParam = cadParam & "pIndPar=""" & DBLet(Rs!parcelas, "T") & """|"
                numParam = numParam + 1
                cadParam = cadParam & "pIndHda=" & DBSet(Rs!Hanegadas, "N") & "|"
                numParam = numParam + 1
                cadParam = cadParam & "pIndToma=" & CLng(ComprobarCero(DBSet(Rs!toma, "N"))) & "|"
                numParam = numParam + 1
                cadParam = cadParam & "pContador=""" & Text1(0).Text & """|"
                numParam = numParam + 1
                cadParam = cadParam & "pSocioAnt=" & SocioAnt & "|"
                numParam = numParam + 1
                cadParam = cadParam & "pHanegadas=" & DBSet(Text1(6).Text, "N") & "|"
                numParam = numParam + 1
                cadParam = cadParam & "pPoligono=" & DBSet(Text1(4).Text, "T") & "|"
                numParam = numParam + 1
                cadParam = cadParam & "pParcela=" & DBSet(Text1(5).Text, "T") & "|"
                numParam = numParam + 1
                cadParam = cadParam & "pSocio=" & Text1(2).Text & "|"
                numParam = numParam + 1
                cadParam = cadParam & "pToma=" & CLng(ComprobarCero(Text1(1).Text)) Mod 100 & "|"
                numParam = numParam + 1
                

                cadTitulo = "Carta de Comunicación a Indefa"
                cadNombreRPT = nomDocu
                cadFormula = "{rpozos.hidrante} = " & DBSet(Contador, "T")
                
                LlamarImprimir
                
            End If
        End If
      End If
    End If
        
End Sub


Private Sub cmdAux_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 0 ' Socios coopropietarios
            Set frmSoc1 = New frmManSocios
            frmSoc1.DatosADevolverBusqueda = "0|1|"
            frmSoc1.Show vbModal
            Set frmSoc1 = Nothing
            PonerFoco txtAux3(2)
            
        Case 1 ' campos
            Set frmCam1 = New frmManCampos
            frmCam1.DatosADevolverBusqueda = "0|"
'            frmCam.CodigoActual = Text1(5).Text
            frmCam1.Show vbModal
            Set frmCam1 = Nothing
            PonerFoco txtAux4(2)
        
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1

End Sub


Private Sub Form_Activate()
Dim Rs As ADODB.Recordset
Dim Sql As String

    Screen.MousePointer = vbDefault
    If PrimeraVez Then
        PrimeraVez = False
        DoEvents
        If Not Continuar Then Unload Me
        
        If Not MostradoAviso Then
            If ConexionIndefa Then
'                SQL = "select * from rae_visitas_hidtomas "
'
'                Set RS = New ADODB.Recordset
'                RS.Open SQL, ConnIndefa, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                If Not RS.EOF Then
'                    If RS.Fields.Count > 52 Then
'                        If MsgBox("Han cambiado la estructura de Contadores Indefa, hay datos que no se van a mostrar." & vbCrLf & "¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
'                            Set RS = Nothing
'                            Exit Sub
'                        End If
'                    End If
'                End If
'
'                SQL = "select * from rae_visitas_hidrantes "
'                Set RS = New ADODB.Recordset
'                RS.Open SQL, ConnIndefa, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                If Not RS.EOF Then
'                    If RS.Fields.Count > 38 Then
'                        If MsgBox("Han cambiado la estructura de Hidrantes Indefa, hay datos que no se van a mostrar." & vbCrLf & "¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
'                            Set RS = Nothing
'                            Exit Sub
'                        End If
'                    End If
'                End If
'
'                SQL = "select * from rae_visitas_desagues "
'                Set RS = New ADODB.Recordset
'                RS.Open SQL, ConnIndefa, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                If Not RS.EOF Then
'                    If RS.Fields.Count > 20 Then
'                        If MsgBox("Han cambiado la estructura de Desagües Indefa, hay datos que no se van a mostrar." & vbCrLf & "¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
'                            Set RS = Nothing
'                            Exit Sub
'                        End If
'                    End If
'                End If
'
            End If
        End If
        
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
    Screen.MousePointer = vbDefault
    CerrarConexionIndefa
End Sub

Private Sub Form_Load()
Dim i As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    Continuar = True
    
    '[Monica]08/05/2012: añadida Escalona que funciona como Utxera
    Me.Caption = "Contadores"
    Me.Label1(0).Caption = "Contador TCH"
    
'[Monica]22/02/2019: quito lo de la conexion a indefa
'    ConexionIndefa = False
'    If AbrirConexionIndefa() = False Then
'        If MsgBox("No se ha podido acceder a los datos de Indefa. " & vbCrLf & "¿ Desea Continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
'            Continuar = False
'            Exit Sub
'        End If
'    Else
'        ConexionIndefa = True
'    End If
    
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
        'el 10  son separadors
        .Buttons(8).Image = 10  'Imprimir
    End With
    
    With Me.Toolbar5
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 21  'Buscar diferencias
        .Buttons(2).Image = 26  'Actualizar desde datos de indefa
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
    
    'cargar IMAGES de busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    For i = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(i).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next i


'    Me.cmdSigpac.Picture = frmPpal.imgListComun16.ListImages(29).Picture
    With Me.Toolbar2
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 29   'Insertar
    End With
    ' *** si n'hi han tabs, per a que per defecte sempre es pose al 1r***
    Me.SSTab1.Tab = 0
    ' *******************************************************************
    
    '[Monica]08/02/2013: cargamos todos los toolbar de camara de fotos
    For i = 0 To 7
        With Me.Toolbar3(i)
            .HotImageList = frmPpal.imgListComun_OM
            .DisabledImageList = frmPpal.imgListComun_BN
            .ImageList = frmPpal.imgListComun
            .Buttons(1).Image = 36   'camara
        End With
    Next i
    
    LimpiarCampos   'Neteja els camps TextBox
    
    CodTipoMov = "NOC"

    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "rpozos"
    Ordenacion = " ORDER BY hidrante "
    '************************************************
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    '***** canviar el nom de la PK de la capçalera; repasar codEmpre *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where hidrante is null"
    Data1.Refresh
       
    ModoLineas = 0
    
    Me.SSTab2.Tab = 0
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1 'búsqueda
        ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
        Text1(0).BackColor = vbLightBlue 'codclien
        ' ****************************************************************************
    End If
End Sub


Private Sub LimpiarCampos()
Dim i As Integer

    On Error Resume Next
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
'    Me.chkAbonos(0).Value = 0
    
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
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    DesplazamientoVisible b And Data1.Recordset.RecordCount > 1
    
    '---------------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    CmdAceptar.visible = b
       
    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a més neteja els camps Text1
    
' cambio la siguiente expresion por la de abajo
'   BloquearText1 Me, Modo
    For i = 0 To Text1.Count - 1
        BloquearTxt Text1(i), Not (Modo = 1 Or Modo = 3 Or Modo = 4)
    Next i
    BloquearCombo Me, Modo
    
    b = (Modo <> 1)
    BloquearTxt Text1(19), b
    
    FrameAux2.Enabled = (Modo = 2)
    FrameAux3.Enabled = (Modo = 2)
    FrameAux4.Enabled = (Modo = 2)
    FrameAux5.Enabled = (Modo = 2)
    FrameAux6.Enabled = (Modo = 2)
    For i = 0 To txtAux1.Count - 1
        txtAux1(i).Locked = True
    Next i
    For i = 0 To txtAux5.Count - 1
        txtAux5(i).Locked = True
    Next i
    For i = 0 To txtaux6.Count - 1
        txtaux6(i).Locked = True
    Next i
    For i = 0 To txtaux7.Count - 1
        txtaux7(i).Locked = True
    Next i
    'Campos Nº entrada bloqueado y en azul
'    BloquearTxt Text1(0), Modo = 4
    
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
'    b = (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) And (NumTabMto = 0))
    ' ****************************************************************************
            
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    PonerLongCampos

    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 0, False
        CargaGrid 1, False
    End If
    
    b = (Modo = 4) Or (Modo = 2)
    

    DataGridAux(0).Enabled = b
    
    Me.Toolbar2.Enabled = (Modo = 2)
    Me.Toolbar2.visible = (Modo = 2)
    ' las camaras
    For i = 0 To 7
        Me.Toolbar3(i).Enabled = (Modo = 2)
        Me.Toolbar3(i).visible = (Modo = 2)
    Next i
        
'        Me.Toolbar3.Enabled = (Modo = 2)
'        Me.Toolbar3.visible = (Modo = 2)
'        Me.Toolbar4.Enabled = (Modo = 2)
'        Me.Toolbar4.visible = (Modo = 2)

'     '-----------------------------
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
    'Buscar diferencias con indefa
    Toolbar5.Buttons(1).Enabled = b And ConexionIndefa
    Me.mnDiferencias.Enabled = b And ConexionIndefa
    
    'Insertar
    Toolbar1.Buttons(1).Enabled = b
    Me.mnNuevo.Enabled = b
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) 'And Not DeConsulta
    
    'actualizar registro con datos de indefa         '[Monica]27/01/2014: solo para el caso de escalona
    Toolbar5.Buttons(2).Enabled = b And ConexionIndefa And (vParamAplic.Cooperativa = 10)
    Me.mnActualizar.Enabled = b And ConexionIndefa And (vParamAplic.Cooperativa = 10)
    
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b
    'Imprimir
    Toolbar1.Buttons(8).Enabled = b
       
       
       
       
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
    b = (Modo = 3 Or Modo = 4 Or Modo = 2)
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.Adoaux(i).Recordset.RecordCount > 0)
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i
    ' ****************************************
       
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botons de Desplaçament; per a desplaçar-se pels registres de control Data
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
'    indice = CByte(Me.cmdAux(0).Tag + 2)
'    txtaux1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmC1_Selec(vFecha As Date)
Dim indice As Byte
    indice = CByte(Me.imgFec(0).Tag)
    Text1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmCam_DatoSeleccionado(CadenaSeleccion As String)
    Text1(18).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo de campo
    FormateaCampo Text1(18)
    PonerDatosCampo Text1(18).Text
End Sub

Private Sub frmCam1_DatoSeleccionado(CadenaSeleccion As String)
    txtAux4(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codcampo
    FormateaCampo txtAux4(2)
    PonerDatosCampoLineas txtAux4(2)

End Sub


Private Sub frmHidPrev_DatoSeleccionado(CadenaSeleccion As String)
Dim cadB As String
    
    If CadenaSeleccion <> "" Then
        cadB = "hidrante = " & DBSet(RecuperaValor(CadenaSeleccion, 1), "T")
        
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub frmMen2_DatoSeleccionado(CadenaSeleccion As String)
    Orden = CadenaSeleccion
    If CadenaSeleccion = "" Then Orden = "pOrden={rpozos.hidrante}"
End Sub

Private Sub frmMen3_DatoSeleccionado(CadenaSeleccion As String)
    cadB = ""
    If CadenaSeleccion <> "" Then
        cadB = "hidrante in (" & Mid(CadenaSeleccion, 1, Len(CadenaSeleccion) - 1) & ")"
        HacerBusquedaDiferencias
    End If
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    Text1(18).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo de campo
    FormateaCampo Text1(18)
    PonerDatosCampo Text1(18).Text
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

Private Sub frmSoc1_DatoSeleccionado(CadenaSeleccion As String)
    txtAux3(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codsocio
    FormateaCampo txtAux3(2)
    txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'nomsocio
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
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
            Case 0
                indice = 8
            Case 1
                indice = 10
            Case 2
                indice = 16
            Case 3
                indice = 17
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
            indice = 11
            frmZ.pTitulo = "Observaciones del Hidrante"
            frmZ.pValor = Text1(indice).Text
            frmZ.pModo = Modo
            frmZ.Show vbModal
            Set frmZ = Nothing
            PonerFoco Text1(indice)
    End Select
End Sub

Private Sub mnActualizar_Click()
    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************
    If BLOQUEADesdeFormulario2(Me, Data1, 2, "Frame2") Then BotonActualizar
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub


Private Sub mnDiferencias_Click()
    BotonBuscarDiferencias
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
    
    If BLOQUEADesdeFormulario2(Me, Data1, 2, "Frame2") Then BotonModificar
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


Private Sub SSTab2_Click(PreviousTab As Integer)
'    PonerCampos
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 1  'Nou
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
        Case 5  'Búscar
            mnBuscar_Click
        Case 6  'Tots
            mnVerTodos_Click
        Case 8 'Imprimir
            mnImprimir_Click
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

Private Sub BotonActualizar()
Dim i As Integer
Dim Sql As String
Dim Contador As String
Dim Rs As ADODB.Recordset
Dim cadena As String

Dim Orden As Long

    On Error GoTo eBotonActualizar

    Contador = Text1(0).Text
    
    If Contador = "" Or Len(Contador) < 6 Then Exit Sub
    
    Sql = "select poligono, parcelas, hanegadas, socio_revisado, toma from rae_visitas_hidtomas where sector = " & DBSet(CInt(Mid(Contador, 1, 2)), "T")
    Sql = Sql & " and hidrante = " & DBSet(CInt(Mid(Contador, 3, 2)), "T")
    '[Monica]18/07/2013:cambio
                                    '[Monica]27/01/2014: lo cambio a numerico
    Sql = Sql & " and salida_tch = " & DBSet(CInt(Mid(Contador, 5, 2)), "N")

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, ConnIndefa, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
      '[Monica]19/11/2013: si el socio_revisado esta vacio no hacemos ninguna comprobacion
      If DBLet(Rs!socio_revisado, "N") <> 0 Then
      
        '[Monica]07/11/2013: puede que indefa nos haya metido un socio que no existe
        If CInt(DBLet(Rs!socio_revisado, "N") <> Text1(2).Text And DBLet(Rs!socio_revisado, "N") <> 0) Then
            Sql = "select * from rsocios where codsocio = " & DBSet(Rs!socio_revisado, "N")
            If TotalRegistros(Sql) = 0 Then
                MsgBox "Debe dar de alta el socio " & DBLet(Rs!socio_revisado, "N"), vbExclamation
                Exit Sub
            End If
        End If
    
        If (DBLet(Rs!Poligono, "T") <> Text1(4).Text) Or (Mid(DBLet(Rs!parcelas, "T"), 1, 25) <> Mid(Text1(5).Text, 1, 25)) Or (Int(ComprobarCero(DBLet(Rs!Hanegadas, "N"))) <> Int(ComprobarCero(Text1(6).Text))) Or CInt(DBLet(Rs!socio_revisado, "N") <> Text1(2).Text And DBLet(Rs!socio_revisado, "N") <> 0) Or _
           (CLng(ComprobarCero(DBLet(Rs!toma, "N"))) <> CLng(ComprobarCero(Text1(1).Text)) Mod 100) Then
            
            cadena = ""
            If (DBLet(Rs!Poligono, "T") <> Text1(4).Text) Then
                cadena = cadena & " Pol:" & Trim(Text1(4).Text) & "-" & DBLet(Rs!Poligono, "T") & "·"
            End If
            If (Mid(DBLet(Rs!parcelas, "T"), 1, 25) <> Mid(Text1(5).Text, 1, 25)) Then
                cadena = cadena & "Par:" & Trim(Text1(5).Text) & "-" & (Mid(DBLet(Rs!parcelas, "T"), 1, 25)) & "·"
            End If
            If (Int(ComprobarCero(DBLet(Rs!Hanegadas, "N"))) <> Int(ComprobarCero(Text1(6).Text))) Then
                cadena = cadena & "Hdas:" & Int(ComprobarCero(Text1(6).Text)) & "-" & Int(ComprobarCero(DBLet(Rs!Hanegadas, "N"))) & "·"
            End If
            If CLng((DBLet(Rs!socio_revisado, "N")) <> CLng(Text1(2).Text) And DBLet(Rs!socio_revisado, "N") <> 0) Then
                cadena = cadena & "Soc:" & Trim(Text1(2).Text) & "-" & DBLet(Rs!socio_revisado, "N") & "·"
            End If
            If (CLng(ComprobarCero(DBLet(Rs!toma, "N"))) <> CLng(ComprobarCero(Text1(1).Text)) Mod 100) Then
                cadena = cadena & "Toma:" & CLng(Text1(1).Text) Mod 100 & "-" & CLng(ComprobarCero(DBLet(Rs!toma, "N"))) & "·"
            End If
            

            '------------------------------------------------------------------------------
            '  LOG de acciones
            Set LOG = New cLOG
            LOG.Insertar 11, vUsu, "Contador:" & Contador & vbCrLf & " " & cadena
            Set LOG = Nothing
            '-----------------------------------------------------------------------------
            
            Sql = "update rpozos set poligono = " & DBSet(Rs!Poligono, "T")
            Sql = Sql & ", parcelas = " & DBSet(Mid(Rs!parcelas, 1, 25), "T")
            Sql = Sql & ", hanegada = " & DBSet(Rs!Hanegadas, "N")
            '[Monica]23/10/2013: daba error cuando no me han insertado el socio
            If DBLet(Rs!socio_revisado, "N") <> 0 Then
                Sql = Sql & ", codsocio = " & DBSet(Rs!socio_revisado, "N")
            End If
            '[Monica]30/10/2013: hemos de actualizar tambien el nro de orden con la toma de indefa
            Orden = (CLng(Text1(1).Text) \ 100) * 100 + CLng(ComprobarCero(DBLet(Rs!toma, "N")))
            Sql = Sql & ", nroorden = " & DBSet(Orden, "N")
            
            Sql = Sql & " where hidrante = " & DBSet(Contador, "T")
            
            conn.Execute Sql
            TerminaBloquear
            Data1.Refresh
    '        PonerCampos
            SituarData Data1, "hidrante = " & DBSet(Contador, "T"), Me.lblIndicador, False
            PonerCampos
        
            MsgBox "Proceso realizado correctamente.", vbExclamation
            
        End If
      End If
    End If
    Exit Sub
    
eBotonActualizar:
    MuestraError Err.Number, "Actualizar Datos Indefa", Err.Description
End Sub

Private Sub BotonBuscarDiferencias()
Dim i As Integer

    LimpiarCampos
    
    
    Set frmMen3 = New frmMensajes
    frmMen3.OpcionMensaje = 44
    frmMen3.Show vbModal
    
    Set frmMen3 = Nothing
    
    
' ******************************************************************************
End Sub

Private Sub HacerBusquedaDiferencias()
    
    If chkVistaPrevia(0) = 1 Then
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

Private Sub HacerBusqueda()

    '[Monica]09/01/2014: nuevo tipo para que no lleve los asteriscos implicitos
    Text1(4).Tag = "Polígono|TT|S|||rpozos|poligono||N|"
    
    cadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
    Text1(4).Tag = "Polígono|T|S|||rpozos|poligono||N|"
    
    If chkVistaPrevia(0) = 1 Then
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

    Set frmHidPrev = New frmBasico2
    
    AyudaPOZHidrantesPrev frmHidPrev, , cadB
    
    Set frmHidPrev = Nothing

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
'    Text1(0).Text = SugerirCodigoSiguienteStr("rentradas", "numnotac")
'    FormateaCampo Text1(0)
       
    PonerFoco Text1(0) '*** 1r camp visible que siga PK ***
    ' ***********************************************************
    ' *** si n'hi han camps de descripció a la capçalera ***
    PosarDescripcions
    ' ******************************************************

    Text1(7).Text = "0"
    Text1(9).Text = "0"
    
    ' *** si n'hi han tabs, em posicione al 1r ***
    Me.SSTab1.Tab = 0
    ' ********************************************
End Sub


Private Sub BotonModificar()
Dim i As Integer

    PonerModo 4
    
    SocioAnt = Text1(2).Text


    Set ListOrigen = New Collection

    For i = 0 To Text1.Count - 1
        ListOrigen.Add Text1(i).Text
    Next i


    ' *** bloquejar els camps visibles de la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    ' *************************************************************************
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(1)
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
    Cad = "¿Seguro que desea eliminar el Hidrante?"
    Cad = Cad & vbCrLf & "Hidrante: " & Data1.Recordset.Fields(0)
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
    PonerCamposForma2 Me, Data1, 2, "Frame2"   'opcio=1: posa el format o els camps de la capçalera
    
    ' *** si n'hi han llínies en datagrids ***
    For i = 0 To DataGridAux.Count - 1
        CargaGrid i, True
        If Not Adoaux(i).Recordset.EOF Then _
            PonerCamposForma2 Me, Adoaux(i), 2, "FrameAux" & i
    Next i
    ' *******************************************
    SumaTotalPorcentajes 0
    
    PosarDescripcions
    
    '[Monica]15/05/2013: Visualizamos los cobros pendientes del socio
    ComprobarCobrosSocio CStr(Data1.Recordset!Codsocio), ""
    
    If ConexionIndefa Then
        PosarDescripcionsIndefa
        PosarDescripcionsIndefa2
        PosarDescripcionsIndefa3
        PosarDescripcionsIndefa4
        PosarDescripcionsIndefa5
    End If

' lo he quitado de aqui pq el consumo está almacenado en un campo de la tabla rpozos
'    CalcularConsumo
    ' ********************************************************************************
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu   'Activar opcions de menú según nivell
                        'de permisos de l'usuari
    
End Sub


Private Sub PosarDescripcionsIndefa()
Dim NomEmple As String
Dim CodPobla As String
Dim Sql As String
Dim Contador As String
Dim Rs As ADODB.Recordset
Dim i As Integer

    On Error GoTo EPosarDescripcions


    ' Limpiamos los campos de indefa
    For i = 0 To 19
        txtAux1(i).Text = ""
    Next i
    
    Contador = Text1(0).Text
    
    If Contador = "" Or Len(Contador) < 6 Or Not IsNumeric(Contador) Then Exit Sub
    '[Monica]19/02/2014: en la columna de nombre_socio(cc1) antes estaba ""codigotch"", no viene en la base de datos nombre_socio de rae_visitas_hidtomas
    Sql = "select ""Contador"" as contador1, socio_revisado, '' as ccc1, poligono, parcelas, hanegadas, dn_contador, dn_valvula_esfera, dn_toma, codigotoma,"
    Sql = Sql & """Recibido"" as aaa , fecha_entrada, instaladora, fecha_revision, fecha_turno,turno, q_instantaneo,verificacion,""observaciones_RAE"" as ccc,"
    Sql = Sql & """Comentarios_INDEFA"" as hhh"
    Sql = Sql & " from rae_visitas_hidtomas where sector = " & DBSet(Int(Mid(Contador, 1, 2)), "T")
    Sql = Sql & " and hidrante = " & DBSet(Int(Mid(Contador, 3, 2)), "T")
    
    '[Monica]18/07/2013: cambio
                                    '[Monica]27/01/2014: lo cambio a numerico
    Sql = Sql & " and salida_tch = " & DBSet(Int(Mid(Contador, 5, 2)), "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, ConnIndefa, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        For i = 0 To 19
            txtAux1(i).Text = DBLet(Rs.Fields(i).Value)
        Next i
    End If
    If txtAux1(5).Text <> "" Then txtAux1(5).Text = Format(txtAux1(5).Text, "###,##0.00")
    Set Rs = Nothing
    
EPosarDescripcions:
    If Err.Number <> 0 Then MsgBox "Han cambiado datos de Contadores de Indefa", vbExclamation
End Sub
' ************************************************************

Private Sub PosarDescripcionsIndefa2()
Dim NomEmple As String
Dim CodPobla As String
Dim Sql As String
Dim Contador As String
Dim Rs As ADODB.Recordset
Dim i As Integer

    On Error GoTo EPosarDescripcions


    ' Limpiamos los campos de indefa
    For i = 0 To txtaux8.Count - 1
        txtaux8(i).Text = ""
    Next i
    
    Contador = Text1(0).Text
    
    If Contador = "" Or Len(Contador) < 6 Or Not IsNumeric(Contador) Then Exit Sub
    
    Sql = "select sector, hidrante1, hidrante2, valvula_aislamiento, punto_entrega, fto_valvula, tipo_arqueta, tipo_tapa, observaciones, foto_desague, foto_general "
    Sql = Sql & " from rae_visitas_desagues "
    Sql = Sql & " where sector = " & DBSet(Int(Mid(Contador, 1, 2)), "T")
    Sql = Sql & " and hidrante1 = " & DBSet(Int(Mid(Contador, 3, 2)), "T")

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, ConnIndefa, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Me.Toolbar3(2).Tag = ""
    Me.Toolbar3(3).Tag = ""
    
    If Not Rs.EOF Then
        For i = 0 To txtaux8.Count - 1
            txtaux8(i).Text = DBLet(Rs.Fields(i).Value)
        Next i
    
        If Dir(App.Path & "\FotosHts\" & Rs!foto_desague) <> "" Then
            Me.Toolbar3(2).Tag = App.Path & "\FotosHts\" & Rs!foto_desague
            Image8.Picture = LoadPicture(Me.Toolbar3(2).Tag)
            
        End If
        
        If Dir(App.Path & "\FotosHts\" & Rs!foto_general) <> "" Then
            Me.Toolbar3(3).Tag = App.Path & "\FotosHts\" & Rs!foto_general
            Image9.Picture = LoadPicture(Me.Toolbar3(3).Tag)
        End If
    End If
    Set Rs = Nothing
    
    
EPosarDescripcions:
    If Err.Number <> 0 Then MsgBox "Han cambiado datos de Desagües de Indefa", vbExclamation
End Sub
' ************************************************************



Private Sub PosarDescripcionsIndefa3()
Dim NomEmple As String
Dim CodPobla As String
Dim Sql As String
Dim Contador As String
Dim Rs As ADODB.Recordset
Dim i As Integer
Dim vTabAnt As Integer

    On Error GoTo EPosarDescripcions



    ' Limpiamos los campos de indefa
    For i = 0 To txtAux5.Count - 1
        txtAux5(i).Text = ""
    Next i
    
    Contador = Text1(0).Text
    
    If Contador = "" Or Len(Contador) < 6 Or Not IsNumeric(Contador) Then Exit Sub
    
    Sql = "select codigo_hidrante, constructora, fecha, valvula_compuerta, estado_colector, tch_tipo, tch_fijacion_colector, caja_empalmes_tch, nivelacion_arqueta_verticalidad, observaciones,  "
    Sql = Sql & " foto1, foto2"
    Sql = Sql & " from rae_visitas_hidrantes "
    Sql = Sql & " where sector = " & DBSet(Int(Mid(Contador, 1, 2)), "T")
    Sql = Sql & " and hidrante = " & DBSet(Int(Mid(Contador, 3, 2)), "T")
    

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, ConnIndefa, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Me.Toolbar3(0).Tag = ""
    Me.Toolbar3(1).Tag = ""
    Image2.Picture = LoadPicture("")
    Image3.Picture = LoadPicture("")
    
    If Not Rs.EOF Then
    
        For i = 0 To txtAux5.Count - 1
            txtAux5(i).Text = DBLet(Rs.Fields(i).Value)
        Next i
        If Dir(App.Path & "\FotosHts\" & Rs!foto1 & ".jpg") <> "" Then
            Me.Toolbar3(0).Tag = App.Path & "\FotosHts\" & Rs!foto2 & ".jpg"
            Image2.Picture = LoadPicture(Me.Toolbar3(0).Tag)
        End If
        
        If Dir(App.Path & "\FotosHts\" & Rs!foto2 & ".jpg") <> "" Then
            Me.Toolbar3(1).Tag = App.Path & "\FotosHts\" & Rs!foto1 & ".jpg"
            Image3.Picture = LoadPicture(Me.Toolbar3(1).Tag)
        End If
    End If
    DoEvents

'    Me.Toolbar3(0).Buttons(1).Enabled = (Me.Toolbar3(0).Tag <> "")
'    Me.Toolbar3(1).Buttons(1).Enabled = (Me.Toolbar3(1).Tag <> "")
    Set Rs = Nothing

EPosarDescripcions:
    If Err.Number <> 0 Then MsgBox "Han cambiado datos de Hidrantes de Indefa", vbExclamation
End Sub
' ************************************************************


Private Sub PosarDescripcionsIndefa4()
Dim NomEmple As String
Dim CodPobla As String
Dim Sql As String
Dim Contador As String
Dim Rs As ADODB.Recordset
Dim i As Integer

    On Error GoTo EPosarDescripcions


    ' Limpiamos los campos de indefa
    For i = 0 To txtaux6.Count - 1
        txtaux6(i).Text = ""
    Next i
    
    Contador = Text1(0).Text
    
    If Contador = "" Or Len(Contador) < 6 Or Not IsNumeric(Contador) Then Exit Sub
    
    Sql = "select sector, hidrante1, hidrante2, dn_tuberia_instalada, valvula_mariposa, uniones, eje_en_valvula_de_mariposa, comprobacion,"
    Sql = Sql & " tipologia_arqueta, tapa_fundicion_en_conos, situacion_tapa_arqueta, observaciones, foto2, foto_valvulas_aislamiento"
    Sql = Sql & " from rae_visitas_valvulas "
    Sql = Sql & " where sector = " & DBSet(Int(Mid(Contador, 1, 2)), "T")
    Sql = Sql & " and hidrante1 = " & DBSet(Int(Mid(Contador, 3, 2)), "T")

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, ConnIndefa, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    Me.Toolbar3(4).Tag = ""
    Me.Toolbar3(5).Tag = ""
    Me.Image4.Picture = LoadPicture("")
    Me.Image5.Picture = LoadPicture("")
    
    If Not Rs.EOF Then
        For i = 0 To txtaux6.Count - 1
            txtaux6(i).Text = DBLet(Rs.Fields(i).Value)
        Next i
    
        If InStr(1, Rs!foto_valvulas_aislamiento, "http") <> 0 Then
            Me.Toolbar3(4).Tag = Rs!foto_valvulas_aislamiento
'            Image5.Picture = LoadPicture(Me.Toolbar3(4).Tag)
        Else
            If DBLet(Rs!foto_valvulas_aislamiento, "T") <> "" Then
                If Dir(App.Path & "\FotosHts\" & Rs!foto_valvulas_aislamiento) <> "" Then
                    Me.Toolbar3(4).Tag = App.Path & "\FotosHts\" & Rs!foto_valvulas_aislamiento
                    Image5.Picture = LoadPicture(Me.Toolbar3(4).Tag)
                End If
            End If
        End If
        If InStr(1, Rs!foto2, "http") <> 0 Then
            Me.Toolbar3(5).Tag = Rs!foto2
'            Image4.Picture = LoadPicture(Me.Toolbar3(5).Tag)
        Else
            If DBLet(Rs!foto2, "T") <> "" Then
                If Dir(App.Path & "\FotosHts\" & Rs!foto2) <> "" Then
                    Me.Toolbar3(5).Tag = App.Path & "\FotosHts\" & Rs!foto2
                    Image4.Picture = LoadPicture(Me.Toolbar3(5).Tag)
                End If
            End If
        End If
    End If
    
    Set Rs = Nothing
    
    
EPosarDescripcions:
    If Err.Number <> 0 Then MsgBox "Han cambiado datos de Válvulas de Indefa", vbExclamation
End Sub
' ************************************************************



Private Sub PosarDescripcionsIndefa5()
Dim NomEmple As String
Dim CodPobla As String
Dim Sql As String
Dim Contador As String
Dim Rs As ADODB.Recordset
Dim i As Integer

    On Error GoTo EPosarDescripcions


    ' Limpiamos los campos de indefa
    For i = 0 To txtaux7.Count - 1
        txtaux7(i).Text = ""
    Next i
    
    Contador = Text1(0).Text
    
    If Contador = "" Or Len(Contador) < 6 Or Not IsNumeric(Contador) Then Exit Sub
    
    Sql = "select sector, hidrante1, hidrante2, diametro_ventosa, diametro_tuberia_plano, aislamiento, comprobacion, tipologia_arqueta, tapa_arqueta, situacion, observaciones, foto2, foto_ventosa"
    Sql = Sql & " from rae_visitas_ventosas "
    Sql = Sql & " where sector = " & DBSet(Int(Mid(Contador, 1, 2)), "T")
    Sql = Sql & " and hidrante1 = " & DBSet(Int(Mid(Contador, 3, 2)), "T")

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, ConnIndefa, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Me.Toolbar3(6).Tag = ""
    Me.Toolbar3(7).Tag = ""
    
    If Not Rs.EOF Then
        For i = 0 To txtaux7.Count - 1
            txtaux7(i).Text = DBLet(Rs.Fields(i).Value)
        Next i
        
        If InStr(1, Rs!foto_ventosa, "http") <> 0 Then
            Me.Toolbar3(6).Tag = Rs!foto_ventosa
'            Image6.Picture = LoadPicture(Me.Toolbar3(6).Tag)
        Else
            If DBLet(Rs!foto_ventosa, "T") <> "" Then
                If Dir(App.Path & "\FotosHts\" & Rs!foto_ventosa) <> "" Then
                    Me.Toolbar3(6).Tag = App.Path & "\FotosHts\" & Rs!foto_ventosa
                    Image6.Picture = LoadPicture(Me.Toolbar3(6).Tag)
                End If
            End If
        End If
        If InStr(1, Rs!foto2, "http") <> 0 Then
            Me.Toolbar3(7).Tag = Rs!foto2
'            Image7.Picture = LoadPicture(Me.Toolbar3(7).Tag)
        Else
            If DBLet(Rs!foto2, "T") <> "" Then
                If Dir(App.Path & "\FotosHts\" & Rs!foto2 & ".jpg") <> "" Then
                    Me.Toolbar3(7).Tag = App.Path & "\FotosHts\" & Rs!foto2
                    Image7.Picture = LoadPicture(Me.Toolbar3(7).Tag)
                End If
            End If
        End If
    End If
    
    
    Set Rs = Nothing
    
    
EPosarDescripcions:
    If Err.Number <> 0 Then MsgBox "Han cambiado datos de Ventosas de Indefa", vbExclamation
End Sub
' ************************************************************











Private Sub CalcularConsumo()
Dim Sql As String
Dim Inicio As Long
Dim Fin As Long
Dim Consumo As Long
Dim NroDig As Integer
Dim Limite As Long

    If Text1(9).Text = "" Then Exit Sub

    Inicio = 0
    Fin = 0
    
    If Text1(7).Text <> "" Then Inicio = CLng(Text1(7).Text)
    If Text1(9).Text <> "" Then Fin = CLng(Text1(9).Text)
    
    NroDig = CCur(Text1(12).Text)
    Limite = (10 ^ NroDig)
    
    If Fin >= Inicio Then
        Consumo = Fin - Inicio
    Else
        If MsgBox("¿ Es un reinicio de contador ? ", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
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
    '[Monica]11/06/2013: cambio el formato del consumo
    Text1(19).Text = Format(Consumo, "########0")

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
                    If NumTabMto = 0 Or NumTabMto = 1 Or NumTabMto = 2 Or NumTabMto = 3 Then
                        DataGridAux(NumTabMto).AllowAddNew = False
                        ' **** repasar si es diu Data1 l'adodc de la capçalera ***
                        'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
                        ' ********************************************************
                        LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                        DataGridAux(NumTabMto).Enabled = True
                        DataGridAux(NumTabMto).SetFocus

                    ' *** els tabs que no tenen datagrid ***
                    ElseIf NumTabMto = 3 Then
                        If BLOQUEADesdeFormulario2(Me, Data1, 2, "Frame2") Then BotonModificar 'Modificar
'                        CargaFrame 3, True
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
            
            SumaTotalPorcentajes NumTabMto

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
Dim b As Boolean
Dim Sql As String
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOk = False
    
    b = CompForm2(Me, 2, "Frame2")
    If Not b Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
        'comprobar si existe ya el cod. del campo clave primaria
        If ExisteCP(Text1(0)) Then b = False
    End If
    
    If b And (Modo = 3 Or Modo = 4) Then
        If b Then
'            If Not EstaSocioDeAlta(Text1(2).Text) Then
'            ' comprobamos que el socio no este dado de baja
'                SQL = "El socio introducido está dado de baja. Reintroduzca. " & vbCrLf & vbCrLf
'                MsgBox SQL, vbExclamation
'                b = False
'                PonerFoco Text1(2)
'            End If

'            If Text2(0).Text = "" Then Text2(0).Text = "0"
'            Text1(19).Text = Text2(0).Text
'            If CCur(Text2(0).Text) < 0 Then
'                MsgBox "El consumo no puede ser negativo. Revise.", vbExclamation
'                PonerFoco Text1(9)
'                b = False
'            End If
        
            CalcularConsumo
            If CCur(ComprobarCero(Text1(19).Text)) < 0 Then
                MsgBox "El consumo no puede ser negativo. Revise.", vbExclamation
                PonerFoco Text1(9)
                b = False
            End If
        End If
        
        '[Monica]04/02/2013: si el campo introducido no existe daba un error sin controlar
        If b Then
            If Text1(18).Text <> "" Then
                Sql = "select count(*) from rcampos where codcampo = " & DBSet(Text1(18).Text, "N")
                If TotalRegistros(Sql) = 0 Then
                    MsgBox "No existe el campo. Reintroduzca.", vbExclamation
                    PonerFoco Text1(18)
                    b = False
                End If
            End If
        End If
        
        If b Then
            If Text1(18).Text <> "" And Text1(2).Text <> "" Then
                Sql = "select count(*) from rcampos where codcampo = " & DBSet(Text1(18).Text, "N")
                Sql = Sql & " and codsocio = " & DBSet(Text1(2).Text, "N")
                If TotalRegistros(Sql) = 0 Then
                    If MsgBox("El campo introducido no es del socio. ¿ Desea continuar ? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                        PonerFoco Text1(18)
                        b = False
                    End If
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
    Cad = "(hidrante='" & Text1(0).Text & "')"
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


    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    Select Case Index
        Case 1 'nro de orden
            PonerFormatoEntero Text1(1)

        Case 2 'SOCIO
            If Modo = 1 Then Exit Sub
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
                    If Modo = 3 Then
                        If TotalRegistros("select count(*) from rcampos where codsocio = " & DBSet(Text1(Index).Text, "N")) > 0 Then
                            Set frmMens = New frmMensajes
                            frmMens.cadWhere = "and rcampos.codsocio = " & DBSet(Text1(Index).Text, "N")
                            frmMens.campo = Text1(Index).Text
                            frmMens.OpcionMensaje = 29
                            frmMens.Show vbModal
                            Set frmMens = Nothing
                        End If
                    End If
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
                '[Monica]01/10/2014: en el caso de que exista la partida hemos de refrescar el nombre de la poblacion
                Else
                    
                    Text2(1).Text = DevuelveValor("select despobla from rpartida inner join rpueblos on rpartida.codpobla = rpueblos.codpobla where rpartida.codparti = " & DBSet(Text1(3).Text, "N"))
                
                End If
            Else
                Text2(Index).Text = ""
            End If
                
        Case 6 'hanegadas
            PonerFormatoDecimal Text1(Index), 7
                
        Case 7, 9 ' CONTADORES
            If Modo = 1 Then Exit Sub
            PonerFormatoEntero Text1(Index)
            CalcularConsumo
                
        Case 8, 10 'Fecha no comprobaremos que esté dentro de campaña
                    'Fecha de alta y fecha de baja
            PonerFormatoFecha Text1(Index), True
            
        Case 14 ' calibre
            PonerFormatoEntero Text1(Index)
            
        Case 15 ' acciones
            PonerFormatoEntero Text1(Index)
            
        Case 16, 17 ' fecha de alta y baja del contador
            '[Monica]28/08/2013: comprobamos que la fecha esté en la campaña
            PonerFormatoFecha Text1(Index)
            
        Case 13 'Tipo de pozo
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rtipopozos", "nompozo", "codpozo", "N")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Pozo: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmPoz = New frmPOZPozos
                        frmPoz.DatosADevolverBusqueda = "0|1|"
                        frmPoz.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmPoz.Show vbModal
                        Set frmPoz = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        
        Case 18
            If PonerFormatoEntero(Text1(Index)) Then
                PonerDatosCampo (Text1(Index))
            End If
                        
    End Select
End Sub

Private Sub PonerDatosCampo(campo As String)
Dim Cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim Rs As ADODB.Recordset
Dim Sql As String


    If campo = "" Then Exit Sub
    
    If Not (Modo = 3 Or Modo = 4) Then Exit Sub

'[Monica]08/02/2013: quito esto pq si quieren traer los datos del campo desplegaran la lupa
'    '[Monica]22/11/2012: Preguntamos si quiere traer los datos del socio del campo
'    If (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10) And Modo = 4 Then
'        Sql = "select rcampos.codsocio, rsocios.nomsocio from rcampos inner join rsocios on rcampos.codsocio = rsocios.codsocio where rcampos.codcampo = " & DBSet(Text1(18).Text, "N")
'
'        Set RS = New ADODB.Recordset
'        RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'        If DBLet(RS.Fields(0)) <> CLng(ComprobarCero(Text1(2).Text)) And Modo = 3 Then
'            Text1(2).Text = Format(DBLet(RS!CodSocio, "N"), "000000") ' codigo de socio del campo
'            Text2(2).Text = DBLet(RS!nomsocio, "T") ' nombre de socio
'
'           'If MsgBox("¿ Desea traer los datos de RAE al contador ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
'        End If
'
'        Set RS = Nothing
'
'        Exit Sub
'
'    End If

    Cad = "rcampos.codcampo = " & DBSet(campo, "N") & " and rcampos.fecbajas is null"
     
    Cad1 = "select rcampos.codparti, rpartida.nomparti, rcampos.poligono, rcampos.parcela, rcampos.supcoope, rpueblos.despobla, rcampos.subparce, rcampos.codsocio "
    Cad1 = Cad1 & " from rcampos, rpartida, rpueblos "
    Cad1 = Cad1 & " where " & Cad
    Cad1 = Cad1 & " and rcampos.codparti = rpartida.codparti "
    Cad1 = Cad1 & " and rpartida.codpobla = rpueblos.codpobla"
     
    Set Rs = New ADODB.Recordset
    Rs.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Text1(3).Text = ""
        Text1(4).Text = ""
        Text1(5).Text = ""
        Text1(6).Text = ""
        Text2(1).Text = ""
        
        Text1(18).Text = campo
        PonerFormatoEntero Text1(18)
        Text1(3).Text = DBLet(Rs.Fields(0).Value, "N") ' codigo de partida
        If Text1(3).Text <> "" Then Text1(3).Text = Format(Text1(3).Text, "0000")
        Text2(3).Text = DBLet(Rs.Fields(1).Value, "T") ' nombre de partida
        Text2(1).Text = DBLet(Rs.Fields(5).Value, "T") ' nombre de poblacion
        Text1(4).Text = DBLet(Rs.Fields(2).Value, "N") ' poligono
'[Monica]03/08/2012: quito el formato de poligono y parcela
'        If Text1(4).Text <> "" Then Text1(4).Text = Format(Text1(4).Text, "0000")
        Text1(5).Text = DBLet(Rs.Fields(3).Value, "N") ' parcela
        
        If vParamAplic.Cooperativa = 10 Then Text1(5).Text = Text1(5).Text & " " & DBLet(Rs.Fields(6).Value)
        
'        If Text1(5).Text <> "" Then Text1(5).Text = Format(Text1(5).Text, "000000")
        
        'hanegadas
        Text1(6).Text = Format(Round2(DBLet(Rs.Fields(4).Value, "N") / vParamAplic.Faneca, 2), "##,##0.00")
    End If
    
    Set Rs = Nothing
    
End Sub

Private Sub PonerDatosCampoLineas(campo As String)
Dim Cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim Rs As ADODB.Recordset
Dim i As Integer


    If campo = "" Then Exit Sub

    Cad = "rcampos.codcampo = " & DBSet(campo, "N") & " and rcampos.fecbajas is null"
     
    Cad1 = "select rcampos.codparti, rpartida.nomparti, rcampos.poligono, rcampos.parcela, rcampos.supcoope, rpueblos.despobla "
    Cad1 = Cad1 & " from rcampos, rpartida, rpueblos "
    Cad1 = Cad1 & " where " & Cad
    Cad1 = Cad1 & " and rcampos.codparti = rpartida.codparti "
    Cad1 = Cad1 & " and rpartida.codpobla = rpueblos.codpobla"
     
    Set Rs = New ADODB.Recordset
    Rs.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        For i = 1 To 5
            txtAux2(i).Text = ""
        Next i
        
        txtAux4(2).Text = campo
        PonerFormatoEntero txtAux4(2)
        txtAux2(1).Text = DBLet(Rs.Fields(1).Value, "T") ' nombre de partida
        txtAux2(2).Text = DBLet(Rs.Fields(5).Value, "T") ' nombre de poblacion
        txtAux2(4).Text = DBLet(Rs.Fields(2).Value, "N") ' poligono
        If txtAux2(4).Text <> "" Then txtAux2(4).Text = Format(txtAux2(4).Text, "0000")
        txtAux2(5).Text = DBLet(Rs.Fields(3).Value, "N") ' parcela
        If txtAux2(5).Text <> "" Then txtAux2(5).Text = Format(txtAux2(5).Text, "000000")
        
        'hanegadas
        txtAux2(3).Text = Format(Round2(DBLet(Rs.Fields(4).Value, "N") / vParamAplic.Faneca, 2), "##,##0.00")
    End If
    
    Set Rs = Nothing
    
End Sub


Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 2: KEYBusqueda KeyAscii, 0 'socio
                Case 3: KEYBusqueda KeyAscii, 1 'partida
                Case 8: KEYFecha KeyAscii, 0 'fecha desde
                Case 10: KEYFecha KeyAscii, 1 'fecha desde
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


Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

' **** si n'hi han camps de descripció a la capçalera ****
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



' *** si n'hi han formularis de buscar codi a les llínies ***
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
            Set frmCam = New frmManCampos
            frmCam.DatosADevolverBusqueda = "0|1|"
'            frmCam.CodigoActual = Text1(18).Text
            frmCam.Show vbModal
            Set frmCam = Nothing
            PonerFoco Text1(18)
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 2, "Frame2"
End Sub


' ***** si n'hi han varios nivells de tabs *****
Private Sub SituarTab(numTab As Integer)
    On Error Resume Next
    
    If numTab = 0 Then
        SSTab1.Tab = 1
    ElseIf numTab = 1 Then
        SSTab1.Tab = 2
    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub
' **********************************************


Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " codsocio=" & Val(Text1(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function


Private Sub printNou()
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim Sql As String
Dim ConFechaBaja As Boolean

    ' pedimos el orden del informe
    Set frmMen2 = New frmMensajes
    
    frmMen2.OpcionMensaje = 38
    frmMen2.Show vbModal
    
    Set frmMen2 = Nothing
    
    ConFechaBaja = False
    If MsgBox("¿ Desea imprimir los contadores con fecha de baja ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        ConFechaBaja = True
    End If
    
    indRPT = 78 ' personalizacion del informe de hidrantes
    
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
    
    
    With frmImprimir2
        .cadTabla2 = "rpozos"
        .Informe2 = nomDocu
        If cadB <> "" Then
            If InStr(cadB, "in (") <> 0 Then
                .cadRegSelec = ""
            
                Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
                conn.Execute Sql
            
                Sql = "insert into tmpinformes (codusu, nombre1) select " & vUsu.Codigo & ", " & Replace(Data1.RecordSource, "select * ", "hidrante ")
                conn.Execute Sql
            
                .Informe2 = Replace(nomDocu, ".rpt", "1.rpt")
            
            Else
                .cadRegSelec = SQL2SF(cadB)
            
            End If
        Else
            .cadRegSelec = ""
        End If
        If ConFechaBaja Then
            If .cadRegSelec <> "" Then .cadRegSelec = .cadRegSelec & " and "
            .cadRegSelec = .cadRegSelec & "isnull({rpozos.fechabaja})"
        End If
        .cadRegActua = POS2SF(Data1, Me)
        '[Monica]13/07/2012: falla si hay un solo registro seleccionado y apretamos registros buscados
        If Data1.Recordset.RecordCount = 1 Then .cadRegSelec = .cadRegActua
        .cadTodosReg = ""
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|" & Orden & "|pUsu=" & vUsu.Codigo & "|"
        .NumeroParametros2 = 2
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = True
        .SubInformeConta = ""
        .Show vbModal
    End With
End Sub

Private Sub printNouIndefa()
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
    
    
    indRPT = 78 ' personalizacion del informe de hidrantes
    
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
    
    nomDocu = "EscContadorIndefa.rpt"
    
    With frmImprimir2
        .cadTabla2 = "rpozos"
        .Informe2 = nomDocu
        If cadB <> "" Then
            .cadRegSelec = SQL2SF(cadB)
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




'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGridAux_GotFocus(Index As Integer)
'  WheelHook DataGridAux(Index)
'End Sub
'Private Sub DataGridAux_LostFocus(Index As Integer)
'  WheelUnHook
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

Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String

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

Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim Sql As String

    On Error Resume Next
    
    Sql = " numnotac= " & Text1(0).Text
    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function



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
        Case 0
            Sql = "select rpozos_cooprop.hidrante, rpozos_cooprop.numlinea, rpozos_cooprop.codsocio, rsocios.nomsocio, "
            Sql = Sql & " rpozos_cooprop.porcentaje "
            Sql = Sql & " FROM rpozos_cooprop INNER JOIN rsocios ON rpozos_cooprop.codsocio = rsocios.codsocio "
            Sql = Sql & " and rpozos_cooprop.codsocio = rsocios.codsocio "
            If enlaza Then
                Sql = Sql & " WHERE rpozos_cooprop.hidrante = '" & Trim(Text1(0).Text) & "'"
            Else
                Sql = Sql & " WHERE rpozos_cooprop.hidrante is null"
            End If
            Sql = Sql & " ORDER BY rpozos_cooprop.codsocio "
        
        Case 1
            Sql = "select rpozos_campos.hidrante, rpozos_campos.numlinea, rpozos_campos.codcampo, rpartida.nomparti, "
            Sql = Sql & " rpueblos.despobla, round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",2) supcoope, rcampos.poligono, rcampos.parcela "
            Sql = Sql & " FROM ((rpozos_campos INNER JOIN rcampos ON rpozos_campos.codcampo = rcampos.codcampo)"
            Sql = Sql & " INNER JOIN rpartida ON rcampos.codparti = rpartida.codparti) "
            Sql = Sql & " INNER JOIN rpueblos ON rpartida.codpobla = rpueblos.codpobla "
            If enlaza Then
                Sql = Sql & " WHERE rpozos_campos.hidrante = '" & Trim(Text1(0).Text) & "'"
            Else
                Sql = Sql & " WHERE rpozos_campos.hidrante is null"
            End If
            Sql = Sql & " ORDER BY rpozos_campos.codcampo "
       
    End Select
    ' ********************************************************************************
    
    MontaSQLCarga = Sql
End Function
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
        Case 0 'coopropietarios
            Sql = "¿Seguro que desea eliminar el coopropietario?"
            Sql = Sql & vbCrLf & "Coopropietario: " & Adoaux(Index).Recordset!Codsocio & " - " & Adoaux(Index).Recordset!nomsocio
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM rpozos_cooprop"
                Sql = Sql & " WHERE rpozos_cooprop.hidrante = " & DBSet(Adoaux(Index).Recordset!Hidrante, "T")
                Sql = Sql & " and codsocio = " & Adoaux(Index).Recordset!Codsocio
            End If
        Case 1 ' campos
            Sql = "¿Seguro que desea eliminar el campo del hidrante?"
            Sql = Sql & vbCrLf & "Campo: " & Adoaux(Index).Recordset!codCampo
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM rpozos_campos"
                Sql = Sql & " WHERE rpozos_campos.hidrante = " & DBSet(Adoaux(Index).Recordset!Hidrante, "T")
                Sql = Sql & " and numlinea = " & Adoaux(Index).Recordset!NumLinea
            End If
        
    End Select

    If Eliminar Then
        NumRegElim = Adoaux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        conn.Execute Sql
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        If Index <> 3 Then _
            CargaGrid Index, True
        ' ***************************************************
        If Not SituarDataTrasEliminar(Adoaux(Index), NumRegElim, True) Then
            
        End If
        SumaTotalPorcentajes NumTabMto
        ' *** si n'hi han tabs sense datagrid ***
        ' ***************************************
        If BLOQUEADesdeFormulario2(Me, Data1, 2, "Frame2") Then BotonModificar
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
        Case 0: vtabla = "rpozos_cooprop"
        Case 1: vtabla = "rpozos_campos"
    End Select
    ' ********************************************************

    vWhere = ObtenerWhereCab(False)

    Select Case Index
         Case 0, 1, 2 'clasificacion
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            Select Case Index
                Case 0
                    NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", "rpozos_cooprop.hidrante = '" & Trim(Text1(0).Text) & "'")
                Case 1
                    NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", "rpozos_campos.hidrante = '" & Trim(Text1(0).Text) & "'")
            End Select
            ' ***************************************************************

            AnyadirLinea DataGridAux(Index), Adoaux(Index)

            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If

            LLamaLineas Index, ModoLineas, anc

            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
                Case 0 'copropietarios
                    For i = 0 To txtAux3.Count - 1
                        txtAux3(i).Text = ""
                    Next i
                    txtAux2(0).Text = ""
                    txtAux3(0).Text = Text1(0).Text 'codcampo
                    txtAux3(1).Text = NumF 'numlinea
                    txtAux3(2).Text = ""
                    PonerFoco txtAux3(2)
                Case 1 ' campos
                    For i = 0 To txtAux4.Count - 1
                        txtAux4(i).Text = ""
                    Next i
                    For i = 1 To 5
                        txtAux2(i).Text = ""
                    Next i
                    txtAux4(0).Text = Text1(0).Text ' codcampo
                    txtAux4(1).Text = NumF 'numlinea
                    PonerFoco txtAux4(2)
                
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
        Case 0 'coopropietarios
            txtAux3(0).Text = DataGridAux(Index).Columns(0).Text
            txtAux3(1).Text = DataGridAux(Index).Columns(1).Text
            txtAux3(2).Text = DataGridAux(Index).Columns(2).Text
            
            txtAux2(0).Text = DataGridAux(Index).Columns(3).Text
            txtAux3(3).Text = DataGridAux(Index).Columns(4).Text
        
        Case 1 ' campos
            txtAux4(0).Text = DataGridAux(Index).Columns(0).Text
            txtAux4(1).Text = DataGridAux(Index).Columns(1).Text
            txtAux4(2).Text = DataGridAux(Index).Columns(2).Text
            
            txtAux2(1).Text = DataGridAux(Index).Columns(3).Text
            txtAux2(2).Text = DataGridAux(Index).Columns(4).Text
            txtAux2(3).Text = DataGridAux(Index).Columns(5).Text
            txtAux2(4).Text = DataGridAux(Index).Columns(6).Text
            txtAux2(5).Text = DataGridAux(Index).Columns(7).Text
        
    
    End Select

    LLamaLineas Index, ModoLineas, anc

    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'coopropietarios
            PonerFoco txtAux3(2)
        Case 1 ' campos
            PonerFoco txtAux4(2)
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
        Case 1 ' campos
            For jj = 2 To txtAux4.Count - 1
                txtAux4(jj).visible = b
                txtAux4(jj).Top = alto
            Next jj
            For jj = 1 To 5
                txtAux2(jj).visible = b
                txtAux2(jj).Top = alto
            Next jj
            cmdAux(1).visible = b
            cmdAux(1).Top = txtAux4(2).Top
            cmdAux(1).Height = txtAux4(2).Height
    
    End Select
End Sub




Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Búscar
            cmdSigpac_Click
    End Select
End Sub


Private Sub Toolbar3_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    
    Select Case Index
        Case 0, 1, 2, 3
            Select Case Button.Index
                Case 1  'Búscar
                    ' si existe el fichero de ampliacion lo mostramos
                    If Dir(Toolbar3(Index).Tag) <> "0" Then
                    
                        Set frmMensImg = New frmMensajes
                        
                        frmMensImg.cadena = Toolbar3(Index).Tag
                        frmMensImg.OpcionMensaje = 45
                        frmMensImg.Show vbModal
                        
                        Set frmMens = Nothing
                    
                    End If
            End Select
        Case 4, 5, 6, 7
            Select Case Button.Index
                Case 1  'Búscar
                    If InStr(1, Toolbar3(Index).Tag, "http") <> 0 Then
                        Screen.MousePointer = vbHourglass
                    
                        If LanzaHomeGnral(Toolbar3(Index).Tag) Then espera 2
                        
                        Screen.MousePointer = vbDefault
                    Else
                        If Dir(Toolbar3(Index).Tag) <> "0" Then
                        
                            Set frmMensImg = New frmMensajes
                            
                            frmMensImg.cadena = Toolbar3(Index).Tag
                            frmMensImg.OpcionMensaje = 45
                            frmMensImg.Show vbModal
                            
                            Set frmMensImg = Nothing
                        
                        End If
                    End If
            End Select
    End Select
     
End Sub

Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 ' buscar diferencias
            mnDiferencias_Click
        Case 2 ' Actualizar registro con datos de indefa
            mnActualizar_Click
    End Select
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
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
            If txtAux3(2).Text <> "" Then CmdAceptar.SetFocus
    
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

'AAAAAAAAAAAAAAAAAAAAAAA
Private Sub TxtAux4_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim i As Integer
Dim Sql As String


    If Not PerderFocoGnral(txtAux4(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 2 ' campo
            If PonerFormatoEntero(txtAux4(Index)) Then
                Sql = ""
                Sql = DevuelveDesdeBDNew(cAgro, "rcampos", "codcampo", "codcampo", txtAux4(Index).Text, "N")
                If Sql = "" Then
                    cadMen = "No existe el Campo: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCam1 = New frmManCampos
                        frmCam1.DatosADevolverBusqueda = "0|1|"
'                        frmCamp.NuevoCodigo = Text1(Index).Text
                        txtAux4(Index).Text = ""
                        TerminaBloquear
                        frmCam1.Show vbModal
                        Set frmCam1 = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtAux4(Index).Text = ""
                    End If
                    PonerFoco txtAux4(Index)
                Else
                    If Not EstaCampoDeAlta(txtAux4(Index).Text) Then
                        MsgBox "El campo no está dado de alta. Reintroduzca.", vbExclamation
                        txtAux4(Index).Text = ""
                        PonerFoco txtAux4(Index)
                    Else
                        PonerDatosCampoLineas (txtAux4(Index))
                    End If
                End If
            Else
                For i = 1 To 5
                    txtAux2(i).Text = ""
                Next i
            End If
            
        Case 3 'porcentaje de
            PonerFormatoDecimal txtAux4(Index), 4
            If txtAux4(2).Text <> "" Then CmdAceptar.SetFocus
    
    End Select

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


'AAAAAAAAAAAAAAAAAAAAAAA
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


Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
Dim b As Boolean
Dim i As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    b = DataGridAux(Index).Enabled
    DataGridAux(Index).Enabled = False
    
    Adoaux(Index).ConnectionString = conn
    Adoaux(Index).RecordSource = MontaSQLCarga(Index, enlaza)
    Adoaux(Index).CursorType = adOpenDynamic
    Adoaux(Index).LockType = adLockPessimistic
    DataGridAux(Index).ScrollBars = dbgNone
    Adoaux(Index).Refresh
    Set DataGridAux(Index).DataSource = Adoaux(Index)
    
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
    
    CargaGridGnral Me.DataGridAux(Index), Me.Adoaux(Index), tots, PrimeraVez
    
    'DataGridAux(Index).Enabled = b
'    PrimeraVez = False
    
    Select Case Index
        Case 0 ' coopropietarios
            tots = "N||||0|;N||||0|;S|txtaux3(2)|T|Código|1200|;S|cmdAux(0)|B|||;" 'codsocio,numlinea
            tots = tots & "S|txtAux2(0)|T|Nombre|4270|;"
            tots = tots & "S|txtaux3(3)|T|Porcentaje|1500|;"
            arregla tots, DataGridAux(Index), Me
        
            DataGridAux(Index).Columns(4).Alignment = dbgRight
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))

'            BloquearTxt txtAux(14), Not b
'            BloquearTxt txtAux(15), Not b

            If (enlaza = True) And (Not Adoaux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
'                SumaTotalPorcentajes
            Else
                For i = 0 To 3
                    txtAux3(i).Text = ""
                Next i
                txtAux2(0).Text = ""
            End If
        Case 1 ' CAMPOS
            tots = "N||||0|;N||||0|;S|txtaux4(2)|T|Campo|1200|;S|cmdAux(1)|B|||;" 'codsocio,numlinea
            tots = tots & "S|txtAux2(1)|T|Partida|2000|;"
            tots = tots & "S|txtAux2(2)|T|Población|1670|;"
            tots = tots & "S|txtAux2(3)|T|Hdas|800|;"
            tots = tots & "S|txtAux2(4)|T|Pol|600|;"
            tots = tots & "S|txtAux2(5)|T|Par|800|;"
            
            arregla tots, DataGridAux(Index), Me
        
            DataGridAux(Index).Columns(5).Alignment = dbgRight
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))


            If (enlaza = True) And (Not Adoaux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
'                SumaTotalPorcentajes
            Else
                For i = 2 To 2
                    txtAux4(i).Text = ""
                Next i
                For i = 1 To 5
                    txtAux2(i).Text = ""
                Next i
            End If
         
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


Private Sub InsertarLinea()
'Inserta registre en les taules de Llínies
Dim nomframe As String
Dim b As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'coopropietarios
        Case 1: nomframe = "FrameAux1" 'campos
    End Select
    ' ***************************************************************
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomframe) Then
            ' *** si n'hi ha que fer alguna cosa abas d'insertar
            ' *************************************************
            b = BLOQUEADesdeFormulario2(Me, Data1, 2, "Frame2")
            
            '++monica: en caso de estar insertando seccion y que no existan las
            'cuentas contables hacemos esto para que las inserte en contabilidad.
'            If NumTabMto = 1 Then
'               txtAux2(4).Text = PonerNombreCuenta(txtaux1(4), 3, Text1(0))
'               txtAux2(5).Text = PonerNombreCuenta(txtaux1(5), 3, Text1(0))
'            End If
            
            Select Case NumTabMto
                Case 0, 1 ' *** els index de les llinies en grid (en o sense tab) ***
                    CargaGrid NumTabMto, True
                    If b Then BotonAnyadirLinea NumTabMto
'                Case 3 ' *** els index dels tabs que NO tenen grid ***
'                    CargaFrame 3, True
'                    If b Then BotonModificar
'                    ModoLineas = 0
'                LLamaLineas NumTabMto, 0
                
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
        Case 1: nomframe = "FrameAux1" 'campos
        Case 2: nomframe = "FrameAux2" 'parcelas
    End Select
    ' **************************************************************

    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
            ' *** si cal que fer alguna cosa abas d'insertar ***
            If NumTabMto = 0 Then
            End If
            ' ******************************************************
            If BLOQUEADesdeFormulario2(Me, Data1, 2, "Frame2") Then BotonModificar
            ModoLineas = 0

            If NumTabMto <> 3 Then
                V = Adoaux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
                CargaGrid NumTabMto, True
            End If

            ' *** si n'hi han tabs ***
            'SituarTab (NumTabMto)

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


Private Sub SumaTotalPorcentajes(numTab As Integer)
Dim Sql As String
Dim i As Currency
Dim Rs As ADODB.Recordset
   
   Select Case numTab
        Case 0 ' coopropietarios
            Sql = "select sum(porcentaje) from rpozos_cooprop where rpozos_cooprop.hidrante = " & DBSet(Data1.Recordset!Hidrante, "T")
            
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
   
        
   End Select

End Sub


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
    
    
Private Sub cmdSigpac_Click()
Dim Direccion As String
Dim Pobla As String
Dim Municipio As String
Dim Rs As ADODB.Recordset
Dim Sql As String

    TerminaBloquear

    
    'http://sigpac.mapa.es/fega/visor/LayerInfo.aspx?layer=PARCELA&id=OID&image=ORTOFOTOS
'    Direccion = "http://sigpac.mapa.es/fega/visor/LayerInfo.aspx?layer=PARCELA&id=" & Trim(Text1(18).Text) & "&image=ORTOFOTOS"
    
    If vParamAplic.SigPac <> "" Then
        If InStr(1, vParamAplic.SigPac, "NUMOID") <> 0 Then
            Sql = "select numeroid from rcampos where codcampo = " & DBSet(Text1(18).Text, "N")
            
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
            If Not Rs.EOF Then
                Direccion = Replace(vParamAplic.SigPac, "NUMOID", DBLet(Rs!numeroid))
            End If
        Else
            If txtAux1(3).Text <> "" And txtAux1(4).Text <> "" Then
'                Sql = "select codparti, recintos from rcampos where codcampo = " & DBSet(Text1(18).Text, "N")
'
'                Set Rs = New ADODB.Recordset
'                Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                If Not Rs.EOF Then
            
'                    Pobla = DevuelveDesdeBDNew(cAgro, "rpartida", "codpobla", "codparti", DBLet(Rs!codparti), "N")
                    Pobla = vParam.CPostal
                    If Pobla = "" Then
                        MsgBox "No existe el código de poblacion de la partida", vbExclamation
                    Else
                        Municipio = DevuelveDesdeBDNew(cAgro, "rpueblos", "codsigpa", "codpobla", Pobla, "T")
                        Direccion = Replace(vParamAplic.SigPac, "[PR]", Mid(Pobla, 1, 2))
                        Direccion = Replace(Direccion, "[MN]", CInt(Municipio))
                        Direccion = Replace(Direccion, "[PL]", CInt(ComprobarCero(txtAux1(3).Text)))
                        
                        If InStr(txtAux1(14).Text, ",") Then
                            'cogemos unicamente la primera parcela
                            Direccion = Replace(Direccion, "[PC]", CInt(ComprobarCero(Mid(txtAux1(4).Text, 1, InStr(txtAux1(4).Text, ",") - 1))))
                        Else
                            Direccion = Replace(Direccion, "[PC]", CInt(ComprobarCero(txtAux1(4).Text)))
                        End If
                        Direccion = Replace(Direccion, "[RC]", 1) 'CInt(ComprobarCero(Rs!recintos)))
                    End If
'                End If
            Else
                MsgBox "No existe el polígono y/o parcela Indefa.", vbExclamation
                Exit Sub
            End If
        End If
    Else
        MsgBox "No tiene configurada en parámetros la dirección de Sigpac. Llame a Soporte.", vbExclamation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass

    If LanzaHomeGnral(Direccion) Then espera 2
    Screen.MousePointer = vbDefault


End Sub

