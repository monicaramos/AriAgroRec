VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPOZRecibosMonast 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hist�rico de Recibos "
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   10350
   Icon            =   "frmPOZRecibosMonast.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameRectifica 
      Caption         =   "Factura Rectificativa "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1005
      Left            =   180
      TabIndex        =   142
      Top             =   3135
      Width           =   15
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
         Index           =   45
         Left            =   180
         MaxLength       =   3
         TabIndex        =   145
         Tag             =   "Tipo Movimiento Fra Rectifica|T|S|||rrecibpozos|codtipomrec|||"
         Top             =   540
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
         Index           =   47
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   144
         Tag             =   "Fecha Factura Rectificativa|F|S|||rrecibpozos|fecfacturec|dd/mm/yyyy||"
         Text            =   "123"
         Top             =   540
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
         Index           =   46
         Left            =   1800
         MaxLength       =   7
         TabIndex        =   143
         Tag             =   "N� Factura Rectifica|N|S|||rrecibpozos|numfacturec|0000000||"
         Text            =   "Text1"
         Top             =   540
         Width           =   930
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   6
         Left            =   4035
         Picture         =   "frmPOZRecibosMonast.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Recibo"
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
         Index           =   17
         Left            =   180
         TabIndex        =   148
         Top             =   255
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "N� Recibo"
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
         Index           =   14
         Left            =   1800
         TabIndex        =   147
         Top             =   255
         Width           =   1065
      End
      Begin VB.Label Label22 
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
         Left            =   2940
         TabIndex        =   146
         Top             =   255
         Width           =   1035
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   225
      TabIndex        =   153
      Top             =   0
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   154
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
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3855
      TabIndex        =   151
      Top             =   0
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   152
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
      Left            =   6795
      TabIndex        =   150
      Top             =   180
      Width           =   1605
   End
   Begin VB.Frame Frame5 
      Caption         =   "Total Factura"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1050
      Index           =   0
      Left            =   180
      TabIndex        =   42
      Top             =   3090
      Width           =   10035
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
         Left            =   5955
         MaxLength       =   10
         TabIndex        =   23
         Tag             =   "Importe Iva|N|S|||rrecibpozos|imporiva|###,##0.00||"
         Top             =   540
         Width           =   1650
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
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   21
         Tag             =   "Tipo Iva|N|S|||rrecibpozos|tipoiva|00||"
         Text            =   "Text1"
         Top             =   540
         Width           =   600
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
         Left            =   5100
         MaxLength       =   6
         TabIndex        =   22
         Tag             =   "Porc.Iva|N|S|||rrecibpozos|porc_iva|##0.00||"
         Text            =   "123"
         Top             =   540
         Width           =   735
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
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   83
         Text            =   "Text2"
         Top             =   540
         Width           =   2550
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
         Left            =   180
         MaxLength       =   10
         TabIndex        =   20
         Tag             =   "Base Imponible|N|N|||rrecibpozos|baseimpo|###,##0.00||"
         Top             =   540
         Width           =   1530
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CAE3FD&
         BeginProperty Font 
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
         Left            =   7905
         MaxLength       =   10
         TabIndex        =   24
         Tag             =   "Total Factura|N|N|||rrecibpozos|totalfact|###,##0.00||"
         Top             =   540
         Width           =   1830
      End
      Begin VB.Label Label2 
         Caption         =   "% Iva"
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
         Left            =   5070
         TabIndex        =   47
         Top             =   300
         Width           =   630
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   8
         Left            =   1800
         TabIndex        =   46
         Top             =   300
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   2220
         ToolTipText     =   "Buscar Iva"
         Top             =   300
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Importe IVA"
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
         Left            =   5985
         TabIndex        =   45
         Top             =   300
         Width           =   1500
      End
      Begin VB.Label Label1 
         Caption         =   "TOTAL FACTURA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   7905
         TabIndex        =   44
         Top             =   300
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
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
         Index           =   10
         Left            =   180
         TabIndex        =   43
         Top             =   300
         Width           =   1545
      End
   End
   Begin VB.PictureBox cmdRectificativa 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   6765
      ScaleHeight     =   435
      ScaleWidth      =   525
      TabIndex        =   141
      ToolTipText     =   "Rectificativa"
      Top             =   510
      Width           =   525
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Es Contado"
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
      Index           =   3
      Left            =   4485
      TabIndex        =   135
      Tag             =   "Es Contado|N|N|0|1|rrecibpozos|escontado|0||"
      Top             =   2805
      Width           =   1635
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
      Index           =   40
      Left            =   3105
      MaxLength       =   10
      TabIndex        =   15
      Tag             =   "Fec.Albaran|F|S|||rrecibpozos|fecalbar|dd/mm/yyyy||"
      Text            =   "1234567890123456789012345"
      Top             =   2685
      Width           =   1275
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
      Index           =   39
      Left            =   1710
      MaxLength       =   10
      TabIndex        =   14
      Tag             =   "Ticket|N|S|||rrecibpozos|numalbar|0000000||"
      Text            =   "1234567890"
      Top             =   2685
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
      Index           =   38
      Left            =   8685
      MaxLength       =   10
      TabIndex        =   10
      Tag             =   "Toma|N|S|||rrecibpozos|nroorden|000000||"
      Text            =   "1234567890123456789012345"
      Top             =   1650
      Width           =   1470
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
      Index           =   37
      Left            =   6150
      MaxLength       =   25
      TabIndex        =   9
      Tag             =   "Parcelas|T|S|||rrecibpozos|parcelas|||"
      Text            =   "1234567890123456789012345"
      Top             =   1650
      Width           =   2460
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
      Index           =   36
      Left            =   4470
      MaxLength       =   10
      TabIndex        =   8
      Tag             =   "Poligono|T|S|||rrecibpozos|poligono|||"
      Text            =   "1234567890"
      Top             =   1650
      Width           =   1545
   End
   Begin VB.PictureBox cmdCampos 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   7455
      ScaleHeight     =   435
      ScaleWidth      =   525
      TabIndex        =   114
      ToolTipText     =   "Campos"
      Top             =   510
      Width           =   525
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Pasa Aridoc"
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
      Index           =   2
      Left            =   4485
      TabIndex        =   19
      Tag             =   "Pasa Aridoc|N|N|0|1|rrecibpozos|pasaridoc|0||"
      Top             =   2565
      Width           =   1650
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
      Index           =   35
      Left            =   3105
      MaxLength       =   10
      TabIndex        =   13
      Tag             =   "Precio|N|S|||rrecibpozos|precio|###,##0.0000||"
      Text            =   "1234567890"
      Top             =   2295
      Width           =   1275
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
      Index           =   34
      Left            =   1710
      MaxLength       =   10
      TabIndex        =   12
      Tag             =   "Importe Dto|N|S|||rrecibpozos|impdto|##,###,##0.00||"
      Top             =   2295
      Width           =   1215
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
      Index           =   33
      Left            =   210
      MaxLength       =   10
      TabIndex        =   11
      Tag             =   "Porc.Dto|N|S|||rrecibpozos|porcdto|##0.00||"
      Top             =   2295
      Width           =   1305
   End
   Begin VB.PictureBox cmdHidrantes 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   8175
      ScaleHeight     =   435
      ScaleWidth      =   525
      TabIndex        =   100
      ToolTipText     =   "Hidrantes"
      Top             =   510
      Width           =   525
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
      Index           =   32
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   7
      Tag             =   "Diferencia Dias|N|S|||rrecibpozos|difdias|###,##0||"
      Text            =   "1234567"
      Top             =   1650
      Width           =   690
   End
   Begin VB.PictureBox cmdParticipa 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   8895
      ScaleHeight     =   435
      ScaleWidth      =   525
      TabIndex        =   88
      ToolTipText     =   "Participaciones"
      Top             =   510
      Width           =   525
   End
   Begin VB.PictureBox cmdConceptos 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   9615
      ScaleHeight     =   435
      ScaleWidth      =   525
      TabIndex        =   82
      ToolTipText     =   "Conceptos"
      Top             =   510
      Width           =   525
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
      TabIndex        =   0
      Top             =   990
      Width           =   1650
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   20
      Left            =   630
      MaxLength       =   10
      TabIndex        =   66
      Tag             =   "Tipo de Fichero|T|S|||rrecibpozos|codtipom||S|"
      Top             =   990
      Width           =   1065
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
      Left            =   4470
      MaxLength       =   6
      TabIndex        =   3
      Tag             =   "Cod.Socio|N|N|0|999999|rrecibpozos|codsocio|000000|N|"
      Text            =   "Text1"
      Top             =   990
      Width           =   1005
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
      Left            =   210
      MaxLength       =   10
      TabIndex        =   4
      Tag             =   "Hidrante|T|S|||rrecibpozos|hidrante||N|"
      Text            =   "1234567890"
      Top             =   1650
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
      Index           =   12
      Left            =   1710
      MaxLength       =   7
      TabIndex        =   5
      Tag             =   "Consumo|N|S|||rrecibpozos|consumo|||"
      Text            =   "1234567"
      Top             =   1650
      Width           =   975
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
      Height          =   780
      Index           =   15
      Left            =   6135
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Tag             =   "Conceptol|T|S|||rrecibpozos|concepto|||"
      Text            =   "frmPOZRecibosMonast.frx":0097
      Top             =   2250
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
      Index           =   13
      Left            =   2700
      MaxLength       =   10
      TabIndex        =   6
      Tag             =   "Cuota|N|S|||rrecibpozos|impcuota|###,##0.00||"
      Text            =   "1234567"
      Top             =   1650
      Width           =   1020
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
      Index           =   0
      Left            =   1920
      MaxLength       =   7
      TabIndex        =   1
      Tag             =   "N� Factura|N|S|||rrecibpozos|numfactu|0000000|S|"
      Top             =   990
      Width           =   1065
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
      Left            =   3045
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "Fecha Factura|F|N|||rrecibpozos|fecfactu|dd/mm/yyyy|S|"
      Top             =   990
      Width           =   1335
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
      Left            =   5505
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   65
      Text            =   "Text2"
      Top             =   990
      Width           =   4650
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Contabilizado"
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
      Index           =   1
      Left            =   4485
      TabIndex        =   18
      Tag             =   "Contabilizado|N|N|0|1|rrecibpozos|contabilizado|0||"
      Top             =   2325
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Impreso"
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
      Index           =   0
      Left            =   4485
      TabIndex        =   17
      Tag             =   "Impreso|N|N|0|1|rrecibpozos|impreso|0||"
      Top             =   2055
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Height          =   525
      Index           =   0
      Left            =   210
      TabIndex        =   40
      Top             =   6930
      Width           =   2220
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
         Left            =   240
         TabIndex        =   41
         Top             =   180
         Width           =   1755
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
      Left            =   9090
      TabIndex        =   38
      Top             =   7035
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
      Left            =   7935
      TabIndex        =   37
      Top             =   7050
      Width           =   1095
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
      Left            =   9090
      TabIndex        =   39
      Top             =   7050
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   2880
      Top             =   6450
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   2880
      Top             =   6330
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
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
   Begin MSAdodcLib.Adodc Data4 
      Height          =   360
      Left            =   2970
      Top             =   6270
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   635
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
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   3150
      Top             =   3210
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Height          =   285
      Index           =   31
      Left            =   420
      MaxLength       =   7
      TabIndex        =   86
      Tag             =   "Linea|N|N|||rrecibpozos|numlinea|0000000|S|"
      Text            =   "Text1"
      Top             =   1650
      Width           =   885
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   9735
      TabIndex        =   155
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
   Begin VB.Frame FrameCampos 
      Caption         =   "Campos"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2730
      Left            =   180
      TabIndex        =   115
      Top             =   4170
      Visible         =   0   'False
      Width           =   10035
      Begin VB.Frame FrameAux1 
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   120
         TabIndex        =   116
         Top             =   300
         Width           =   9810
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
            Index           =   11
            Left            =   6720
            MaxLength       =   15
            TabIndex        =   149
            Text            =   "SP"
            Top             =   1530
            Visible         =   0   'False
            Width           =   555
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
            Index           =   10
            Left            =   6060
            MaxLength       =   9
            TabIndex        =   129
            Tag             =   "SubParcela|T|S|||rrecibpozos_cam|subparce|||"
            Text            =   "SP"
            Top             =   1530
            Visible         =   0   'False
            Width           =   555
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
            Index           =   9
            Left            =   5460
            MaxLength       =   6
            TabIndex        =   128
            Tag             =   "Parcela|N|S|||rrecibpozos_cam|parcela|#####0||"
            Text            =   "Parcela"
            Top             =   1530
            Visible         =   0   'False
            Width           =   555
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
            Index           =   8
            Left            =   4860
            MaxLength       =   3
            TabIndex        =   127
            Tag             =   "Poligono|N|S|||rrecibpozos_cam|poligono|##0||"
            Text            =   "Poligono"
            Top             =   1530
            Visible         =   0   'False
            Width           =   555
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
            Index           =   7
            Left            =   4260
            MaxLength       =   9
            TabIndex        =   126
            Tag             =   "Precio2|N|S|||rrecibpozos_cam|precio2|#,##0.0000||"
            Text            =   "Hanegada"
            Top             =   1530
            Visible         =   0   'False
            Width           =   555
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
            Left            =   3660
            MaxLength       =   9
            TabIndex        =   125
            Tag             =   "Precio1|N|S|||rrecibpozos_cam|precio1|#,##0.0000||"
            Text            =   "Hanegada"
            Top             =   1530
            Visible         =   0   'False
            Width           =   555
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
            Index           =   3
            Left            =   2160
            MaxLength       =   7
            TabIndex        =   120
            Tag             =   "Linea|N|N|||rrecibpozos_cam|numlinea|0000000|S|"
            Text            =   "Linea"
            Top             =   1530
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtAux5 
            Alignment       =   1  'Right Justify
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
            Index           =   2
            Left            =   1350
            MaxLength       =   10
            TabIndex        =   119
            Tag             =   "Fecha Factura|F|N|||rrecibpozos_cam|fecfactu|dd/mm/yyyy|S|"
            Text            =   "fecha"
            Top             =   1530
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.TextBox txtAux5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
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
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   1
            Left            =   750
            MaxLength       =   7
            TabIndex        =   118
            Tag             =   "N� Factura|N|N|||rrecibpozos_cam|numfactu|0000000|S|"
            Text            =   "recibo"
            Top             =   1530
            Visible         =   0   'False
            Width           =   585
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
            Left            =   180
            MaxLength       =   6
            TabIndex        =   117
            Tag             =   "Tipo Mov|T|N|||rrecibpozos_cam|codtipom||S|"
            Text            =   "TipoM"
            Top             =   1530
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
            Index           =   5
            Left            =   3060
            MaxLength       =   9
            TabIndex        =   124
            Tag             =   "Hanegada|N|S|||rrecibpozos_cam|hanegada|#,##0.0000||"
            Text            =   "Hanegada"
            Top             =   1530
            Visible         =   0   'False
            Width           =   555
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
            Left            =   2580
            MaxLength       =   10
            TabIndex        =   122
            Tag             =   "Campo|N|N|||rrecibpozos_cam|codcampo|00000000|S|"
            Text            =   "Hidrante"
            Top             =   1530
            Visible         =   0   'False
            Width           =   465
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   1
            Left            =   90
            TabIndex        =   121
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
            Left            =   3765
            Top             =   330
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
            Bindings        =   "frmPOZRecibosMonast.frx":009F
            Height          =   1395
            Index           =   1
            Left            =   105
            TabIndex        =   123
            Top             =   420
            Width           =   9520
            _ExtentX        =   16801
            _ExtentY        =   2461
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
   Begin VB.Frame FrameHidrantes 
      Caption         =   "Hidrantes"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2730
      Left            =   180
      TabIndex        =   101
      Top             =   4170
      Visible         =   0   'False
      Width           =   9990
      Begin VB.Frame FrameAux0 
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   120
         TabIndex        =   102
         Top             =   300
         Width           =   9665
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
            Left            =   2580
            MaxLength       =   10
            TabIndex        =   107
            Tag             =   "Hidrante|T|N|||rrecibpozos_hid|hidrante||S|"
            Text            =   "Hidrante"
            Top             =   1530
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
            Index           =   5
            Left            =   3060
            MaxLength       =   9
            TabIndex        =   108
            Tag             =   "Hanegada|N|S|||rrecibpozos_hid|hanegada|#,##0.0000||"
            Text            =   "Hanegada"
            Top             =   1530
            Visible         =   0   'False
            Width           =   555
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
            Left            =   300
            MaxLength       =   6
            TabIndex        =   106
            Tag             =   "Tipo Mov|T|N|||rrecibpozos_hid|codtipom||S|"
            Text            =   "TipoM"
            Top             =   1530
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtAux4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
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
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   1
            Left            =   750
            MaxLength       =   7
            TabIndex        =   105
            Tag             =   "N� Factura|N|N|||rrecibpozos_hid|numfactu|0000000|S|"
            Text            =   "recibo"
            Top             =   1530
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox txtAux4 
            Alignment       =   1  'Right Justify
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
            Index           =   2
            Left            =   1350
            MaxLength       =   10
            TabIndex        =   104
            Tag             =   "Fecha Factura|F|N|||rrecibpozos_hid|fecfactu|dd/mm/yyyy|S|"
            Text            =   "fecha"
            Top             =   1530
            Visible         =   0   'False
            Width           =   765
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
            Index           =   3
            Left            =   2160
            MaxLength       =   7
            TabIndex        =   103
            Tag             =   "Linea|N|N|||rrecibpozos_hid|numlinea|0000000|S|"
            Text            =   "Linea"
            Top             =   1530
            Visible         =   0   'False
            Width           =   375
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   0
            Left            =   90
            TabIndex        =   109
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
            Left            =   3765
            Top             =   330
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
            Bindings        =   "frmPOZRecibosMonast.frx":00B7
            Height          =   1395
            Index           =   0
            Left            =   105
            TabIndex        =   110
            Top             =   420
            Width           =   3120
            _ExtentX        =   5503
            _ExtentY        =   2461
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
   Begin VB.Frame FrameConceptos 
      Caption         =   "Conceptos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2745
      Left            =   180
      TabIndex        =   67
      Top             =   4170
      Width           =   10005
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
         Index           =   30
         Left            =   8040
         MaxLength       =   100
         TabIndex        =   77
         Tag             =   "Importe Art.4|N|S|||rrecibpozos|importear4|###,##0.00||"
         Top             =   2280
         Width           =   1530
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
         Index           =   29
         Left            =   240
         MaxLength       =   100
         TabIndex        =   76
         Tag             =   "Concepto Articulo 4|T|S|||rrecibpozos|conceptoar4|||"
         Text            =   "1234567"
         Top             =   2280
         Width           =   7605
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
         Left            =   8040
         MaxLength       =   100
         TabIndex        =   75
         Tag             =   "Importe Art.3|N|S|||rrecibpozos|importear3|###,##0.00||"
         Top             =   1890
         Width           =   1530
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
         Left            =   240
         MaxLength       =   100
         TabIndex        =   74
         Tag             =   "Concepto Articulo 3|T|S|||rrecibpozos|conceptoar3|||"
         Text            =   "1234567"
         Top             =   1890
         Width           =   7605
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
         Left            =   8040
         MaxLength       =   100
         TabIndex        =   73
         Tag             =   "Importe Art.2|N|S|||rrecibpozos|importear2|###,##0.00||"
         Top             =   1500
         Width           =   1530
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
         Index           =   25
         Left            =   240
         MaxLength       =   100
         TabIndex        =   72
         Tag             =   "Concepto Articulo 2|T|S|||rrecibpozos|conceptoar2|||"
         Text            =   "1234567"
         Top             =   1500
         Width           =   7605
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
         Left            =   8040
         MaxLength       =   100
         TabIndex        =   71
         Tag             =   "Importe Art.1|N|S|||rrecibpozos|importear1|###,##0.00||"
         Top             =   1110
         Width           =   1530
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
         Left            =   240
         MaxLength       =   100
         TabIndex        =   70
         Tag             =   "Concepto Articulo 1|T|S|||rrecibpozos|conceptoar1|||"
         Text            =   "1234567"
         Top             =   1110
         Width           =   7605
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
         Index           =   22
         Left            =   8040
         MaxLength       =   10
         TabIndex        =   69
         Tag             =   "Importe MO|N|S|||rrecibpozos|importemo|###,##0.00||"
         Top             =   495
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
         Index           =   21
         Left            =   240
         MaxLength       =   100
         TabIndex        =   68
         Tag             =   "Concepto MO|T|S|||rrecibpozos|conceptomo|||"
         Text            =   "1234567"
         Top             =   495
         Width           =   7605
      End
      Begin VB.Label Label11 
         Caption         =   "Importe"
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
         Left            =   8070
         TabIndex        =   81
         Top             =   810
         Width           =   1035
      End
      Begin VB.Label Label10 
         Caption         =   "Importe"
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
         Left            =   8040
         TabIndex        =   80
         Top             =   225
         Width           =   1035
      End
      Begin VB.Label Label8 
         Caption         =   "Art�culos"
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
         TabIndex        =   79
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label7 
         Caption         =   "Mano de Obra"
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
         TabIndex        =   78
         Top             =   225
         Width           =   1530
      End
   End
   Begin VB.Frame FrameParticipaciones 
      Caption         =   "Participaciones"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2730
      Left            =   180
      TabIndex        =   87
      Top             =   4170
      Visible         =   0   'False
      Width           =   9990
      Begin VB.Frame FrameAux2 
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   120
         TabIndex        =   89
         Top             =   300
         Width           =   8865
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
            Height          =   290
            Index           =   6
            Left            =   2160
            MaxLength       =   7
            TabIndex        =   98
            Tag             =   "Linea|N|N|||rrecibpozos_acc|numlinea|0000000|S|"
            Text            =   "Linea"
            Top             =   1530
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
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
            Index           =   5
            Left            =   1350
            MaxLength       =   10
            TabIndex        =   97
            Tag             =   "Fecha Factura|F|N|||rrecibpozos_acc|fecfactu|dd/mm/yyyy|S|"
            Text            =   "fecha"
            Top             =   1530
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
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
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   4
            Left            =   750
            MaxLength       =   7
            TabIndex        =   96
            Tag             =   "N� Factura|N|S|||rrecibpozos_acc|numfactu|0000000|S|"
            Text            =   "recibo"
            Top             =   1530
            Visible         =   0   'False
            Width           =   585
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
            Height          =   290
            Index           =   0
            Left            =   300
            MaxLength       =   6
            TabIndex        =   95
            Tag             =   "Tipo Mov|T|N|||rrecibpozos_acc|codtipom||S|"
            Text            =   "TipoM"
            Top             =   1530
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
            Height          =   290
            Index           =   2
            Left            =   3060
            MaxLength       =   9
            TabIndex        =   91
            Tag             =   "Acciones|N|N|||rrecibpozos_acc|acciones|##0.00||"
            Text            =   "Acciones"
            Top             =   1530
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
            Height          =   290
            Index           =   3
            Left            =   3690
            MaxLength       =   30
            TabIndex        =   92
            Tag             =   "Observaciones|T|S|||rrecibpozos_acc|observac|||"
            Text            =   "observaciones"
            Top             =   1530
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
            Height          =   290
            Index           =   1
            Left            =   2580
            MaxLength       =   9
            TabIndex        =   90
            Tag             =   "Numero Fases|N|N|||rrecibpozos_acc|numfases|000|S|"
            Text            =   "Fases"
            Top             =   1530
            Visible         =   0   'False
            Width           =   465
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   2
            Left            =   90
            TabIndex        =   93
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
            Left            =   3765
            Top             =   330
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
            Bindings        =   "frmPOZRecibosMonast.frx":00CF
            Height          =   1335
            Index           =   2
            Left            =   105
            TabIndex        =   94
            Top             =   420
            Width           =   7950
            _ExtentX        =   14023
            _ExtentY        =   2355
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
   Begin VB.Frame Frame3 
      Caption         =   "Lectura Anterior"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2700
      Left            =   180
      TabIndex        =   48
      Top             =   4200
      Width           =   3165
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
         Index           =   42
         Left            =   1455
         MaxLength       =   10
         TabIndex        =   28
         Tag             =   "Fecha lectura anterior2|F|S|||rrecibpozos|fech_ant2|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   1980
         Width           =   1290
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
         Index           =   41
         Left            =   1455
         MaxLength       =   7
         TabIndex        =   27
         Tag             =   "Lectura Anterior2|N|S|||rrecibpozos|lect_ant2|0000000||"
         Text            =   "1234567"
         Top             =   1500
         Width           =   1290
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
         Left            =   1455
         MaxLength       =   7
         TabIndex        =   25
         Tag             =   "Lectura Anterior|N|S|||rrecibpozos|lect_ant|0000000||"
         Text            =   "1234567"
         Top             =   420
         Width           =   1290
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
         Index           =   11
         Left            =   1455
         MaxLength       =   10
         TabIndex        =   26
         Tag             =   "Fecha lectura anterior|F|S|||rrecibpozos|fech_ant|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   960
         Width           =   1290
      End
      Begin VB.Label Label19 
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
         Left            =   195
         TabIndex        =   137
         Top             =   2010
         Width           =   705
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   4
         Left            =   1065
         Picture         =   "frmPOZRecibosMonast.frx":00E7
         ToolTipText     =   "Buscar fecha"
         Top             =   2010
         Width           =   240
      End
      Begin VB.Label Label17 
         Caption         =   "Contador"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   195
         TabIndex        =   136
         Top             =   1530
         Width           =   1125
      End
      Begin VB.Label Label23 
         Caption         =   "Contador"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   195
         TabIndex        =   50
         Top             =   450
         Width           =   1125
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1065
         Picture         =   "frmPOZRecibosMonast.frx":0172
         ToolTipText     =   "Buscar fecha"
         Top             =   990
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
         Left            =   195
         TabIndex        =   49
         Top             =   990
         Width           =   660
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Lectura Actual"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2700
      Left            =   3435
      TabIndex        =   51
      Top             =   4200
      Width           =   3375
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
         Index           =   44
         Left            =   1410
         MaxLength       =   7
         TabIndex        =   31
         Tag             =   "Contador Actual2|N|S|||rrecibpozos|lect_act2|0000000||"
         Text            =   "1234567"
         Top             =   1470
         Width           =   1335
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
         Index           =   43
         Left            =   1410
         MaxLength       =   10
         TabIndex        =   32
         Tag             =   "Fecha Lectura Actual2|F|S|||rrecibpozos|fech_act2|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   1980
         Width           =   1335
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
         Index           =   10
         Left            =   1410
         MaxLength       =   10
         TabIndex        =   30
         Tag             =   "Fecha Lectura Actual|F|S|||rrecibpozos|fech_act|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   960
         Width           =   1335
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
         Left            =   1410
         MaxLength       =   7
         TabIndex        =   29
         Tag             =   "Contador Actual|N|S|||rrecibpozos|lect_act|0000000||"
         Text            =   "1234567"
         Top             =   420
         Width           =   1335
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   5
         Left            =   1050
         Picture         =   "frmPOZRecibosMonast.frx":01FD
         ToolTipText     =   "Buscar fecha"
         Top             =   1980
         Width           =   240
      End
      Begin VB.Label Label21 
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
         Left            =   270
         TabIndex        =   139
         Top             =   1980
         Width           =   705
      End
      Begin VB.Label Label20 
         Caption         =   "Contador"
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
         Left            =   270
         TabIndex        =   138
         Top             =   1500
         Width           =   1035
      End
      Begin VB.Label Label9 
         Caption         =   "Contador"
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
         Left            =   270
         TabIndex        =   53
         Top             =   450
         Width           =   1035
      End
      Begin VB.Label Label12 
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
         Left            =   270
         TabIndex        =   52
         Top             =   990
         Width           =   705
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1050
         Picture         =   "frmPOZRecibosMonast.frx":0288
         ToolTipText     =   "Buscar fecha"
         Top             =   990
         Width           =   240
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Precios Aplicados"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2700
      Left            =   6900
      TabIndex        =   54
      Top             =   4200
      Width           =   3270
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
         Left            =   1500
         MaxLength       =   8
         TabIndex        =   33
         Tag             =   "Consumo 1|N|S|||rrecibpozos|consumo1|0000000||"
         Text            =   "m3"
         Top             =   450
         Width           =   1380
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
         Left            =   1500
         MaxLength       =   9
         TabIndex        =   36
         Tag             =   "Precio 2|N|S|||rrecibpozos|precio2|#,##0.000||"
         Text            =   "precio2"
         Top             =   1980
         Width           =   1380
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
         Left            =   1500
         MaxLength       =   9
         TabIndex        =   34
         Tag             =   "Precio 1|N|S|||rrecibpozos|precio1|#,##0.000||"
         Text            =   "precio1"
         Top             =   870
         Width           =   1380
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
         Left            =   1500
         MaxLength       =   8
         TabIndex        =   35
         Tag             =   "Consumo 2|N|S|||rrecibpozos|consumo2|0000000||"
         Text            =   "m3"
         Top             =   1530
         Width           =   1380
      End
      Begin VB.Label Label14 
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   300
         TabIndex        =   85
         Top             =   1980
         Width           =   1245
      End
      Begin VB.Label Label13 
         Caption         =   "Hasta m3."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   300
         TabIndex        =   84
         Top             =   1590
         Width           =   1245
      End
      Begin VB.Label Label28 
         Caption         =   "Hasta m3."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   300
         TabIndex        =   56
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label Label27 
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   300
         TabIndex        =   55
         Top             =   870
         Width           =   1245
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "P A G A D O"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   13
      Left            =   2460
      TabIndex        =   140
      Top             =   7020
      Width           =   4515
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "/"
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
      Left            =   2970
      TabIndex        =   134
      Top             =   2745
      Width           =   135
   End
   Begin VB.Image imgFec 
      Height          =   240
      Index           =   3
      Left            =   1440
      Picture         =   "frmPOZRecibosMonast.frx":0313
      ToolTipText     =   "Buscar fecha"
      Top             =   2700
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Albar�n/Fec"
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
      Index           =   12
      Left            =   240
      TabIndex        =   133
      Top             =   2715
      Width           =   1740
   End
   Begin VB.Label Label1 
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
      Index           =   11
      Left            =   8685
      TabIndex        =   132
      Top             =   1350
      Width           =   1215
   End
   Begin VB.Label Label1 
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
      Index           =   6
      Left            =   6150
      TabIndex        =   131
      Top             =   1380
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Poligono"
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
      Left            =   4470
      TabIndex        =   130
      Top             =   1380
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Precio"
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
      Left            =   3105
      TabIndex        =   113
      Top             =   2025
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Importe"
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
      Left            =   1710
      TabIndex        =   112
      Top             =   2025
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "%Bon/Rec"
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
      Left            =   210
      TabIndex        =   111
      Top             =   2025
      Width           =   1710
   End
   Begin VB.Label Label15 
      Caption         =   "D�as"
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
      TabIndex        =   99
      Top             =   1380
      Width           =   705
   End
   Begin VB.Label Label1 
      Caption         =   "N� Recibo"
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
      Index           =   28
      Left            =   1950
      TabIndex        =   64
      Top             =   720
      Width           =   1080
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   5295
      ToolTipText     =   "Buscar Socio"
      Top             =   720
      Width           =   240
   End
   Begin VB.Label Label1 
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
      Index           =   0
      Left            =   4500
      TabIndex        =   63
      Top             =   720
      Width           =   735
   End
   Begin VB.Image imgFec 
      Height          =   240
      Index           =   0
      Left            =   4140
      Picture         =   "frmPOZRecibosMonast.frx":039E
      ToolTipText     =   "Buscar fecha"
      Top             =   720
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha "
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
      Index           =   29
      Left            =   3090
      TabIndex        =   62
      Top             =   720
      Width           =   780
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo Recibo"
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
      TabIndex        =   61
      Top             =   720
      Width           =   945
   End
   Begin VB.Label Label3 
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
      Left            =   1740
      TabIndex        =   60
      Top             =   1380
      Width           =   960
   End
   Begin VB.Label Label4 
      Caption         =   "Cuota"
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
      TabIndex        =   59
      Top             =   1380
      Width           =   810
   End
   Begin VB.Label Label5 
      Caption         =   "Hidrante"
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
      TabIndex        =   58
      Top             =   1380
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Concepto"
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
      Left            =   6135
      TabIndex        =   57
      Top             =   1980
      Width           =   1035
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
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         HelpContextID   =   2
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
Attribute VB_Name = "frmPOZRecibosMonast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'Si se llama de la busqueda en el frmAlmMovimArticulos se accede
'a las tablas del Albaran o de Facturas de movimiento seleccionado (solo consulta)
Public hcoCodMovim As String 'cod. movim
Public hcoCodTipoM As String 'Codigo detalle de Movimiento(ALC)
Public hcoFechaMov As String 'fecha del movimiento

'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmLFac As frmManLinFactSocios 'Lineas de variedades de facturas socios
Attribute frmLFac.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmTIva As frmTipIVAConta 'Tipos de iva de conta
Attribute frmTIva.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'Mensajes para sacar los hidrantes de un socio
Attribute frmMens.VB_VarHelpID = -1

Private WithEvents frmVar As frmComVar 'Form Mto de variedades de comercial
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'Form Mto de socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmCal As frmManCalidades 'Form Mto de calidades
Attribute frmCal.VB_VarHelpID = -1
Private WithEvents frmPOZRecPrev As frmPOZRecibosMonastPrev 'Form Mto de recibos de pozos
Attribute frmPOZRecPrev.VB_VarHelpID = -1

Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

Dim ModificaLineas As Byte
'1.- A�adir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim CodTipoMov As String
'Codigo tipo de movimiento en funci�n del valor en la tabla de par�metros: stipom

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

Dim EsCabecera As Boolean
'Para saber en MandaBusquedaPrevia si busca en la tabla scapla o en la tabla sdirec
Dim NumTabMto As Integer 'Indica quin n� de Tab est� en modo Mantenimient


Dim EsDeVarios As Boolean
'Si el cliente mostrado es de Varios o No

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnPrimero As Byte
'Variable que indica el n�mero del Boton  PrimerRegistro en la Toolbar1


Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal
Dim Indice As Byte
Dim Facturas As String

Dim Cliente As String
Private BuscaChekc As String

Private Sub Check1_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "check1(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "check1(" & Index & ")|"
    End If
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkAux_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkAux(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "chkAux(" & Index & ")|"
    End If
End Sub


Private Sub chkAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim I As Integer

    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'B�SQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOK Then InsertarCabecera
        
        Case 4  'MODIFICAR
            If DatosOK Then
                If ModificaDesdeFormulario2(Me, 1) Then
                
                    '------------------------------------------------------------------------------
                    '  LOG de acciones
                    Set LOG = New cLOG
                    LOG.InsertarCambiosRegistros 14, vUsu, "Cambio Recibos Pozos-U", "rrecibpozos", ObtenerWhereCab(False)
                    Set LOG = Nothing
                    '-----------------------------------------------------------------------------
                
                    TerminaBloquear
                    PosicionarData
                End If
            End If
        ' *** si n'hi han ll�nies ***
        
        Case 5 'LL�NIES
            Select Case ModificaLineas
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


Private Sub cmdCancelar_Click()
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
            Select Case ModificaLineas
                Case 1 'afegir ll�nia
                    ModificaLineas = 0
                    ' *** les ll�nies que tenen datagrid (en o sense tab) ***
                    If NumTabMto = 0 Or NumTabMto = 1 Or NumTabMto = 2 Or NumTabMto = 4 Then
                        DataGridAux(NumTabMto).AllowAddNew = False
                        ' **** repasar si es diu Data1 l'adodc de la cap�alera ***
                        'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
                        ' ********************************************************
                        LLamaLineas NumTabMto, ModificaLineas 'ocultar txtAux
                        DataGridAux(NumTabMto).Enabled = True
                        DataGridAux(NumTabMto).SetFocus

                        ' *** si n'hi han camps de descripci� dins del grid, els neteje ***
                        'txtAux2(2).text = ""
                        ' *****************************************************************
                    End If
                    
                    If Not Adoaux(NumTabMto).Recordset.EOF Then
                        Adoaux(NumTabMto).Recordset.MoveFirst
                    End If

                Case 2 'modificar ll�nies
                    ModificaLineas = 0
                    
                    
                    PonerModo 4
                    If Not Adoaux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de ll�nies ***
                        V = Adoaux(NumTabMto).Recordset.Fields(1) 'el 2 es el n� de llinia
                        Adoaux(NumTabMto).Recordset.Find (Adoaux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
                        ' ***************************************************************
                    End If

                    ' ***  bloquejar els camps fora dels grids ***

                    ' ***  bloquejar els camps fora dels grids ***
                    LLamaLineas NumTabMto, ModificaLineas 'ocultar txtAux
        
        
            End Select
            TerminaBloquear

            PosicionarData
            
        
    End Select
End Sub

Private Sub BotonAnyadir()
Dim vSeccion As CSeccion
Dim B As Boolean

    LimpiarCampos 'Huida els TextBox
    PonerModo 3
    
    '08/09/2010: numlinea
    Text1(31).Text = "1"
    
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la cap�alera *******
'    Text1(0).Text = SugerirCodigoSiguienteStr("palets", "numpalet")
'    FormateaCampo Text1(0)
    Combo1(0).ListIndex = -1
    
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    
    Text1(5).Text = 0
    Text1(6).Text = 0
    Text1(7).Text = 0
    Text1(3).Text = vParamAplic.CodIvaPOZ
    
    
    
    Set vSeccion = New CSeccion
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        B = vSeccion.AbrirConta
        If B Then
            Text2(3).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(3).Text, "N")
            Text1(4).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", Text1(3).Text, "N")
            FormateaCampo Text1(4)
        End If
    End If
    Set vSeccion = Nothing
    
    Combo1(0).ListIndex = 0
    Combo1(0).SetFocus
'    PonerFoco Text1(1) '*** 1r camp visible que siga PK ***
    ' *** si n'hi han camps de descripci� a la cap�alera ***
    'PosarDescripcions
End Sub


Private Sub BotonBuscar()
Dim anc As Single

    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        PonerModo 1
        
        'Si pasamos el control aqui lo ponemos en amarillo
        Combo1(0).SetFocus
        Combo1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
'            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
    If chkVistaPrevia.Value = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia ""
    Else
        LimpiarCampos
        CadenaConsulta = "Select rrecibpozos.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla
'        CadenaConsulta = CadenaConsulta & " WHERE scafac.codtipom='" & CodTipoMov & "'"
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index, True
    PonerCampos
End Sub


Private Sub BotonModificar()
Dim DeVarios As Boolean

'    'solo se puede modificar la factura si no esta contabilizada
'    If FactContabilizada Then
'        TerminaBloquear
'        Exit Sub
'    End If
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    CargarValoresAnteriores Me, 1
    
    PonerFoco Text1(14) '*** 1r camp visible que siga PK ***
        
End Sub



Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Facturas (scafac)
' y los registros correspondientes de las tablas cab. albaranes (scafac1)
' y las lineas de la factura (slifac)
Dim cad As String
'Dim vTipoMov As CTiposMov

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    
    Select Case Data1.Recordset.Fields(0).Value
        Case "RCP"
            cad = "Recibo de Consumo." & vbCrLf
    
        Case "RMP"
            cad = "Recibo de Mantenimiento." & vbCrLf
            
        Case "TAL"
            cad = "Recibo de Talla." & vbCrLf
        
        Case "RVP"
            cad = "Recibo de Contadores." & vbCrLf
        
    End Select
    
    cad = cad & "-------------------------------------      " & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar el Recibo del Socio:            "
    cad = cad & vbCrLf & "N� Recibo:  " & Format(Text1(0).Text, "0000000")
    cad = cad & vbCrLf & "Fecha:  " & Format(Text1(1).Text, "dd/mm/yyyy")

    cad = cad & vbCrLf & vbCrLf & " �Desea Eliminarlo? "

    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
'        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
'        NumPedElim = Data1.Recordset.Fields(1).Value
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            PonerModo 0
        End If
        
'        'Devolvemos contador, si no estamos actualizando
'        Set vTipoMov = New CTiposMov
'        vTipoMov.DevolverContador CodTipoMov, NumPedElim
'        Set vTipoMov = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEliminar:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminar Factura", Err.Description
End Sub


Private Sub cmdConceptos_Click()
    If Text1(20).Text = "RVP" Then
        FrameConceptos.visible = Not FrameConceptos.visible
        FrameConceptos.Enabled = Not FrameConceptos.Enabled
        FrameHidrantes.visible = Not FrameConceptos.visible
        FrameHidrantes.Enabled = Not FrameConceptos.visible
        FrameParticipaciones.visible = Not FrameConceptos.visible
        FrameParticipaciones.Enabled = Not FrameConceptos.visible
        FrameCampos.visible = False
        FrameCampos.Enabled = False
        Frame3.visible = Not FrameConceptos.visible
        Frame3.Enabled = Not FrameConceptos.visible
        Frame4.visible = Not FrameConceptos.visible
        Frame4.Enabled = Not FrameConceptos.visible
        Frame6.visible = Not FrameConceptos.visible
        Frame6.Enabled = Not FrameConceptos.visible
        
        If Me.cmdConceptos.Picture = frmPpal.imgListPpal.ListImages(24).Picture Then
            Me.cmdConceptos.Picture = frmPpal.imgListPpal.ListImages(36).Picture
        Else
            Me.cmdConceptos.Picture = frmPpal.imgListPpal.ListImages(24).Picture
        End If
    Else
        FrameConceptos.visible = False
        FrameConceptos.Enabled = False
        FrameHidrantes.visible = False
        FrameHidrantes.Enabled = False
        FrameCampos.visible = False
        FrameCampos.Enabled = False
        FrameParticipaciones.visible = False
        FrameParticipaciones.Enabled = False
        Frame3.visible = Not FrameConceptos.visible
        Frame3.Enabled = Not FrameConceptos.visible
        Frame4.visible = Not FrameConceptos.visible
        Frame4.Enabled = Not FrameConceptos.visible
        Frame6.visible = Not FrameConceptos.visible
        Frame6.Enabled = Not FrameConceptos.visible
       
        Me.cmdConceptos.Picture = frmPpal.imgListPpal.ListImages(24).Picture
    End If
End Sub

Private Sub cmdHidrantes_Click()
    If Text1(20).Text = "RMP" Then
        FrameHidrantes.visible = Not FrameHidrantes.visible
        FrameHidrantes.Enabled = FrameHidrantes.visible
        FrameCampos.visible = False
        FrameCampos.Enabled = False
        FrameConceptos.visible = False
        FrameConceptos.Enabled = False
        FrameParticipaciones.visible = False
        FrameParticipaciones.Enabled = False
        Frame3.visible = Not FrameHidrantes.visible
        Frame3.Enabled = Not FrameHidrantes.visible
        Frame4.visible = Not FrameHidrantes.visible
        Frame4.Enabled = Not FrameHidrantes.visible
        Frame6.visible = Not FrameHidrantes.visible
        Frame6.Enabled = Not FrameHidrantes.visible
        
        If Me.cmdHidrantes.Picture = frmPpal.imgListPpal.ListImages(24).Picture Then
            Me.cmdHidrantes.Picture = frmPpal.imgListPpal.ListImages(36).Picture
        Else
            Me.cmdHidrantes.Picture = frmPpal.imgListPpal.ListImages(24).Picture
        End If
    Else
        FrameHidrantes.visible = False
        FrameHidrantes.Enabled = False
        FrameCampos.visible = False
        FrameCampos.Enabled = False
        FrameConceptos.visible = False
        FrameConceptos.Enabled = False
        FrameParticipaciones.visible = False
        FrameParticipaciones.Enabled = False
        Frame3.visible = Not FrameHidrantes.visible
        Frame3.Enabled = Not FrameHidrantes.visible
        Frame4.visible = Not FrameHidrantes.visible
        Frame4.Enabled = Not FrameHidrantes.visible
        Frame6.visible = Not FrameHidrantes.visible
        Frame6.Enabled = Not FrameHidrantes.visible
       
        Me.cmdHidrantes.Picture = frmPpal.imgListPpal.ListImages(24).Picture
    End If

End Sub


Private Sub cmdCampos_Click()
    '[Monica]05/05/2014: a�adimos los recibos de consumo de manta
    If Text1(20).Text = "TAL" Or Text1(20).Text = "RMT" Then
        FrameCampos.visible = Not FrameCampos.visible
        FrameCampos.Enabled = FrameCampos.visible
        FrameHidrantes.visible = False
        FrameHidrantes.Enabled = False
        FrameConceptos.visible = False
        FrameConceptos.Enabled = False
        FrameParticipaciones.visible = False
        FrameParticipaciones.Enabled = False
        Frame3.visible = Not FrameHidrantes.visible
        Frame3.Enabled = Not FrameHidrantes.visible
        Frame4.visible = Not FrameHidrantes.visible
        Frame4.Enabled = Not FrameHidrantes.visible
        Frame6.visible = Not FrameHidrantes.visible
        Frame6.Enabled = Not FrameHidrantes.visible
        
        If Me.cmdCampos.Picture = frmPpal.imgListPpal.ListImages(24).Picture Then
            Me.cmdCampos.Picture = frmPpal.imgListPpal.ListImages(36).Picture
        Else
            Me.cmdCampos.Picture = frmPpal.imgListPpal.ListImages(24).Picture
        End If
    Else
        FrameCampos.visible = False
        FrameCampos.Enabled = False
        FrameHidrantes.visible = False
        FrameHidrantes.Enabled = False
        FrameConceptos.visible = False
        FrameConceptos.Enabled = False
        FrameParticipaciones.visible = False
        FrameParticipaciones.Enabled = False
        Frame3.visible = Not FrameHidrantes.visible
        Frame3.Enabled = Not FrameHidrantes.visible
        Frame4.visible = Not FrameHidrantes.visible
        Frame4.Enabled = Not FrameHidrantes.visible
        Frame6.visible = Not FrameHidrantes.visible
        Frame6.Enabled = Not FrameHidrantes.visible
       
        Me.cmdCampos.Picture = frmPpal.imgListPpal.ListImages(24).Picture
    End If

End Sub



Private Sub cmdParticipa_Click()
    If Text1(20).Text = "RCP" Then
        FrameParticipaciones.visible = Not FrameParticipaciones.visible
        FrameParticipaciones.Enabled = FrameParticipaciones.visible
        FrameConceptos.visible = False
        FrameConceptos.Enabled = False
        FrameCampos.visible = False
        FrameCampos.Enabled = False
        Frame3.visible = Not FrameParticipaciones.visible
        Frame3.Enabled = Not FrameParticipaciones.visible
        Frame4.visible = Not FrameParticipaciones.visible
        Frame4.Enabled = Not FrameParticipaciones.visible
        Frame6.visible = Not FrameParticipaciones.visible
        Frame6.Enabled = Not FrameParticipaciones.visible
        
        If Me.cmdParticipa.Picture = frmPpal.imgListPpal.ListImages(24).Picture Then
            Me.cmdParticipa.Picture = frmPpal.imgListPpal.ListImages(36).Picture
        Else
            Me.cmdParticipa.Picture = frmPpal.imgListPpal.ListImages(24).Picture
        End If
    Else
        FrameParticipaciones.visible = False
        FrameParticipaciones.Enabled = False
        FrameConceptos.visible = False
        FrameConceptos.Enabled = False
        FrameCampos.visible = False
        FrameCampos.Enabled = False
        Frame3.visible = Not FrameParticipaciones.visible
        Frame3.Enabled = Not FrameParticipaciones.visible
        Frame4.visible = Not FrameParticipaciones.visible
        Frame4.Enabled = Not FrameParticipaciones.visible
        Frame6.visible = Not FrameParticipaciones.visible
        Frame6.Enabled = Not FrameParticipaciones.visible
       
        Me.cmdParticipa.Picture = frmPpal.imgListPpal.ListImages(24).Picture
    End If
End Sub





Private Sub cmdRectificativa_Click()
    Me.FrameRectifica.visible = Not (Me.FrameRectifica.visible)
    Me.Frame5(0).visible = Not Me.FrameRectifica.visible
End Sub

Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        If Not Data1.Recordset.EOF Then _
            Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

    Else 'Se llama desde alg�n Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ning�n registro devuelto.", vbExclamation
            Exit Sub
        End If
        cad = Data1.Recordset.Fields(0) & "|"
        cad = cad & Data1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub

Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbYellow
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim I As Integer
    If Combo1(Index).BackColor = vbYellow Then Combo1(Index).BackColor = vbWhite
    
    I = Combo1(Index).ListIndex
    Text1(20).Text = Mid(Trim(Combo1(Index).List(I)), 1, 3)
    CodTipoMov = Text1(20).Text
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    If PrimeraVez Then
        PonerFocoChk Me.chkVistaPrevia
        PrimeraVez = False
    End If
    'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
    If hcoCodMovim <> "" And Not Data1.Recordset.EOF Then PonerCadenaBusqueda
    
'    PonerCadenaBusqueda
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim I As Integer
    'Icono del formulario
    Me.Icon = frmPpal.Icon

     'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next kCampo

    
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
     
    ' ******* si n'hi han ll�nies *******
    'ICONETS DE LES BARRES ALS TABS DE LL�NIA
    For kCampo = 0 To 2
        With Me.ToolAux(kCampo)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next kCampo
   ' ***********************************
    
    Me.cmdConceptos.Picture = frmPpal.imgListPpal.ListImages(24).Picture
    Me.cmdParticipa.Picture = frmPpal.imgListPpal.ListImages(24).Picture
    Me.cmdHidrantes.Picture = frmPpal.imgListPpal.ListImages(24).Picture
    Me.cmdCampos.Picture = frmPpal.imgListPpal.ListImages(24).Picture
    Me.cmdRectificativa.Picture = frmPpal.imgListPpal.ListImages(37).Picture

    Me.FrameConceptos.visible = False
    Me.FrameConceptos.Enabled = False
    Me.cmdConceptos.visible = False
    Me.cmdConceptos.Enabled = False

    Me.FrameParticipaciones.visible = False
    Me.FrameParticipaciones.Enabled = False
    Me.cmdParticipa.visible = False
    Me.cmdParticipa.Enabled = False
    
    
    Me.FrameHidrantes.visible = False
    Me.FrameHidrantes.Enabled = False
    Me.cmdHidrantes.visible = False
    Me.cmdHidrantes.Enabled = False
    
    Me.FrameCampos.visible = False
    Me.FrameCampos.Enabled = False
    Me.cmdCampos.visible = False
    Me.cmdCampos.Enabled = False
    
    Me.FrameRectifica.visible = False
'    Me.FrameRectifica.Enabled = False
    
    
    
    LimpiarCampos   'Limpia los campos TextBox
    CargaCombo
    
    CodTipoMov = "RCP"
    VieneDeBuscar = False
    
    '[Monica]08/05/2012: a�adida Escalona que funciona como Utxera
    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        Label28.Caption = "Consumo m3."
        Label13.Caption = "Consumo m3."
        Label27.Caption = "Precio 1"
        Label14.Caption = "Precio 2"
    End If
    
    '[Monica]20/07/2015: solo para Escalona
    If vParamAplic.Cooperativa = 10 Then
        Me.Caption = "Duplicado de Recibos"
    End If
    
    
    '## A mano
    NombreTabla = "rrecibpozos"
    Ordenacion = " ORDER BY rrecibpozos.numfactu"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    CadenaConsulta = "select * from rrecibpozos "
    If hcoCodMovim <> "" Then
        'Se llama desde Dobleclick en frmManSocios
        CadenaConsulta = CadenaConsulta & " WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
    Else
        CadenaConsulta = CadenaConsulta & " where numfactu=-1"
    End If
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    'Poner los grid sin apuntar a nada
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1 'b�squeda
        ' *** posar de groc els camps visibles de la clau primaria de la cap�alera ***
        Text1(0).BackColor = vbYellow 'codclien
        ' ****************************************************************************
    End If

End Sub


Private Sub LimpiarCampos()
Dim I As Integer

    On Error Resume Next
    
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Combo1(0).ListIndex = -1
    For I = 0 To Check1.Count - 1
        Me.Check1(I).Value = 0
    Next I
'    Label2(2).Caption = ""
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Modo = 4 Then TerminaBloquear
End Sub



Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(20), CadenaDevuelta, 1)
        CadB = CadB & Aux
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 2)
        CadB = CadB & " and " & Aux
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 3)
        CadB = CadB & " and " & Aux
        Aux = ValorDevueltoFormGrid(Text1(31), CadenaDevuelta, 4)
        CadB = CadB & " and " & Aux
        
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB '& " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    Text1(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub


Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
' hidrante del socio
    Text1(14).Text = RecuperaValor(CadenaSeleccion, 1)
End Sub

Private Sub frmPOZRecPrev_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
    
    If CadenaSeleccion <> "" Then
        CadB = "codtipom = " & DBSet(RecuperaValor(CadenaSeleccion, 1), "T")
        CadB = CadB & " and numfactu = " & DBSet(RecuperaValor(CadenaSeleccion, 2), "N")
        CadB = CadB & " and fecfactu = " & DBSet(RecuperaValor(CadenaSeleccion, 3), "F")
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub frmTIva_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento Tipo de iva
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codiva
    FormateaCampo Text1(3)
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'nomiva
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 3) 'porcentaje iva
    FormateaCampo Text1(4)
End Sub


Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Socios
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000") 'Cod Socios
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Descripcion
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(Indice).Text = vCampo
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim vSeccion As CSeccion

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 1 'Tipo de IVA
            Indice = 3
            PonerFoco Text1(Indice)
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
                If vSeccion.AbrirConta Then
                    Set frmTIva = New frmTipIVAConta
                    frmTIva.DeConsulta = True
                    frmTIva.DatosADevolverBusqueda = "0|1|2|"
                    frmTIva.CodigoActual = Text1(3).Text
                    frmTIva.Show vbModal
                    Set frmTIva = Nothing
                    PonerFoco Text1(3)
                End If
            End If
            Set vSeccion = Nothing
        
        Case 0 'Socios
            Indice = 2
            PonerFoco Text1(Indice)
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco Text1(Indice)
            
            
    End Select
    
    Screen.MousePointer = vbDefault
End Sub



Private Sub imgFec_Click(Index As Integer)
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
    
    menu = Me.Height - Me.ScaleHeight 'ac� tinc el heigth del men� i de la toolbar

    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40

    Select Case Index
        Case 0
            Indice = 1
        Case 1
            Indice = 10
        Case 2
            Indice = 11
        Case 3
            Indice = 40
        Case 4
            Indice = 42
        Case 5
            Indice = 43
        Case 6
            Indice = 47
    End Select
    
    imgFec(0).Tag = Indice '<===
    If Text1(Indice).Text <> "" Then frmC.NovaData = Text1(Indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
     PonerFoco Text1(CByte(imgFec(0).Tag)) '<===
    ' ********************************************
End Sub


Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        Indice = 15
        frmZ.pTitulo = "Observaciones del Albar�n"
        frmZ.pValor = Text1(Indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco Text1(Indice)
    End If
End Sub


Private Sub mnBuscar_Click()
    BotonBuscar
End Sub


Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
'         BotonEliminarLinea
    Else   'Eliminar Pedido
         BotonEliminar
    End If
End Sub


Private Sub mnImprimir_Click()
'Imprimir Factura
    
    If Data1.Recordset.EOF Then Exit Sub
    
    BotonImprimir
End Sub



Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub


Private Sub mnModificar_Click()

'    If Data1.Recordset!impreso = 1 Then
'        If MsgBox("Este albar�n est� facturado y/o cobrado. � Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
'            Exit Sub
'        End If
'    End If

    'bloquea la tabla cabecera de factura: scafac
    If BLOQUEADesdeFormulario(Me) Then
        'bloquear la tabla cabecera de albaranes de la factura: scafac1
        BotonModificar
    End If

End Sub


Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    If (Modo = 5) Then 'Modo 5: Mto Lineas
        '1:Insertar linea, 2: Modificar
        If ModificaLineas = 1 Or ModificaLineas = 2 Then cmdCancelar_Click
        cmdRegresar_Click
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


Private Sub Text1_Change(Index As Integer)
'    If Index = 9 Then HaCambiadoCP = True 'Cod. Postal
End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
'    If Index = 9 Then HaCambiadoCP = False 'CPostal
'    If Index = 1 And Modo = 1 Then
'        SendKeys "{tab}"
'        Exit Sub
'    End If
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)
Dim devuelve As String
Dim cadMen As String
Dim SQL As String
Dim vSeccion As CSeccion
Dim vSocio As cSocio
Dim Rs As ADODB.Recordset

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1 'Fecha factura
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
    
        Case 2 'Socio
            If Modo = 1 Then Exit Sub
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rsocios", "nomsocio")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Socio: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmSoc = New frmManSocios
                        frmSoc.DatosADevolverBusqueda = "0|1|"
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
                        PonerHidrantesSocio
                    Else
                        MsgBox "El socio est� dado de baja. Reintroduzca.", vbExclamation
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    End If
                End If
            End If
            
        Case 3 'Tipo de IVA
            If Text1(Index).Text <> "" Then
                Set vSeccion = New CSeccion
                If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
                    If vSeccion.AbrirConta Then
                        Text2(3).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(3).Text, "N")
                        Text1(4).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", Text1(3).Text, "N")
                        
                        CalculoTotales
                    End If
                End If
                Set vSeccion = Nothing
            End If
            
        Case 4, 5 ' base imponible, importe iva, total factura
            If PonerFormatoDecimal(Text1(Index), 3) Then
                CalculoTotales
            End If
            
        Case 6 ' importe de iva
            PonerFormatoDecimal Text1(Index), 3
            
        Case 12 'consumo
            PonerFormatoEntero Text1(Index)
            
        Case 13 'cuota
            If Modo <> 1 Then PonerFormatoDecimal Text1(Index), 3
        
        
        Case 14 ' contador para el caso de escalona
            If Modo = 1 Then Exit Sub
            If Text1(Index).Text <> "" Then
                If Modo = 3 Or Modo = 4 Then
                    SQL = "select poligono, parcelas, nroorden from rpozos where hidrante = " & DBSet(Text1(Index).Text, "T")
                    Set Rs = New ADODB.Recordset
                    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    If Not Rs.EOF Then
                        Text1(36).Text = DBLet(Rs!Poligono, "T")
                        Text1(37).Text = DBLet(Rs!parcelas, "T")
                        Text1(38).Text = DBLet(Rs!nroorden, "N")
                    End If
                End If
            End If
        
        Case 8, 9, 16, 18, 41, 44 'contadores
            PonerFormatoEntero Text1(Index)
            
        Case 10, 11, 42, 43, 47 'fechas
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
                        
        Case 17, 19 ' precios aplicados
            If Modo <> 1 Then PonerFormatoDecimal Text1(Index), 5
            
        Case 32 ' diferencia de dias
            PonerFormatoEntero Text1(32)
            
        Case 33 ' Porcentaje de bonificacion / Recargo
            If Modo <> 1 Then PonerFormatoDecimal Text1(Index), 4
            
        Case 34 ' Importe de bonificacion / Recargo
            If Modo <> 1 Then PonerFormatoDecimal Text1(Index), 3
        
        Case 35 ' Precio
            If Modo <> 1 Then PonerFormatoDecimal Text1(Index), 11
    End Select
End Sub


Private Sub HacerBusqueda()
Dim CadB As String
Dim cadAux As String
    
'    '--- Laura 12/01/2007
'    cadAux = Text1(5).Text
'    If Text1(4).Text <> "" Then Text1(5).Text = ""
'    '---
    
'    '--- Laura 12/01/2007
'    Text1(5).Text = cadAux
'    '---
'    CadB = ObtenerBusqueda(Me)
    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)

    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select rrecibpozos.* from " & NombreTabla
        CadenaConsulta = CadenaConsulta & " WHERE " & CadB & Ordenacion
'        CadenaConsulta = "select palets.* from " & NombreTabla
'        CadenaConsulta = CadenaConsulta & " WHERE " & CadB & " GROUP BY palets.numpalet " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
'    'Llamamos a al form
'    '##A mano
'    cad = ""
''    Cad = Cad & "Tipo|if(rfactsoc.codtipom='FAA','Anticipo','Liquidaci�n') as a|T||10�"
'    cad = cad & "Tipo Fichero|case rrecibpozos.codtipom when ""RCP"" then ""RCP-Consumo"" when ""RMP"" then ""RMP-Mantenim"" when ""RVP"" then ""RVP-Contadores"" when ""RMT"" then ""RMT-Manta"" when ""TAL"" then ""TAL-Talla"" when ""RRC"" then ""RRC-Rect Consumo""  when ""RRM"" then ""RRM-Rect Mto"" when ""RRT"" then ""RRT-Rect Manta""  when ""RRV"" then ""RRV-Rect Cont"" when ""RTA"" then ""RTA-Rect Talla""  when ""FIN"" then ""FIN-Internas"" end as tipo|N||22�"
'    cad = cad & "Tipo|rrecibpozos.codtipom|N||6�" ' ParaGrid(Combo1(0), 0, "Tipo")
'    cad = cad & "N�.Factura|rrecibpozos.numfactu|N||12�"
'    cad = cad & "Fecha|rrecibpozos.fecfactu|F||15�"
'    cad = cad & "Lin|rrecibpozos.numlinea|N||6�"
'    cad = cad & "C�digo|rrecibpozos.codsocio|N|000000|12�"
'    cad = cad & "Socio|rsocios.nomsocio|N||38�"
'
'    tabla = NombreTabla & " inner join rsocios on rrecibpozos.codsocio = rsocios.codsocio "
'    Titulo = "Recibos de Contadores"
'    devuelve = "1|2|3|4|"
'
'    If cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vtabla = tabla
'        frmB.vSQL = CadB
'        HaDevueltoDatos = False
'        '###A mano
'        frmB.vDevuelve = devuelve
'        frmB.vTitulo = Titulo
'        frmB.vSelElem = 0
''        frmB.vConexionGrid = cAgro  'Conexi�n a BD: Ariagro
'        If Not EsCabecera Then frmB.Label1.FontSize = 11
''        frmB.vBuscaPrevia = chkVistaPrevia
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
''        If EsCabecera Then
''            PonerCadenaBusqueda
''            Text1(0).Text = Format(Text1(0).Text, "0000000")
''        End If
'        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
'        End If
'    End If
'    Screen.MousePointer = vbDefault

    Set frmPOZRecPrev = New frmPOZRecibosMonastPrev
    
    frmPOZRecPrev.cWhere = CadB
    frmPOZRecPrev.DatosADevolverBusqueda = "1|2|3|"
    frmPOZRecPrev.Show vbModal
    
    Set frmPOZRecPrev = Nothing


End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass

    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then
            PonerFoco Text1(kCampo)
'            Text1(0).BackColor = vbYellow
        End If
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        PonerCampos
    End If


    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim BrutoFac As Single
Dim B As Boolean
Dim vSeccion As CSeccion

    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1

    PosicionarCombo2 Combo1(0), Text1(20).Text

    CargaGrid 2, True
    If Not Adoaux(2).Recordset.EOF Then _
        PonerCamposForma2 Me, Adoaux(2), 2, "FrameAux2"
    
    CargaGrid 0, True
    If Not Adoaux(0).Recordset.EOF Then _
        PonerCamposForma2 Me, Adoaux(0), 2, "FrameAux0"
    
    CargaGrid 1, True
    If Not Adoaux(1).Recordset.EOF Then _
        PonerCamposForma2 Me, Adoaux(1), 2, "FrameAux1"
    
    
    
'    cmdConceptos_Click
    
    cmdConceptos.visible = (Text1(20).Text = "RVP")
    cmdConceptos.Enabled = (Text1(20).Text = "RVP")
  
   
    cmdParticipa.visible = (Text1(20).Text = "RCP") And vParamAplic.Cooperativa = 1
    cmdParticipa.Enabled = (Text1(20).Text = "RCP") And vParamAplic.Cooperativa = 1
   
    cmdHidrantes.visible = (Text1(20).Text = "RMP") And vParamAplic.Cooperativa = 10
    cmdHidrantes.Enabled = (Text1(20).Text = "RMP") And vParamAplic.Cooperativa = 10
    
    cmdCampos.visible = (Text1(20).Text = "TAL" Or Text1(20).Text = "RMT") And vParamAplic.Cooperativa = 10
    cmdCampos.Enabled = (Text1(20).Text = "TAL" Or Text1(20).Text = "RMT") And vParamAplic.Cooperativa = 10
   
   
    '[Monica]15/09/2015: si es Escalona indicamos si el recibo est� pagado
    If vParamAplic.Cooperativa = 10 Then
        If ReciboCobrado(Mid(Combo1(0).Text, 1, 3), Text1(0).Text, Text1(1).Text) Then
            Label1(13).visible = True
            Label1(13).Caption = "P A G A D O"
'            Timer1.Enabled = False
        Else
            Label1(13).Caption = "R E C I B O  P E N D I E N T E"
'            Timer1.Enabled = True
        End If
    End If
            

'    FormatoDatosTotales
    
    'poner descripcion campos
    Modo = 4
    
    Text2(2).Text = PonerNombreDeCod(Text1(2), "rsocios", "nomsocio", "codsocio", "N") 'socios
    
    Set vSeccion = New CSeccion
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        B = vSeccion.AbrirConta
        If B Then
            Text2(3).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(3).Text, "N")
        End If
    End If
    Set vSeccion = Nothing
'    MostrarCadena Text1(3), Text1(4)
    
    Modo = 2
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu seg�n modo
    PonerOpcionesMenu 'Activar opciones de menu seg�n nivel de permisos del usuario

    If Err.Number <> 0 Then Err.Clear
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim I As Byte, NumReg As Byte
Dim B As Boolean
Dim b1 As Boolean

    On Error GoTo EPonerModo

    BuscaChekc = ""

    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    B = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Or hcoCodMovim <> "" Then
        cmdRegresar.visible = B
    Else
        cmdRegresar.visible = False
    End If
    
    
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    DesplazamientoVisible B And Data1.Recordset.RecordCount > 1
          
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    BloquearCombo Me, Modo
    
    '+++ bloqueamos el combo1(0) como si tuviera tag
    b1 = (Modo = 3 Or Modo = 4 Or Modo = 1) '06/09/2005, lleve el modo 5 per a que no es puga modificar la cap�alera mentre treballe en les ll�nies
    
    If (Modo = 4 Or Modo = 5) Then
        Combo1(0).Enabled = False
        Combo1(0).BackColor = &H80000018 'groc
    Else
        Combo1(0).Enabled = b1
        If b1 Then
            Combo1(0).BackColor = vbWhite
        Else
            Combo1(0).BackColor = &H80000018 'Amarillo Claro
        End If
        If Modo = 3 Then Combo1(0).ListIndex = 0 'Modo 3: Insertar (si vamos a Insertar ade+ Limpiamos el campo)
    End If
    '+++

    Text1(0).Enabled = (Modo = 1)
    
    For I = 0 To Check1.Count - 1
        If I <> 3 Then
            Me.Check1(I).Enabled = (Modo = 1)
        Else
            Me.Check1(I).Enabled = (Modo = 1 Or Modo = 3 Or Modo = 4)
        End If
    Next I
    
    B = (Modo <> 1)
    'Campos N� Recibo bloqueado y en azul
    BloquearTxt Text1(0), B, True
    
    '---------------------------------------------
    B = (Modo <> 0 And Modo <> 2)
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    
    BloquearImgBuscar Me, Modo, ModificaLineas
    BloquearImgFec Me, 0, Modo
    
    B = (Modo = 1 Or Modo = 3)
    Text1(1).Enabled = B
    Text1(2).Enabled = B Or Modo = 4
    imgFec(0).Enabled = B
    imgFec(0).visible = B
    imgBuscar(0).Enabled = B Or Modo = 4
    imgBuscar(0).visible = B Or Modo = 4
    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    BloquearTxt Text1(4), (Modo <> 1)
    BloquearTxt Text1(6), (Modo <> 1)
       
    For I = 1 To 2
        imgFec(I).Enabled = (Modo = 1 Or Modo = 3 Or Modo = 4)
    Next I
        
    ' ***************************
    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
'lineas
    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 2, False
        CargaGrid 0, False
        CargaGrid 1, False
    End If
    
    B = (Modo = 4) Or (Modo = 2)
    DataGridAux(2).Enabled = B
    
    B = (Modo = 5)
    For I = 1 To 3
        BloquearTxt txtAux3(I), Not B
    Next I
    B = (Modo = 5) And ModificaLineas = 2
    BloquearTxt txtAux3(1), B
    
    B = (Modo = 5)
    For I = 4 To 10
        BloquearTxt txtAux5(I), Not B
    Next I
    
    '[Monica]15/09/2015: ponemos la situacion
    Label1(13).visible = ((Modo = 2 Or Modo = 4) And vParamAplic.Cooperativa = 10)
    
    
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu seg�n modo
    PonerOpcionesMenu 'Activar opciones de menu seg�n nivel de permisos del usuario
    
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
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOK() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim B As Boolean
Dim SQL As String

    On Error GoTo EDatosOK

    DatosOK = False
    
'    ComprobarDatosTotales

    'comprobamos datos OK de la tabla scafac
    B = CompForm2(Me, 1)  'Comprobar formato datos ok de la cabecera: opcion=1
    If Not B Then Exit Function
    
    If B Then
        If (Modo = 3 Or Modo = 4) And Combo1(0).ListIndex = 0 Then
            '[Monica]17/11/2014: obligamos a que si es de consumo metan el hidrante s�lo para escalona y utxera
            If Text1(14).Text = "" Then
                If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                    MsgBox "El hidrante no puede estar vacio en un recibo de consumo. Revise.", vbExclamation
                    PonerFoco Text1(14)
                    B = False
                End If
            Else
            
                '[Monica]17/11/2014: si el hidrante no existe evitamos el error de clave referencial
                If B Then
                    SQL = DevuelveDesdeBDNew(cAgro, "rpozos", "hidrante", "hidrante", Text1(14).Text, "T")
                    If SQL = "" Then
                        MsgBox "El Hidrante no existe. Revise.", vbExclamation
                        PonerFoco Text1(14)
                        B = False
                    End If
                End If
            
            
                ' comprobamos si insertamos o modificamos que existe el hidrante para el socio
                If B Then
                    SQL = ""
                    SQL = DevuelveDesdeBDNew(cAgro, "rpozos", "hidrante", "hidrante", Text1(14).Text, "T", , "codsocio", Text1(2).Text, "N")
                    
                    If SQL = "" Then
                        If MsgBox("El Hidrante no es del socio introducido. " & vbCrLf & vbCrLf & "� Desea continuar ? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                            PonerFoco Text1(14)
                            B = False
                        Else
                            B = True
                        End If
                    End If
                End If
                    
            End If
        End If
    End If
    
    If B Then
        If Modo = 3 Or Modo = 4 Then
            If Text1(11).Text <> "" And Text1(10).Text <> "" Then
                If CDate(Text1(11).Text) > CDate(Text1(10).Text) Then
                    MsgBox "La Fecha de Lectura Anterior no puede ser superior a la de Lectura Actual. Revise.", vbExclamation
                    PonerFoco Text1(11)
                    B = False
                End If
            End If
            
            If B Then
                If Text1(8).Text <> "" And Text1(9).Text <> "" Then
                    If CLng(Text1(8).Text) > CLng(Text1(9).Text) Then
                        MsgBox "El Contador Anterior no puede ser superior al del Contador Actual. Revise.", vbExclamation
                        PonerFoco Text1(8)
                        B = False
                    End If
                End If
            End If
        End If
    End If
    
    
    If B Then
        If Modo = 3 Then
            '[Monica]20/06/2017: control de fechas que antes no estaba
            ResultadoFechaContaOK = EsFechaOKConta(CDate(Text1(1)))
            If ResultadoFechaContaOK > 0 Then
                If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
                B = False
            End If
        End If
    End If
    
    DatosOK = B
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 5  'Buscar
            mnBuscar_Click
        Case 6  'Todos
            BotonVerTodos
        Case 1  'A�adir
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
        Case 8  ' Impresion de albaran
            mnImprimir_Click
    End Select
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

Private Function Eliminar() As Boolean
Dim SQL As String, LEtra As String, Sql2 As String
Dim B As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String

    On Error GoTo FinEliminar

    If Data1.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    'Eliminar en tablas de cabecera de factura
    '------------------------------------------
    SQL = " " & ObtenerWhereCP(True)
    
    conn.Execute "delete from rrecibpozos_acc " & SQL
    
    conn.Execute "delete from rrecibpozos_hid " & SQL
    
    conn.Execute "delete from rrecibpozos_cam " & SQL
    
    
    'Cabecera de factura (rrecibpozos)
    conn.Execute "Delete from " & NombreTabla & SQL
    
    
    CadenaCambio = "DELETE FROM " & NombreTabla & SQL
    
    '------------------------------------------------------------------------------
    '  LOG de acciones
    ValorAnterior = ""
    Set LOG = New cLOG
    LOG.InsertarCambiosRegistros 14, vUsu, "Cambio Recibos Pozos-D", "rrecibpozos", ObtenerWhereCab(False)
    Set LOG = Nothing
    '-----------------------------------------------------------------------------
    
    
    
    
    'Decrementar contador si borramos el ultima factura
    Set vTipoMov = New CTiposMov
    vTipoMov.DevolverContador Text1(20).Text, Val(Text1(0).Text)
    Set vTipoMov = Nothing
    
    B = True
FinEliminar:
    If Err.Number <> 0 Or Not B Then
        MuestraError Err.Number, "Eliminar Recibo", Err.Description & " " & Mens
        B = False
    End If
    If Not B Then
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         If SituarDataMULTI(Data1, vWhere, Indicador) Then
             PonerModo 2
             lblIndicador.Caption = Indicador
        Else
             LimpiarCampos
             'Poner los grid sin apuntar a nada
             PonerModo 0
         End If
    Else
        'El Data esta vacio, desde el modo de inicio se pulsa Insertar
        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim SQL As String

    On Error Resume Next
    
    SQL = " codtipom= '" & Text1(20).Text & "'"
    SQL = SQL & " and numfactu = " & Text1(0).Text
    SQL = SQL & " and fecfactu = " & DBSet(Text1(1).Text, "F")
    '08/09/2010 : a�adido a la clave primaria
    SQL = SQL & " and numlinea = " & DBSet(Text1(31).Text, "N")

    If conWhere Then SQL = " WHERE " & SQL
    ObtenerWhereCP = SQL
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar seg�n el modo en que estemos
Dim B As Boolean, bAux As Boolean
Dim I As Integer


    B = ((Modo = 2) Or (Modo = 0)) And (hcoCodMovim = "") 'Or (Modo = 5 And ModificaLineas = 0)
    'Buscar
    Toolbar1.Buttons(5).Enabled = B
    Me.mnBuscar.Enabled = B
    'Vore Tots
    Toolbar1.Buttons(6).Enabled = B
    Me.mnVerTodos.Enabled = B
    
    'A�adir
    Toolbar1.Buttons(1).Enabled = B
    Me.mnModificar.Enabled = B
    
    B = (Modo = 2 And Data1.Recordset.RecordCount > 0) And (hcoCodMovim = "") And Not (Check1(1).Value = 1)
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    Me.mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(3).Enabled = B
    Me.mnEliminar.Enabled = B
    'Impresi�n de albaran
    Toolbar1.Buttons(8).Enabled = (Modo = 2) And Data1.Recordset.RecordCount > 0
    Me.mnImprimir.Enabled = (Modo = 2) And Data1.Recordset.RecordCount > 0

    ' *** si n'hi han ll�nies que tenen grids (en o sense tab) ***
    B = (Modo = 3 Or Modo = 4 Or Modo = 2) And DatosADevolverBusqueda = "" And Check1(1).Value = 0
    For I = 0 To 2
        ToolAux(I).Buttons(1).Enabled = B
        If B Then bAux = (B And Me.Adoaux(I).Recordset.RecordCount > 0)
        ToolAux(I).Buttons(2).Enabled = bAux
        ToolAux(I).Buttons(3).Enabled = bAux
    Next I
    ' ****************************************

End Sub


Private Sub BotonImprimir()
Dim cadFormula As String
Dim CadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim NroCopias As Integer

    If Text1(0).Text = "" Then
        MsgBox "Debe seleccionar una Factura para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    CadParam = ""
    cadSelect = ""
    numParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    Select Case Text1(20).Text
        Case "RCP"
            indRPT = 46 'Impresion de recibos de consumo de pozos
        Case "RMP"
            indRPT = 47 'Impresion de recibos de mantenimiento de pozos
        Case "RVP"
            indRPT = 47 'Impresion de recibos de contadores pozos
        Case "TAL"
            indRPT = 47 'Impresion de recibos de talla
        Case "RMT"
            indRPT = 47 'Impresion de recibos de consumo a manta
            
        '[Monica]14/01/2016: las rectificativas
        Case "RRC"
            indRPT = 46 ' impresion de recibos de consumo
        Case "RRM"
            indRPT = 47 'Impresion de recibos de mantenimiento de pozos
        Case "RRV"
            indRPT = 47 'Impresion de recibos de contadores pozos
        Case "RTA"
            indRPT = 47 'Impresion de recibos de talla
        Case "RRT"
            indRPT = 47 'Impresion de recibos de consumo a manta
    End Select
    
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
    
    If Text1(20).Text = "TAL" Then nomDocu = Replace(nomDocu, "Mto.", "Tal.")
    If Text1(20).Text = "RVP" Then nomDocu = Replace(nomDocu, "Mto.", "Cont.")
    If Text1(20).Text = "RMT" Then nomDocu = Replace(nomDocu, "Mto.", "Manta.")
      
    '[Monica]14/01/2016: las rectificativas
    If Text1(20).Text = "RTA" Then nomDocu = Replace(nomDocu, "Mto.", "Tal.")
    If Text1(20).Text = "RRV" Then nomDocu = Replace(nomDocu, "Mto.", "Cont.")
    If Text1(20).Text = "RRM" Then nomDocu = Replace(nomDocu, "Mto.", "Manta.")
      
      
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion N� de recibo
    '---------------------------------------------------
    If Text1(0).Text <> "" Then
        'tipo de fichero
        devuelve = "{" & NombreTabla & ".codtipom}='" & Text1(20).Text & "'"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
        
        'N� factura
        devuelve = "{" & NombreTabla & ".numfactu}=" & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        devuelve = "numfactu = " & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
        
        'Fecha Factura
        devuelve = "{" & NombreTabla & ".fecfactu}=Date(" & Year(Text1(1).Text) & "," & Month(Text1(1).Text) & "," & Day(Text1(1).Text) & ")"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        devuelve = "fecfactu = " & DBSet(Text1(1).Text, "F")
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
        
'        'Socio
'        devuelve = "{" & NombreTabla & ".codsocio}=" & Val(Text1(2).Text)
'        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
'        devuelve = "codsocio = " & Val(Text1(2).Text)
'        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
        
    End If
    
    
    ' si no es escalona
    If vParamAplic.Cooperativa = 10 Then
        If ReciboCobrado(Mid(Combo1(0).Text, 1, 3), Text1(0).Text, Text1(1).Text) Then
            CadParam = CadParam & "pDuplicado=1|"
            numParam = numParam + 1
        End If
    End If
    
    If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub
     
    With frmImprimir
          '[Monica]06/02/2012: a�adido la siguientes 3 lineas para el envio por el outlook
            .outClaveNombreArchiv = Mid(Combo1(0).Text, 1, 3) & Format(Text1(0).Text, "0000000")
            .outCodigoCliProv = Text1(2).Text
            .outTipoDocumento = 100
            
            .FormulaSeleccion = cadFormula
            .OtrosParametros = CadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 0
            .Titulo = "Impresi�n de Recibos de Socios"
            
            '[Monica]11/09/2015: pasamos la contabilidad que es pq tenemos que imprimir que gastos de cobros tiene.
            If vParamAplic.Cooperativa = 10 Then
                vParamAplic.NumeroConta = DevuelveValor("Select empresa_conta from rseccion where codsecci = " & vParamAplic.Seccionhorto)
            End If
            .ConSubInforme = True
            .Show vbModal
    End With

    If frmVisReport.EstaImpreso Then
        ActualizarRegistros "rrecibpozos", cadSelect
    End If
End Sub

Private Function ReciboCobrado(TipoM As String, numfactu As String, fecfactu As String) As Boolean
Dim SQL As String
Dim vSeccion As CSeccion
Dim Rs As ADODB.Recordset

    ReciboCobrado = False

    If Check1(1).Value = 0 Then Exit Function
    

    Set vSeccion = New CSeccion
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If vSeccion.AbrirConta Then
    
            SQL = "SELECT count(*) FROM scobro INNER JOIN usuarios.stipom ON scobro.numserie = stipom.letraser "
            SQL = SQL & " WHERE stipom.codtipom = " & DBSet(TipoM, "T")
            SQL = SQL & " and scobro.codfaccl = " & DBSet(numfactu, "N")
            SQL = SQL & " and scobro.fecfaccl = " & DBSet(fecfactu, "F")
            
            Set Rs = New ADODB.Recordset
            Rs.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If Not Rs.EOF Then
                If Rs.Fields(0).Value = 0 Then
                    ReciboCobrado = True
                    Exit Function
                End If
            End If
            Set Rs = Nothing
            
            
            SQL = "SELECT sum(coalesce(impvenci,0) + coalesce(gastos,0) - coalesce(impcobro,0))  FROM scobro INNER JOIN usuarios.stipom ON scobro.numserie = stipom.letraser "
            SQL = SQL & " WHERE stipom.codtipom = " & DBSet(TipoM, "T")
            SQL = SQL & " and scobro.codfaccl = " & DBSet(numfactu, "N")
            SQL = SQL & " and scobro.fecfaccl = " & DBSet(fecfactu, "F")
            Set Rs = New ADODB.Recordset
            Rs.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not Rs.EOF Then
                ReciboCobrado = (DBLet(Rs.Fields(0).Value) = 0)
            End If
            Set Rs = Nothing
            
    
        End If
    End If
    Set vSeccion = Nothing
End Function


Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim I As Byte
    
    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For I = 0 To Combo1.Count - 1
        Combo1(I).Clear
    Next I
    
    'tipo de fichero
    Combo1(0).AddItem "RCP-Consumo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "RMP-Mantenimiento"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "RVP-Contadores"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    If vParamAplic.Cooperativa = 10 Then
        Combo1(0).AddItem "TAL-Talla"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 3
    
        Combo1(0).AddItem "RMT-Consumo Manta"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 4
        
        '[Monica]14/01/2016: las rectificativas
        Combo1(0).AddItem "RRC-Rect.Consumo"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 5
        Combo1(0).AddItem "RRM-Rect.Mantenimiento"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 6
        Combo1(0).AddItem "RRV-Rect.Contadores"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 7
        Combo1(0).AddItem "RTA-Rect.Talla"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 8
'        Combo1(0).AddItem "RRT-Rect.Consumo Manta"
'        Combo1(0).ItemData(Combo1(0).NewIndex) = 9
    End If
    
    '[Monica]02/02/2016: caso de quatretonda tienen internas
    If vParamAplic.Cooperativa = 7 Then
        Combo1(0).AddItem "FIN-Internas"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 3
    End If
    
End Sub


Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim SQL As String

    On Error GoTo EInsertarCab
    
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        SQL = CadenaInsertarDesdeForm(Me)
        If SQL <> "" Then
            If InsertarOferta(SQL, vTipoMov) Then
                CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
            End If
        End If
        Text1(0).Text = Format(Text1(0).Text, "0000000")
    End If
    Set vTipoMov = Nothing
    
EInsertarCab:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Function InsertarOferta(vSQL As String, vTipoMov As CTiposMov) As Boolean
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim cambiaSQL As Boolean
Dim devuelve As String
Dim SQL As String
Dim NumF As String

    On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Factura
    'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
    Do
        '[Monica]12/06/2014: en el caso de escalona no hay cambio de campa�a
        If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
            devuelve = DevuelveDesdeBDNew(cAgro, NombreTabla, "numfactu", "numfactu", Text1(0).Text, "N", , "codtipom", Text1(20).Text, "T", "year(fecfactu)", Mid(Text1(1).Text, 7, 4), "N")
        Else
            devuelve = DevuelveDesdeBDNew(cAgro, NombreTabla, "numfactu", "numfactu", Text1(0).Text, "N", , "codtipom", Text1(20).Text, "T")
        End If
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
    
'    '08/09/2010 : Monica Damos valor al nro de linea
'    Sql = " codtipom= '" & Text1(20).Text & "'"
'    Sql = Sql & " and numfactu = " & Text1(0).Text
'    Sql = Sql & " and fecfactu = " & DBSet(Text1(1).Text, "F")
'
'    Numf = SugerirCodigoSiguienteStr("rrecibpozos", "numlinea", Sql)
'    Text1(31).Text = Numf
'    '08/09/2010
    
    If cambiaSQL Then vSQL = CadenaInsertarDesdeForm(Me)
    
    
    'Aqui empieza transaccion
    conn.BeginTrans
    MenError = "Error al insertar en la tabla Cabecera de Factura (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    CadenaCambio = vSQL
    
    '[Monica]19/11/2013: a�adido el log de que ha insertado
    '------------------------------------------------------------------------------
    '  LOG de acciones
    ValorAnterior = ""
    
    Set LOG = New cLOG
    LOG.InsertarCambiosRegistros 14, vUsu, "Cambio Recibos Pozos-I", "rrecibpozos", ObtenerWhereCab(False)
    Set LOG = Nothing
    '-----------------------------------------------------------------------------
    
    
    MenError = "Error al actualizar el contador de la Factura."
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


Private Sub CalculoTotales()
Dim Base As Currency
Dim Tiva As Currency
Dim PorIva As Currency
Dim impiva As Currency
Dim BaseReten As Currency
Dim BaseAFO As Currency
Dim PorRet As Currency
Dim ImpRet As Currency
Dim PorAFO As Currency
Dim ImpAFO As Currency
Dim TotFac As Currency

    Base = CCur(ComprobarCero(Text1(5).Text))
    PorIva = CCur(ComprobarCero(Text1(4).Text))
    impiva = Round2(Base * PorIva / 100, 2)
    
    
    TotFac = Base + impiva

    If impiva = 0 Then
        Text1(6).Text = "0"
    Else
        Text1(6).Text = Format(impiva, "###,##0.00")
    End If
    
    If TotFac = 0 Then
        Text1(7).Text = "0"
    Else
        Text1(7).Text = Format(TotFac, "###,##0.00")
    End If
End Sub



Private Sub PonerHidrantesSocio()
Dim cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim Rs As ADODB.Recordset


    If Text1(2).Text = "" Then Exit Sub
    
    If Not (Modo = 3) Then Exit Sub

    cad = "rpozos.codsocio = " & DBSet(Text1(2).Text, "N")
     
    Cad1 = "select count(*) from rpozos where " & cad
     
    NumRegis = TotalRegistros(Cad1)
    
    If NumRegis = 0 Then Exit Sub
    If NumRegis = 1 Then
        Cad1 = "select hidrante, rpozos.codparti, rpartida.nomparti, rpozos.poligono from rpozos inner join rpartida on rpozos.codparti = rpartida.codparti where " & cad
        Set Rs = New ADODB.Recordset
        Rs.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            Text1(14).Text = DBLet(Rs.Fields(0).Value)
        End If
    Else
        Set frmMens = New frmMensajes
        frmMens.cadWHERE = " and " & cad
        frmMens.campo = Text1(14).Text
        frmMens.OpcionMensaje = 23
        frmMens.Show vbModal
        Set frmMens = Nothing
    End If
    
End Sub




Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
Dim B As Boolean
Dim I As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)
    
    CargaGridGnral Me.DataGridAux(Index), Me.Adoaux(Index), tots, PrimeraVez
    
    Select Case Index
        Case 2 'pozos
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;S|txtAux3(1)|T|Fases|900|;" 'codsocio,numfase
            tots = tots & "S|txtAux3(2)|T|Acciones|1200|;"
            tots = tots & "S|txtAux3(3)|T|Observaciones|5280|;"
            arregla tots, DataGridAux(Index), Me, 350
        
'            DataGridAux(Index).Columns(3).Alignment = dbgLeft
        
            B = (Modo = 4) And ((ModificaLineas = 1) Or (ModificaLineas = 2))

        Case 0 'hidrantes
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;S|txtAux4(4)|T|Hidrante|1400|;" 'codsocio,numfase
            tots = tots & "S|txtAux4(5)|T|Hanegadas|1200|;"
            arregla tots, DataGridAux(Index), Me, 350
        
'            DataGridAux(Index).Columns(3).Alignment = dbgLeft
        
            B = (Modo = 4) And ((ModificaLineas = 1) Or (ModificaLineas = 2))

        Case 1 'campos
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;S|txtAux5(4)|T|Campo|1250|;" 'codsocio,numfase
            tots = tots & "S|txtAux5(5)|T|Hdas|900|;S|txtAux5(6)|T|Pr/Ha.Cuota|1400|;S|txtAux5(7)|T|Pr/Ha.Ordin.|1400|;"
            tots = tots & "S|txtAux5(8)|T|Poligono|900|;S|txtAux5(9)|T|Parcela|1100|;S|txtAux5(10)|T|SParcela|850|;S|txtAux5(11)|T|Importe|1100|;"
            arregla tots, DataGridAux(Index), Me, 350
        
            DataGridAux(Index).Columns(11).Alignment = dbgRight
        
            B = (Modo = 4) And ((ModificaLineas = 1) Or (ModificaLineas = 2))


    End Select
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
'    ' **** si n'hi han ll�nies en grids i camps fora d'estos ****
'    If Not AdoAux(Index).Recordset.EOF Then
'        DataGridAux_RowColChange Index, 1, 1
'    Else
'        LimpiarCamposFrame Index
'    End If
'    ' **********************************************************
      
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
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
Dim SQL As String
Dim tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
       Case 0 ' hidrantes
            tabla = "rrecibpozos_hid"
            SQL = "SELECT codtipom,numfactu,fecfactu,numlinea,hidrante, hanegada "
            SQL = SQL & " FROM " & tabla
            If enlaza Then
                SQL = SQL & ObtenerWhereCab(True)
            Else
                SQL = SQL & " WHERE numfactu = -1"
            End If
            SQL = SQL & " ORDER BY " & tabla & ".hidrante "
       
       Case 2 ' pozos
            tabla = "rrecibpozos_acc"
            SQL = "SELECT codtipom,numfactu,fecfactu,numlinea,numfases, acciones,observac "
            SQL = SQL & " FROM " & tabla
            If enlaza Then
                SQL = SQL & ObtenerWhereCab(True)
            Else
                SQL = SQL & " WHERE numfactu = -1"
            End If
            SQL = SQL & " ORDER BY " & tabla & ".numfases "
            
            
       Case 1 ' campos
            tabla = "rrecibpozos_cam"
            SQL = "SELECT codtipom,numfactu,fecfactu,numlinea,codcampo, hanegada, precio1, precio2, poligono, parcela, subparce, if(coalesce(hanegada,0) <> 0,round((coalesce(precio1,0) + coalesce(precio2,0)) * hanegada,2),0) importe "
            SQL = SQL & " FROM " & tabla
            If enlaza Then
                SQL = SQL & ObtenerWhereCab(True)
            Else
                SQL = SQL & " WHERE numfactu = -1"
            End If
            SQL = SQL & " ORDER BY " & tabla & ".codcampo "
            
            
    End Select
    ' ********************************************************************************
    
    MontaSQLCarga = SQL
End Function

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la cap�alera ***
    vWhere = vWhere & " codtipom=" & DBSet(Mid(Combo1(0).Text, 1, 3), "T") & " and numfactu = " & Text1(0).Text & _
                      " and fecfactu = " & DBSet(Text1(1).Text, "F") & " and numlinea = " & Text1(31).Text
    ' *******************************************************
    ObtenerWhereCab = vWhere
End Function

Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vtabla As String
Dim anc As Single
Dim I As Integer
    
    ModificaLineas = 1 'Posem Modo Afegir Ll�nia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Cap�alera
        cmdAceptar_Click
        If ModificaLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la cap�alera ***
    BloquearTxt Text1(0), True
    ' **************************************************

    ' *** posar el nom del les distintes taules de ll�nies ***
    Select Case Index
        Case 0: vtabla = "rrecibpozos_hid"
        Case 1: vtabla = "rrecibpozos_campos"
        Case 2: vtabla = "rrecibpozos_acc"
    End Select
    ' ********************************************************
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0

            AnyadirLinea DataGridAux(Index), Adoaux(Index)
    
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineas Index, ModificaLineas, anc
        
            For I = 0 To txtAux3.Count - 1
                txtAux3(I).Text = ""
            Next I
            
            txtAux4(0).Text = Mid(Combo1(0).Text, 1, 3) ' tipo de movimiento
            txtAux4(1).Text = Text1(0).Text ' numero de factura
            txtAux4(2).Text = Text1(1).Text ' fecha
            txtAux4(3).Text = Text1(31).Text ' numero de linea
            txtAux4(4).Text = "" 'hidrante
            txtAux4(5).Text = "" 'hanegada
            
            PonerFoco txtAux4(4)
         
        
        Case 1

            AnyadirLinea DataGridAux(Index), Adoaux(Index)
    
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineas Index, ModificaLineas, anc
        
            For I = 0 To txtAux5.Count - 1
                txtAux5(I).Text = ""
            Next I
            
            txtAux5(0).Text = Mid(Combo1(0).Text, 1, 3) ' tipo de movimiento
            txtAux5(1).Text = Text1(0).Text ' numero de factura
            txtAux5(2).Text = Text1(1).Text ' fecha
            txtAux5(3).Text = Text1(31).Text ' numero de linea
            txtAux5(4).Text = "" 'campo
            txtAux5(5).Text = "" 'hanegada
            txtAux5(6).Text = "" 'precio1
            txtAux5(7).Text = "" 'precio2
            txtAux5(8).Text = "" 'poligono
            txtAux5(9).Text = "" 'parcela
            txtAux5(10).Text = "" 'subparcela
            
            
            PonerFoco txtAux5(4)
        
        Case 2

            AnyadirLinea DataGridAux(Index), Adoaux(Index)
    
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineas Index, ModificaLineas, anc
        
            For I = 0 To txtAux3.Count - 1
                txtAux3(I).Text = ""
            Next I
            
            txtAux3(0).Text = Mid(Combo1(0).Text, 1, 3) ' tipo de movimiento
            txtAux3(4).Text = Text1(0).Text ' numero de factura
            txtAux3(5).Text = Text1(1).Text ' fecha
            txtAux3(6).Text = Text1(31).Text ' numero de linea
            txtAux3(1).Text = NumF 'numero de fase
            PonerFoco txtAux3(1)
         
            
    End Select
End Sub



Private Sub BotonEliminarLinea(Index As Integer)
Dim SQL As String
Dim vWhere As String
Dim Eliminar As Boolean

    On Error GoTo Error2

    ModificaLineas = 3 'Posem Modo Eliminar Ll�nia
    
    If Modo = 4 Then 'Modificar Cap�alera
        cmdAceptar_Click
        If ModificaLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index

    If Adoaux(Index).Recordset.EOF Then Exit Sub
'    If Not SepuedeBorrar(Index) Then Exit Sub
    NumTabMto = Index
    Eliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 0 'hidrantes
            SQL = "�Seguro que desea eliminar el registro?"
            SQL = SQL & vbCrLf & "Hidrante: " & Adoaux(Index).Recordset!Hidrante
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                SQL = "DELETE FROM rrecibpozos_hid"
                SQL = SQL & vWhere & " AND hidrante= " & DBLet(Adoaux(Index).Recordset!Hidrante, "T")
                
                
                CadenaCambio = SQL
                '------------------------------------------------------------------------------
                '  LOG de acciones
                ValorAnterior = ""
                Set LOG = New cLOG
                LOG.InsertarCambiosRegistros 14, vUsu, "Cambio Recibos Pozos-D", "rrecibpozos_hid", ObtenerWhereCab(False)
                Set LOG = Nothing
                '-----------------------------------------------------------------------------
            End If
        
        Case 1 'campos
            SQL = "�Seguro que desea eliminar el registro?"
            SQL = SQL & vbCrLf & "Campos: " & Adoaux(Index).Recordset!codcampo
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                SQL = "DELETE FROM rrecibpozos_cam"
                SQL = SQL & vWhere & " AND codcampo= " & DBLet(Adoaux(Index).Recordset!codcampo, "N")
                
                CadenaCambio = SQL
                '------------------------------------------------------------------------------
                '  LOG de acciones
                ValorAnterior = ""
                Set LOG = New cLOG
                LOG.InsertarCambiosRegistros 14, vUsu, "Cambio Recibos Pozos-D", "rrecibpozos_cam", ObtenerWhereCab(False)
                Set LOG = Nothing
                '-----------------------------------------------------------------------------
            End If
        
        
        Case 2 'pozos
            SQL = "�Seguro que desea eliminar el registro?"
            SQL = SQL & vbCrLf & "Numero Fase: " & Adoaux(Index).Recordset!numfases
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                SQL = "DELETE FROM rrecibpozos_acc"
                SQL = SQL & vWhere & " AND numfases= " & DBLet(Adoaux(Index).Recordset!numfases, "N")
            
                CadenaCambio = SQL
                '------------------------------------------------------------------------------
                '  LOG de acciones
                ValorAnterior = ""
                Set LOG = New cLOG
                LOG.InsertarCambiosRegistros 14, vUsu, "Cambio Recibos Pozos-D", "rrecibpozos_acc", ObtenerWhereCab(False)
                Set LOG = Nothing
                '-----------------------------------------------------------------------------
            End If
        
    End Select

    If Eliminar Then
        NumRegElim = Adoaux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        conn.Execute SQL
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        If Index <> 3 Then _
            CargaGrid Index, True
        ' ***************************************************
        If Not SituarDataTrasEliminar(Adoaux(Index), NumRegElim, True) Then
'            PonerCampos
            
        End If
        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
    End If
    
    ModificaLineas = 0
    PosicionarData
    
    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub


Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim I As Integer
    Dim J As Integer
    
    If Adoaux(Index).Recordset.EOF Then Exit Sub
    If Adoaux(Index).Recordset.RecordCount < 1 Then Exit Sub
    
    ModificaLineas = 2 'Modificar ll�nia
       
    If Modo = 4 Then 'Modificar Cap�alera
        cmdAceptar_Click
        If ModificaLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    ' *** bloqueje la clau primaria de la cap�alera ***
    BloquearTxt Text1(0), True
    ' *********************************
    
    Select Case Index
        Case 0, 1, 2 ' *** pose els index de ll�nies que tenen datagrid (en o sense tab) ***
            If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
                I = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
                DataGridAux(Index).Scroll 0, I
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
        Case 0 'hidrantes
            For I = 0 To 5
                txtAux4(I).Text = DataGridAux(Index).Columns(I).Text
            Next I
            
            CargarValoresAnteriores Me, 2, "FrameAux0"
            
        Case 1 'campos
            For I = 0 To 10
                txtAux5(I).Text = DataGridAux(Index).Columns(I).Text
            Next I
        
            CargarValoresAnteriores Me, 2, "FrameAux1"
            
        Case 2 'pozos
            For I = 1 To 3
                txtAux3(I + 3).Text = DataGridAux(Index).Columns(I).Text
            Next I
            txtAux3(0).Text = DataGridAux(Index).Columns(0).Text
            txtAux3(1).Text = DataGridAux(Index).Columns(4).Text
            txtAux3(2).Text = DataGridAux(Index).Columns(5).Text
            txtAux3(3).Text = DataGridAux(Index).Columns(6).Text
            
            CargarValoresAnteriores Me, 2, "FrameAux2"
        
    End Select
    
    LLamaLineas Index, ModificaLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 ' hidrantes
            PonerFoco txtAux4(5)
        Case 1 ' campos
            PonerFoco txtAux5(4)
        Case 2 ' pozos
            PonerFoco txtAux3(2)
    End Select
    ' ***************************************************************************************
End Sub


Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim B As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    If Index <> 3 Then DeseleccionaGrid DataGridAux(Index)
    ' ***************************************************
       
    B = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Ll�nies
    Select Case Index
        Case 0 ' hidrantes
            For jj = 4 To 5
                txtAux4(jj).visible = B
                txtAux4(jj).Top = alto
            Next jj
            If xModo = 2 Then txtAux4(4).visible = False
        Case 1 ' campos
            For jj = 4 To 10
                txtAux5(jj).visible = B
                txtAux5(jj).Top = alto
            Next jj
        
        Case 2 ' pozos
            For jj = 1 To 3
                txtAux3(jj).visible = B
                txtAux3(jj).Top = alto
            Next jj
    End Select
End Sub


Private Sub InsertarLinea()
'Inserta registre en les taules de Ll�nies
Dim nomframe As String
Dim B As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'hidrante
        Case 1: nomframe = "FrameAux1" 'campos
        Case 2: nomframe = "FrameAux2" 'pozos
    End Select
    ' ***************************************************************
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomframe) Then
            ' *** si n'hi ha que fer alguna cosa abas d'insertar
            ' *************************************************
            B = BLOQUEADesdeFormulario2(Me, Data1, 1)
            
            Select Case NumTabMto
                Case 0, 1, 2 ' *** els index de les llinies en grid (en o sense tab) ***
                    CargaGrid NumTabMto, True
                    If B Then BotonAnyadirLinea NumTabMto
'                Case 3 ' *** els index dels tabs que NO tenen grid ***
'                    CargaFrame 3, True
'                    If b Then BotonModificar
'                    ModoLineas = 0
'                LLamaLineas NumTabMto, 0
            End Select
           
        End If
    End If
End Sub


Private Sub ModificarLinea()
'Modifica registre en les taules de Ll�nies
Dim nomframe As String
Dim V As Integer
Dim cad As String
Dim TablaAux As String

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'hidrantes
        Case 1: nomframe = "FrameAux1" 'campos
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
                Case 0: TablaAux = "rrecibpozos_hid" 'hidrantes
                Case 1: TablaAux = "rrecibpozos_cam" 'campos
                Case 2: TablaAux = "rrecibpozos_acc" 'pozos
            End Select
    
            '------------------------------------------------------------------------------
            '  LOG de acciones
            Set LOG = New cLOG
            LOG.InsertarCambiosRegistros 14, vUsu, "Cambio Recibos Pozos-U", TablaAux, ObtenerWhereCab(False)
            Set LOG = Nothing
            '-----------------------------------------------------------------------------
            
            ' ******************************************************
            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
            ModificaLineas = 0

            If NumTabMto <> 3 Then
                V = Adoaux(NumTabMto).Recordset.Fields(1) 'el 2 es el n� de llinia
                CargaGrid NumTabMto, True
            End If

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


Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim SQL As Integer
Dim B As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte

    DatosOkLlin = True
    
    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False

    B = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not B Then Exit Function
    
    
    If B And NumTabMto = 2 And ModificaLineas = 1 Then
        SQL = DevuelveValor("select acciones from rrecibpozos_acc where codtipom = " & DBSet(txtAux3(0).Text, "T") & " and numfactu = " & DBSet(txtAux3(4).Text, "N") & " and fecfactu = " & DBSet(txtAux3(5).Text, "F") & " and numlinea = " & DBSet(txtAux3(6).Text, "N") & " and numfase = " & DBSet(txtAux3(1).Text, "N"))
        If SQL <> 0 Then
            MsgBox "El n�mero de fase ya existe. Reintroduzca.", vbExclamation
            B = False
            PonerFoco txtAux3(1)
        End If
    End If
    
    If B And NumTabMto = 0 And ModificaLineas = 1 Then
        SQL = DevuelveValor("select count(*) from rpozos where hidrante = " & DBSet(txtAux4(4).Text, "T"))
        If SQL = 0 Then
            MsgBox "El hidrante no existe. Reintroduzca.", vbExclamation
            B = False
            PonerFoco txtAux4(4)
        End If
        If B Then
            SQL = DevuelveValor("select count(*) from rrecibpozos_hid where codtipom = " & DBSet(txtAux4(0).Text, "T") & " and numfactu = " & DBSet(txtAux4(1).Text, "N") & " and fecfactu = " & DBSet(txtAux4(2).Text, "F") & " and numlinea = " & DBSet(txtAux4(3).Text, "N") & " and hidrante = " & DBSet(txtAux4(4).Text, "T"))
            If SQL <> 0 Then
                MsgBox "El hidrante ya existe en el recibo. Revise.", vbExclamation
                B = False
                PonerFoco txtAux4(4)
            End If
        End If
    End If
    
    If B And NumTabMto = 1 And ModificaLineas = 1 Then
        SQL = DevuelveValor("select count(*) from rcampos where codcampo = " & DBSet(txtAux5(4).Text, "N"))
        If SQL = 0 Then
            MsgBox "El campo no existe. Reintroduzca.", vbExclamation
            B = False
            PonerFoco txtAux5(4)
        End If
        If B Then
            SQL = DevuelveValor("select count(*) from rrecibpozos_cam where codtipom = " & DBSet(txtAux5(0).Text, "T") & " and numfactu = " & DBSet(txtAux5(1).Text, "N") & " and fecfactu = " & DBSet(txtAux5(2).Text, "F") & " and numlinea = " & DBSet(txtAux5(3).Text, "N") & " and codcampo = " & DBSet(txtAux5(4).Text, "T"))
            If SQL <> 0 Then
                MsgBox "El campo ya existe en el recibo. Revise.", vbExclamation
                B = False
                PonerFoco txtAux5(4)
            End If
        End If
    End If
    
    
    
    
    DatosOkLlin = B

EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
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


    ' ******* configurar el LostFocus dels camps de ll�nies (dins i fora grid) ********
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

    ' ******* configurar el LostFocus dels camps de ll�nies (dins i fora grid) ********
    Select Case Index
        Case 5
            PonerFormatoDecimal txtAux4(Index), 11
        
            cmdAceptar.SetFocus
    End Select
    ' ******************************************************************************
End Sub


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
Dim cadena As String
    
    If Not PerderFocoGnral(txtAux5(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    ' ******* configurar el LostFocus dels camps de ll�nies (dins i fora grid) ********
    Select Case Index
        Case 5, 6, 7
            PonerFormatoDecimal txtAux5(Index), 11
        
        Case 10
            cmdAceptar.SetFocus
    End Select
    ' ******************************************************************************
End Sub


Private Sub Timer1_Timer()
    Label1(13).visible = Not Label1(13).visible
End Sub

