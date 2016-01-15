VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPOZRecibos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico de Recibos "
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   9630
   Icon            =   "frmPOZRecibos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameRectifica 
      Caption         =   "Factura Rectificativa "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   180
      TabIndex        =   144
      Top             =   3000
      Width           =   9165
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   45
         Left            =   180
         MaxLength       =   3
         TabIndex        =   147
         Tag             =   "Tipo Movimiento Fra Rectifica|T|S|||rrecibpozos|codtipomrec|||"
         Top             =   540
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   47
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   146
         Tag             =   "Fecha Factura Rectificativa|F|S|||rrecibpozos|fecfacturec|dd/mm/yyyy||"
         Text            =   "123"
         Top             =   540
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   46
         Left            =   1800
         MaxLength       =   7
         TabIndex        =   145
         Tag             =   "Nº Factura Rectifica|N|S|||rrecibpozos|numfacturec|0000000||"
         Text            =   "Text1"
         Top             =   540
         Width           =   930
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   6
         Left            =   3720
         Picture         =   "frmPOZRecibos.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Recibo"
         Height          =   255
         Index           =   17
         Left            =   180
         TabIndex        =   150
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Recibo"
         Height          =   255
         Index           =   14
         Left            =   1800
         TabIndex        =   149
         Top             =   300
         Width           =   885
      End
      Begin VB.Label Label22 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   2940
         TabIndex        =   148
         Top             =   300
         Width           =   495
      End
   End
   Begin VB.PictureBox cmdRectificativa 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   5730
      ScaleHeight     =   435
      ScaleWidth      =   525
      TabIndex        =   143
      ToolTipText     =   "Rectificativa"
      Top             =   510
      Width           =   525
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Es Contado"
      Height          =   195
      Index           =   3
      Left            =   4170
      TabIndex        =   137
      Tag             =   "Es Contado|N|N|0|1|rrecibpozos|escontado|0||"
      Top             =   2760
      Width           =   1185
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   40
      Left            =   3060
      MaxLength       =   10
      TabIndex        =   16
      Tag             =   "Fec.Albaran|F|S|||rrecibpozos|fecalbar|dd/mm/yyyy||"
      Text            =   "1234567890123456789012345"
      Top             =   2640
      Width           =   1050
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   39
      Left            =   1530
      MaxLength       =   10
      TabIndex        =   15
      Tag             =   "Ticket|N|S|||rrecibpozos|numalbar|0000000||"
      Text            =   "1234567890123456789012345"
      Top             =   2640
      Width           =   990
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   38
      Left            =   7740
      MaxLength       =   10
      TabIndex        =   11
      Tag             =   "Toma|N|S|||rrecibpozos|nroorden|000000||"
      Text            =   "1234567890123456789012345"
      Top             =   1650
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   37
      Left            =   5250
      MaxLength       =   25
      TabIndex        =   10
      Tag             =   "Parcelas|T|S|||rrecibpozos|parcelas|||"
      Text            =   "1234567890123456789012345"
      Top             =   1650
      Width           =   2460
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   36
      Left            =   4200
      MaxLength       =   10
      TabIndex        =   9
      Tag             =   "Poligono|T|S|||rrecibpozos|poligono|||"
      Text            =   "1234567890"
      Top             =   1650
      Width           =   1005
   End
   Begin VB.PictureBox cmdCampos 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   6420
      ScaleHeight     =   435
      ScaleWidth      =   525
      TabIndex        =   116
      ToolTipText     =   "Campos"
      Top             =   510
      Width           =   525
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Pasa Aridoc"
      Height          =   195
      Index           =   2
      Left            =   4170
      TabIndex        =   20
      Tag             =   "Pasa Aridoc|N|N|0|1|rrecibpozos|pasaridoc|0||"
      Top             =   2520
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   35
      Left            =   3060
      MaxLength       =   10
      TabIndex        =   14
      Tag             =   "Precio|N|S|||rrecibpozos|precio|###,##0.0000||"
      Text            =   "1234567890"
      Top             =   2250
      Width           =   1050
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   34
      Left            =   1530
      MaxLength       =   10
      TabIndex        =   13
      Tag             =   "Importe Dto|N|S|||rrecibpozos|impdto|##,###,##0.00||"
      Top             =   2250
      Width           =   1350
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   33
      Left            =   210
      MaxLength       =   10
      TabIndex        =   12
      Tag             =   "Porc.Dto|N|S|||rrecibpozos|porcdto|##0.00||"
      Top             =   2250
      Width           =   1080
   End
   Begin VB.PictureBox cmdHidrantes 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   7140
      ScaleHeight     =   435
      ScaleWidth      =   525
      TabIndex        =   102
      ToolTipText     =   "Hidrantes"
      Top             =   510
      Width           =   525
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   32
      Left            =   3360
      MaxLength       =   10
      TabIndex        =   8
      Tag             =   "Diferencia Dias|N|S|||rrecibpozos|difdias|###,##0||"
      Text            =   "1234567"
      Top             =   1650
      Width           =   735
   End
   Begin VB.PictureBox cmdParticipa 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   7860
      ScaleHeight     =   435
      ScaleWidth      =   525
      TabIndex        =   90
      ToolTipText     =   "Participaciones"
      Top             =   510
      Width           =   525
   End
   Begin VB.PictureBox cmdConceptos 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   8580
      ScaleHeight     =   435
      ScaleWidth      =   525
      TabIndex        =   84
      ToolTipText     =   "Conceptos"
      Top             =   510
      Width           =   525
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   210
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   990
      Width           =   1650
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   20
      Left            =   630
      MaxLength       =   10
      TabIndex        =   68
      Tag             =   "Tipo de Fichero|T|S|||rrecibpozos|codtipom||S|"
      Top             =   990
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   4200
      MaxLength       =   6
      TabIndex        =   4
      Tag             =   "Cod.Socio|N|N|0|999999|rrecibpozos|codsocio|000000|N|"
      Text            =   "Text1"
      Top             =   990
      Width           =   780
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   14
      Left            =   210
      MaxLength       =   10
      TabIndex        =   5
      Tag             =   "Hidrante|T|S|||rrecibpozos|hidrante||N|"
      Text            =   "1234567"
      Top             =   1650
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   12
      Left            =   1530
      MaxLength       =   7
      TabIndex        =   6
      Tag             =   "Consumo|N|S|||rrecibpozos|consumo|||"
      Text            =   "1234567"
      Top             =   1650
      Width           =   885
   End
   Begin VB.TextBox Text1 
      Height          =   645
      Index           =   15
      Left            =   5460
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Tag             =   "Conceptol|T|S|||rrecibpozos|concepto|||"
      Text            =   "frmPOZRecibos.frx":0097
      Top             =   2250
      Width           =   3885
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   13
      Left            =   2430
      MaxLength       =   10
      TabIndex        =   7
      Tag             =   "Cuota|N|S|||rrecibpozos|impcuota|###,##0.00||"
      Text            =   "1234567"
      Top             =   1650
      Width           =   885
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Height          =   315
      Index           =   0
      Left            =   2010
      MaxLength       =   7
      TabIndex        =   2
      Tag             =   "Nº Factura|N|S|||rrecibpozos|numfactu|0000000|S|"
      Text            =   "Text1"
      Top             =   990
      Width           =   885
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "Fecha Factura|F|N|||rrecibpozos|fecfactu|dd/mm/yyyy|S|"
      Top             =   990
      Width           =   1065
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   5010
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   67
      Text            =   "Text2"
      Top             =   990
      Width           =   4110
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Contabilizado"
      Height          =   195
      Index           =   1
      Left            =   4170
      TabIndex        =   19
      Tag             =   "Contabilizado|N|N|0|1|rrecibpozos|contabilizado|0||"
      Top             =   2280
      Width           =   1245
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Impreso"
      Height          =   195
      Index           =   0
      Left            =   4170
      TabIndex        =   18
      Tag             =   "Impreso|N|N|0|1|rrecibpozos|impreso|0||"
      Top             =   2010
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   210
      TabIndex        =   41
      Top             =   6660
      Width           =   2175
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
         Left            =   240
         TabIndex        =   42
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8280
      TabIndex        =   39
      Top             =   6765
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7080
      TabIndex        =   38
      Top             =   6780
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Añadir"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
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
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   7410
         TabIndex        =   0
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   8280
      TabIndex        =   40
      Top             =   6780
      Visible         =   0   'False
      Width           =   1035
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
   Begin VB.Frame Frame5 
      Caption         =   "Total Factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1095
      Index           =   0
      Left            =   180
      TabIndex        =   44
      Top             =   3000
      Width           =   9135
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   5370
         MaxLength       =   10
         TabIndex        =   24
         Tag             =   "Importe Iva|N|S|||rrecibpozos|imporiva|###,##0.00||"
         Top             =   540
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   1620
         MaxLength       =   2
         TabIndex        =   22
         Tag             =   "Tipo Iva|N|S|||rrecibpozos|tipoiva|00||"
         Text            =   "Text1"
         Top             =   540
         Width           =   600
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   4560
         MaxLength       =   6
         TabIndex        =   23
         Tag             =   "Porc.Iva|N|S|||rrecibpozos|porc_iva|##0.00||"
         Text            =   "123"
         Top             =   540
         Width           =   645
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   2310
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   85
         Text            =   "Text2"
         Top             =   540
         Width           =   2190
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   180
         MaxLength       =   10
         TabIndex        =   21
         Tag             =   "Base Imponible|N|N|||rrecibpozos|baseimpo|###,##0.00||"
         Top             =   540
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CAE3FD&
         Height          =   315
         Index           =   7
         Left            =   7140
         MaxLength       =   10
         TabIndex        =   25
         Tag             =   "Total Factura|N|N|||rrecibpozos|totalfact|###,##0.00||"
         Top             =   540
         Width           =   1830
      End
      Begin VB.Label Label2 
         Caption         =   "% Iva"
         Height          =   255
         Left            =   4530
         TabIndex        =   49
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Iva"
         Height          =   255
         Index           =   8
         Left            =   1620
         TabIndex        =   48
         Top             =   300
         Width           =   285
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1950
         ToolTipText     =   "Buscar Iva"
         Top             =   300
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Importe IVA"
         Height          =   255
         Index           =   7
         Left            =   5400
         TabIndex        =   47
         Top             =   300
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "TOTAL FACTURA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   7140
         TabIndex        =   46
         Top             =   300
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
         Height          =   255
         Index           =   10
         Left            =   180
         TabIndex        =   45
         Top             =   300
         Width           =   1455
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Height          =   285
      Index           =   31
      Left            =   420
      MaxLength       =   7
      TabIndex        =   88
      Tag             =   "Linea|N|N|||rrecibpozos|numlinea|0000000|S|"
      Text            =   "Text1"
      Top             =   1650
      Width           =   885
   End
   Begin VB.Frame FrameCampos 
      Caption         =   "Campos"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2415
      Left            =   180
      TabIndex        =   117
      Top             =   4170
      Visible         =   0   'False
      Width           =   9135
      Begin VB.Frame FrameAux1 
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   120
         TabIndex        =   118
         Top             =   300
         Width           =   8865
         Begin VB.TextBox txtAux5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   10
            Left            =   6060
            MaxLength       =   9
            TabIndex        =   131
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
            Height          =   290
            Index           =   9
            Left            =   5460
            MaxLength       =   6
            TabIndex        =   130
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
            Height          =   290
            Index           =   8
            Left            =   4860
            MaxLength       =   3
            TabIndex        =   129
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
            Height          =   290
            Index           =   7
            Left            =   4260
            MaxLength       =   9
            TabIndex        =   128
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
            Height          =   290
            Index           =   6
            Left            =   3660
            MaxLength       =   9
            TabIndex        =   127
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
            Height          =   290
            Index           =   3
            Left            =   2160
            MaxLength       =   7
            TabIndex        =   122
            Tag             =   "Linea|N|N|||rrecibpozos_cam|numlinea|0000000|S|"
            Text            =   "Linea"
            Top             =   1530
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtAux5 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   2
            Left            =   1350
            MaxLength       =   10
            TabIndex        =   121
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
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   1
            Left            =   750
            MaxLength       =   7
            TabIndex        =   120
            Tag             =   "Nº Factura|N|N|||rrecibpozos_cam|numfactu|0000000|S|"
            Text            =   "recibo"
            Top             =   1530
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox txtAux5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   0
            Left            =   180
            MaxLength       =   6
            TabIndex        =   119
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
            Height          =   290
            Index           =   5
            Left            =   3060
            MaxLength       =   9
            TabIndex        =   126
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
            Height          =   290
            Index           =   4
            Left            =   2580
            MaxLength       =   10
            TabIndex        =   124
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
            TabIndex        =   123
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
            Bindings        =   "frmPOZRecibos.frx":009F
            Height          =   1395
            Index           =   1
            Left            =   105
            TabIndex        =   125
            Top             =   420
            Width           =   8520
            _ExtentX        =   15028
            _ExtentY        =   2461
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
   End
   Begin VB.Frame FrameHidrantes 
      Caption         =   "Hidrantes"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2415
      Left            =   180
      TabIndex        =   103
      Top             =   4170
      Visible         =   0   'False
      Width           =   9135
      Begin VB.Frame FrameAux0 
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   120
         TabIndex        =   104
         Top             =   300
         Width           =   8865
         Begin VB.TextBox txtAux4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   4
            Left            =   2580
            MaxLength       =   10
            TabIndex        =   109
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
            Height          =   290
            Index           =   5
            Left            =   3060
            MaxLength       =   9
            TabIndex        =   110
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
            Height          =   290
            Index           =   0
            Left            =   300
            MaxLength       =   6
            TabIndex        =   108
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
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   1
            Left            =   750
            MaxLength       =   7
            TabIndex        =   107
            Tag             =   "Nº Factura|N|N|||rrecibpozos_hid|numfactu|0000000|S|"
            Text            =   "recibo"
            Top             =   1530
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox txtAux4 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   2
            Left            =   1350
            MaxLength       =   10
            TabIndex        =   106
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
            Height          =   290
            Index           =   3
            Left            =   2160
            MaxLength       =   7
            TabIndex        =   105
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
            TabIndex        =   111
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
            Bindings        =   "frmPOZRecibos.frx":00B7
            Height          =   1395
            Index           =   0
            Left            =   105
            TabIndex        =   112
            Top             =   420
            Width           =   3120
            _ExtentX        =   5503
            _ExtentY        =   2461
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
   End
   Begin VB.Frame FrameConceptos 
      Caption         =   "Conceptos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2415
      Left            =   180
      TabIndex        =   69
      Top             =   4170
      Width           =   9105
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   30
         Left            =   7680
         MaxLength       =   100
         TabIndex        =   79
         Tag             =   "Importe Art.4|N|S|||rrecibpozos|importear4|###,##0.00||"
         Top             =   2010
         Width           =   1260
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   29
         Left            =   240
         MaxLength       =   100
         TabIndex        =   78
         Tag             =   "Concepto Articulo 4|T|S|||rrecibpozos|conceptoar4|||"
         Text            =   "1234567"
         Top             =   2010
         Width           =   7245
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   28
         Left            =   7680
         MaxLength       =   100
         TabIndex        =   77
         Tag             =   "Importe Art.3|N|S|||rrecibpozos|importear3|###,##0.00||"
         Top             =   1710
         Width           =   1260
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   27
         Left            =   240
         MaxLength       =   100
         TabIndex        =   76
         Tag             =   "Concepto Articulo 3|T|S|||rrecibpozos|conceptoar3|||"
         Text            =   "1234567"
         Top             =   1710
         Width           =   7245
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   26
         Left            =   7680
         MaxLength       =   100
         TabIndex        =   75
         Tag             =   "Importe Art.2|N|S|||rrecibpozos|importear2|###,##0.00||"
         Top             =   1410
         Width           =   1260
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   25
         Left            =   240
         MaxLength       =   100
         TabIndex        =   74
         Tag             =   "Concepto Articulo 2|T|S|||rrecibpozos|conceptoar2|||"
         Text            =   "1234567"
         Top             =   1410
         Width           =   7245
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   7680
         MaxLength       =   100
         TabIndex        =   73
         Tag             =   "Importe Art.1|N|S|||rrecibpozos|importear1|###,##0.00||"
         Top             =   1110
         Width           =   1260
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   23
         Left            =   240
         MaxLength       =   100
         TabIndex        =   72
         Tag             =   "Concepto Articulo 1|T|S|||rrecibpozos|conceptoar1|||"
         Text            =   "1234567"
         Top             =   1110
         Width           =   7245
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   22
         Left            =   7680
         MaxLength       =   10
         TabIndex        =   71
         Tag             =   "Importe MO|N|S|||rrecibpozos|importemo|###,##0.00||"
         Top             =   540
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   21
         Left            =   240
         MaxLength       =   100
         TabIndex        =   70
         Tag             =   "Concepto MO|T|S|||rrecibpozos|conceptomo|||"
         Text            =   "1234567"
         Top             =   540
         Width           =   7245
      End
      Begin VB.Label Label11 
         Caption         =   "Importe"
         Height          =   255
         Left            =   7710
         TabIndex        =   83
         Top             =   810
         Width           =   1035
      End
      Begin VB.Label Label10 
         Caption         =   "Importe"
         Height          =   255
         Left            =   7680
         TabIndex        =   82
         Top             =   270
         Width           =   1035
      End
      Begin VB.Label Label8 
         Caption         =   "Artículos"
         Height          =   255
         Left            =   240
         TabIndex        =   81
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label7 
         Caption         =   "Mano de Obra"
         Height          =   255
         Left            =   240
         TabIndex        =   80
         Top             =   270
         Width           =   1035
      End
   End
   Begin VB.Frame FrameParticipaciones 
      Caption         =   "Participaciones"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2415
      Left            =   180
      TabIndex        =   89
      Top             =   4170
      Visible         =   0   'False
      Width           =   9135
      Begin VB.Frame FrameAux2 
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   120
         TabIndex        =   91
         Top             =   300
         Width           =   8865
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   6
            Left            =   2160
            MaxLength       =   7
            TabIndex        =   100
            Tag             =   "Linea|N|N|||rrecibpozos_acc|numlinea|0000000|S|"
            Text            =   "Linea"
            Top             =   1530
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   315
            Index           =   5
            Left            =   1350
            MaxLength       =   10
            TabIndex        =   99
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
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   4
            Left            =   750
            MaxLength       =   7
            TabIndex        =   98
            Tag             =   "Nº Factura|N|S|||rrecibpozos_acc|numfactu|0000000|S|"
            Text            =   "recibo"
            Top             =   1530
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox txtAux3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   0
            Left            =   300
            MaxLength       =   6
            TabIndex        =   97
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
            Height          =   290
            Index           =   2
            Left            =   3060
            MaxLength       =   9
            TabIndex        =   93
            Tag             =   "Acciones|N|N|||rrecibpozos_acc|acciones|##0.00||"
            Text            =   "Acciones"
            Top             =   1530
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtAux3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   290
            Index           =   3
            Left            =   3690
            MaxLength       =   30
            TabIndex        =   94
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
            Height          =   290
            Index           =   1
            Left            =   2580
            MaxLength       =   9
            TabIndex        =   92
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
            TabIndex        =   95
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
            Bindings        =   "frmPOZRecibos.frx":00CF
            Height          =   1335
            Index           =   2
            Left            =   105
            TabIndex        =   96
            Top             =   420
            Width           =   7950
            _ExtentX        =   14023
            _ExtentY        =   2355
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
   End
   Begin VB.Frame Frame3 
      Caption         =   "Lectura Anterior"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2385
      Left            =   180
      TabIndex        =   50
      Top             =   4200
      Width           =   2985
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   42
         Left            =   1590
         MaxLength       =   10
         TabIndex        =   29
         Tag             =   "Fecha lectura anterior2|F|S|||rrecibpozos|fech_ant2|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   1980
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   41
         Left            =   1590
         MaxLength       =   7
         TabIndex        =   28
         Tag             =   "Lectura Anterior2|N|S|||rrecibpozos|lect_ant2|0000000||"
         Text            =   "1234567"
         Top             =   1500
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   1590
         MaxLength       =   7
         TabIndex        =   26
         Tag             =   "Lectura Anterior|N|S|||rrecibpozos|lect_ant|0000000||"
         Text            =   "1234567"
         Top             =   420
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   1590
         MaxLength       =   10
         TabIndex        =   27
         Tag             =   "Fecha lectura anterior|F|S|||rrecibpozos|fech_ant|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   960
         Width           =   1065
      End
      Begin VB.Label Label19 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   330
         TabIndex        =   139
         Top             =   2010
         Width           =   570
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   4
         Left            =   1200
         Picture         =   "frmPOZRecibos.frx":00E7
         ToolTipText     =   "Buscar fecha"
         Top             =   2010
         Width           =   240
      End
      Begin VB.Label Label17 
         Caption         =   "Contador"
         Height          =   285
         Left            =   330
         TabIndex        =   138
         Top             =   1530
         Width           =   1125
      End
      Begin VB.Label Label23 
         Caption         =   "Contador"
         Height          =   285
         Left            =   330
         TabIndex        =   52
         Top             =   450
         Width           =   1125
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1200
         Picture         =   "frmPOZRecibos.frx":0172
         ToolTipText     =   "Buscar fecha"
         Top             =   990
         Width           =   240
      End
      Begin VB.Label Label18 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   330
         TabIndex        =   51
         Top             =   990
         Width           =   570
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Lectura Actual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2385
      Left            =   3210
      TabIndex        =   53
      Top             =   4200
      Width           =   2925
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   44
         Left            =   1500
         MaxLength       =   7
         TabIndex        =   32
         Tag             =   "Contador Actual2|N|S|||rrecibpozos|lect_act2|0000000||"
         Text            =   "1234567"
         Top             =   1470
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   43
         Left            =   1500
         MaxLength       =   10
         TabIndex        =   33
         Tag             =   "Fecha Lectura Actual2|F|S|||rrecibpozos|fech_act2|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   1980
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   1500
         MaxLength       =   10
         TabIndex        =   31
         Tag             =   "Fecha Lectura Actual|F|S|||rrecibpozos|fech_act|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   960
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   1500
         MaxLength       =   7
         TabIndex        =   30
         Tag             =   "Contador Actual|N|S|||rrecibpozos|lect_act|0000000||"
         Text            =   "1234567"
         Top             =   420
         Width           =   1065
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   5
         Left            =   1140
         Picture         =   "frmPOZRecibos.frx":01FD
         ToolTipText     =   "Buscar fecha"
         Top             =   1980
         Width           =   240
      End
      Begin VB.Label Label21 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   360
         TabIndex        =   141
         Top             =   1980
         Width           =   705
      End
      Begin VB.Label Label20 
         Caption         =   "Contador"
         Height          =   255
         Left            =   360
         TabIndex        =   140
         Top             =   1500
         Width           =   1035
      End
      Begin VB.Label Label9 
         Caption         =   "Contador"
         Height          =   255
         Left            =   360
         TabIndex        =   55
         Top             =   450
         Width           =   1035
      End
      Begin VB.Label Label12 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   360
         TabIndex        =   54
         Top             =   990
         Width           =   705
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1140
         Picture         =   "frmPOZRecibos.frx":0288
         ToolTipText     =   "Buscar fecha"
         Top             =   990
         Width           =   240
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Precios Aplicados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2385
      Left            =   6180
      TabIndex        =   56
      Top             =   4200
      Width           =   3135
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   1500
         MaxLength       =   8
         TabIndex        =   34
         Tag             =   "Consumo 1|N|S|||rrecibpozos|consumo1|0000000||"
         Text            =   "m3"
         Top             =   450
         Width           =   1245
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   1500
         MaxLength       =   9
         TabIndex        =   37
         Tag             =   "Precio 2|N|S|||rrecibpozos|precio2|#,##0.000||"
         Text            =   "precio2"
         Top             =   1890
         Width           =   1245
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   1500
         MaxLength       =   9
         TabIndex        =   35
         Tag             =   "Precio 1|N|S|||rrecibpozos|precio1|#,##0.000||"
         Text            =   "precio1"
         Top             =   870
         Width           =   1245
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   1500
         MaxLength       =   8
         TabIndex        =   36
         Tag             =   "Consumo 2|N|S|||rrecibpozos|consumo2|0000000||"
         Text            =   "m3"
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label Label14 
         Caption         =   "Precio"
         Height          =   285
         Left            =   300
         TabIndex        =   87
         Top             =   1890
         Width           =   1245
      End
      Begin VB.Label Label13 
         Caption         =   "Hasta m3."
         Height          =   285
         Left            =   300
         TabIndex        =   86
         Top             =   1500
         Width           =   1245
      End
      Begin VB.Label Label28 
         Caption         =   "Hasta m3."
         Height          =   285
         Left            =   300
         TabIndex        =   58
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label Label27 
         Caption         =   "Precio"
         Height          =   285
         Left            =   300
         TabIndex        =   57
         Top             =   870
         Width           =   1245
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "P A G A D O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
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
      TabIndex        =   142
      Top             =   6750
      Width           =   4515
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "/"
      Height          =   255
      Left            =   2520
      TabIndex        =   136
      Top             =   2640
      Width           =   225
   End
   Begin VB.Image imgFec 
      Height          =   240
      Index           =   3
      Left            =   2790
      Picture         =   "frmPOZRecibos.frx":0313
      ToolTipText     =   "Buscar fecha"
      Top             =   2640
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Albarán / Fecha"
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   135
      Top             =   2670
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "Toma"
      Height          =   255
      Index           =   11
      Left            =   7740
      TabIndex        =   134
      Top             =   1350
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Parcelas"
      Height          =   255
      Index           =   6
      Left            =   5250
      TabIndex        =   133
      Top             =   1380
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Poligono"
      Height          =   255
      Index           =   5
      Left            =   4200
      TabIndex        =   132
      Top             =   1380
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Precio"
      Height          =   255
      Index           =   4
      Left            =   3060
      TabIndex        =   115
      Top             =   1980
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Importe"
      Height          =   255
      Index           =   3
      Left            =   1530
      TabIndex        =   114
      Top             =   1980
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "% Bon./Recargo"
      Height          =   255
      Index           =   2
      Left            =   210
      TabIndex        =   113
      Top             =   1980
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "Días"
      Height          =   255
      Left            =   3390
      TabIndex        =   101
      Top             =   1380
      Width           =   705
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Recibo"
      Height          =   255
      Index           =   28
      Left            =   2040
      TabIndex        =   66
      Top             =   720
      Width           =   855
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   4800
      ToolTipText     =   "Buscar Socio"
      Top             =   720
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Socio"
      Height          =   255
      Index           =   0
      Left            =   4230
      TabIndex        =   65
      Top             =   720
      Width           =   510
   End
   Begin VB.Image imgFec 
      Height          =   240
      Index           =   0
      Left            =   3840
      Picture         =   "frmPOZRecibos.frx":039E
      ToolTipText     =   "Buscar fecha"
      Top             =   690
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha "
      Height          =   255
      Index           =   29
      Left            =   3000
      TabIndex        =   64
      Top             =   720
      Width           =   780
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo Recibo"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   63
      Top             =   720
      Width           =   945
   End
   Begin VB.Label Label3 
      Caption         =   "Consumo"
      Height          =   255
      Left            =   1560
      TabIndex        =   62
      Top             =   1380
      Width           =   825
   End
   Begin VB.Label Label4 
      Caption         =   "Cuota"
      Height          =   255
      Left            =   2460
      TabIndex        =   61
      Top             =   1380
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Hidrante"
      Height          =   255
      Left            =   240
      TabIndex        =   60
      Top             =   1380
      Width           =   1035
   End
   Begin VB.Label Label6 
      Caption         =   "Concepto"
      Height          =   255
      Left            =   5460
      TabIndex        =   59
      Top             =   1980
      Width           =   1035
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
Attribute VB_Name = "frmPOZRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
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
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

Dim EsCabecera As Boolean
'Para saber en MandaBusquedaPrevia si busca en la tabla scapla o en la tabla sdirec
Dim NumTabMto As Integer 'Indica quin nº de Tab està en modo Mantenimient


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
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1


Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal
Dim indice As Byte
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
        Case 1  'BÚSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then InsertarCabecera
        
        Case 4  'MODIFICAR
            If DatosOk Then
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
        ' *** si n'hi han llínies ***
        
        Case 5 'LLÍNIES
            Select Case ModificaLineas
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


Private Sub cmdCancelar_Click()
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
            Select Case ModificaLineas
                Case 1 'afegir llínia
                    ModificaLineas = 0
                    ' *** les llínies que tenen datagrid (en o sense tab) ***
                    If NumTabMto = 0 Or NumTabMto = 1 Or NumTabMto = 2 Or NumTabMto = 4 Then
                        DataGridAux(NumTabMto).AllowAddNew = False
                        ' **** repasar si es diu Data1 l'adodc de la capçalera ***
                        'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
                        ' ********************************************************
                        LLamaLineas NumTabMto, ModificaLineas 'ocultar txtAux
                        DataGridAux(NumTabMto).Enabled = True
                        DataGridAux(NumTabMto).SetFocus

                        ' *** si n'hi han camps de descripció dins del grid, els neteje ***
                        'txtAux2(2).text = ""
                        ' *****************************************************************
                    End If
                    
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        AdoAux(NumTabMto).Recordset.MoveFirst
                    End If

                Case 2 'modificar llínies
                    ModificaLineas = 0
                    
                    
                    PonerModo 4
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de llínies ***
                        V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
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
Dim b As Boolean

    LimpiarCampos 'Huida els TextBox
    PonerModo 3
    
    '08/09/2010: numlinea
    Text1(31).Text = "1"
    
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la capçalera *******
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
        b = vSeccion.AbrirConta
        If b Then
            Text2(3).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(3).Text, "N")
            Text1(4).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", Text1(3).Text, "N")
            FormateaCampo Text1(4)
        End If
    End If
    Set vSeccion = Nothing
    
    Combo1(0).ListIndex = 0
    Combo1(0).SetFocus
'    PonerFoco Text1(1) '*** 1r camp visible que siga PK ***
    ' *** si n'hi han camps de descripció a la capçalera ***
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
            Text1(kCampo).BackColor = vbYellow
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
    DesplazamientoData Data1, Index
    PonerCampos
End Sub


Private Sub BotonModificar()
Dim DeVarios As Boolean

'    'solo se puede modificar la factura si no esta contabilizada
'    If FactContabilizada Then
'        TerminaBloquear
'        Exit Sub
'    End If
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    CargarValoresAnteriores Me, 1
    
    PonerFoco Text1(14) '*** 1r camp visible que siga PK ***
        
End Sub



Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Facturas (scafac)
' y los registros correspondientes de las tablas cab. albaranes (scafac1)
' y las lineas de la factura (slifac)
Dim Cad As String
'Dim vTipoMov As CTiposMov

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    
    Select Case Data1.Recordset.Fields(0).Value
        Case "RCP"
            Cad = "Recibo de Consumo." & vbCrLf
    
        Case "RMP"
            Cad = "Recibo de Mantenimiento." & vbCrLf
            
        Case "TAL"
            Cad = "Recibo de Talla." & vbCrLf
        
        Case "RVP"
            Cad = "Recibo de Contadores." & vbCrLf
        
    End Select
    
    Cad = Cad & "-------------------------------------      " & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar el Recibo del Socio:            "
    Cad = Cad & vbCrLf & "Nº Recibo:  " & Format(Text1(0).Text, "0000000")
    Cad = Cad & vbCrLf & "Fecha:  " & Format(Text1(1).Text, "dd/mm/yyyy")

    Cad = Cad & vbCrLf & vbCrLf & " ¿Desea Eliminarlo? "

    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
'        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
'        NumPedElim = Data1.Recordset.Fields(1).Value
        If Not eliminar Then
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
    '[Monica]05/05/2014: añadimos los recibos de consumo de manta
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
Dim Cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        If Not Data1.Recordset.EOF Then _
            Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        Cad = Data1.Recordset.Fields(0) & "|"
        Cad = Cad & Data1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(Cad)
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

     'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next kCampo

    ' ICONITOS DE LA BARRA
    btnPrimero = 13
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(4).Image = 3   'Insertar
        .Buttons(5).Image = 4   'Modificar
        .Buttons(6).Image = 5   'Borrar
        .Buttons(8).Image = 10  'Impresión de factura
        .Buttons(10).Image = 11  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    
    ' ******* si n'hi han llínies *******
    'ICONETS DE LES BARRES ALS TABS DE LLÍNIA
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
    
    '[Monica]08/05/2012: añadida Escalona que funciona como Utxera
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
        PonerModo 1 'búsqueda
        ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
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
Dim cadb As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        cadb = ""
        Aux = ValorDevueltoFormGrid(Text1(20), CadenaDevuelta, 1)
        cadb = cadb & Aux
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 2)
        cadb = cadb & " and " & Aux
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 3)
        cadb = cadb & " and " & Aux
        Aux = ValorDevueltoFormGrid(Text1(31), CadenaDevuelta, 4)
        cadb = cadb & " and " & Aux
        
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadb '& " " & Ordenacion
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
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000") 'Cod Socios
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Descripcion
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim vSeccion As CSeccion

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 1 'Tipo de IVA
            indice = 3
            PonerFoco Text1(indice)
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
            indice = 2
            PonerFoco Text1(indice)
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco Text1(indice)
            
            
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
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40

    Select Case Index
        Case 0
            indice = 1
        Case 1
            indice = 10
        Case 2
            indice = 11
        Case 3
            indice = 40
        Case 4
            indice = 42
        Case 5
            indice = 43
        Case 6
            indice = 47
    End Select
    
    imgFec(0).Tag = indice '<===
    If Text1(indice).Text <> "" Then frmC.NovaData = Text1(indice).Text
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
        indice = 15
        frmZ.pTitulo = "Observaciones del Albarán"
        frmZ.pValor = Text1(indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco Text1(indice)
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
'        If MsgBox("Este albarán está facturado y/o cobrado. ¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
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
Dim sql As String
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
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
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
                        MsgBox "El socio está dado de baja. Reintroduzca.", vbExclamation
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
                    sql = "select poligono, parcelas, nroorden from rpozos where hidrante = " & DBSet(Text1(Index).Text, "T")
                    Set Rs = New ADODB.Recordset
                    Rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
Dim cadb As String
Dim cadAux As String
    
'    '--- Laura 12/01/2007
'    cadAux = Text1(5).Text
'    If Text1(4).Text <> "" Then Text1(5).Text = ""
'    '---
    
'    '--- Laura 12/01/2007
'    Text1(5).Text = cadAux
'    '---
'    CadB = ObtenerBusqueda(Me)
    cadb = ObtenerBusqueda2(Me, BuscaChekc, 1)

    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia cadb
    ElseIf cadb <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select rrecibpozos.* from " & NombreTabla
        CadenaConsulta = CadenaConsulta & " WHERE " & cadb & Ordenacion
'        CadenaConsulta = "select palets.* from " & NombreTabla
'        CadenaConsulta = CadenaConsulta & " WHERE " & CadB & " GROUP BY palets.numpalet " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadb As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim Tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
    'Llamamos a al form
    '##A mano
    Cad = ""
'    Cad = Cad & "Tipo|if(rfactsoc.codtipom='FAA','Anticipo','Liquidación') as a|T||10·"
    Cad = Cad & "Tipo Fichero|case rrecibpozos.codtipom when ""RCP"" then ""RCP-Consumo"" when ""RMP"" then ""RMP-Mantenim."" end as tipo|N||22·"
    Cad = Cad & "Tipo|rrecibpozos.codtipom|N||6·" ' ParaGrid(Combo1(0), 0, "Tipo")
    Cad = Cad & "Nº.Factura|rrecibpozos.numfactu|N||12·"
    Cad = Cad & "Fecha|rrecibpozos.fecfactu|F||15·"
    Cad = Cad & "Lin|rrecibpozos.numlinea|N||6·"
    Cad = Cad & "Código|rrecibpozos.codsocio|N|000000|12·"
    Cad = Cad & "Socio|rsocios.nomsocio|N||38·"
    
    Tabla = NombreTabla & " inner join rsocios on rrecibpozos.codsocio = rsocios.codsocio "
    Titulo = "Recibos de Contadores"
    devuelve = "1|2|3|4|"
           
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = Tabla
        frmB.vSQL = cadb
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vSelElem = 0
'        frmB.vConexionGrid = cAgro  'Conexión a BD: Ariagro
        If Not EsCabecera Then frmB.Label1.FontSize = 11
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
'        If EsCabecera Then
'            PonerCadenaBusqueda
'            Text1(0).Text = Format(Text1(0).Text, "0000000")
'        End If
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos Then
            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                cmdRegresar_Click
        Else   'de ha devuelto datos, es decir NO ha devuelto datos
            PonerFoco Text1(kCampo)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass

    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
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
Dim b As Boolean
Dim vSeccion As CSeccion

    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1

    PosicionarCombo2 Combo1(0), Text1(20).Text

    CargaGrid 2, True
    If Not AdoAux(2).Recordset.EOF Then _
        PonerCamposForma2 Me, AdoAux(2), 2, "FrameAux2"
    
    CargaGrid 0, True
    If Not AdoAux(0).Recordset.EOF Then _
        PonerCamposForma2 Me, AdoAux(0), 2, "FrameAux0"
    
    CargaGrid 1, True
    If Not AdoAux(1).Recordset.EOF Then _
        PonerCamposForma2 Me, AdoAux(1), 2, "FrameAux1"
    
    
    
'    cmdConceptos_Click
    
    cmdConceptos.visible = (Text1(20).Text = "RVP")
    cmdConceptos.Enabled = (Text1(20).Text = "RVP")
  
   
    cmdParticipa.visible = (Text1(20).Text = "RCP") And vParamAplic.Cooperativa = 1
    cmdParticipa.Enabled = (Text1(20).Text = "RCP") And vParamAplic.Cooperativa = 1
   
    cmdHidrantes.visible = (Text1(20).Text = "RMP") And vParamAplic.Cooperativa = 10
    cmdHidrantes.Enabled = (Text1(20).Text = "RMP") And vParamAplic.Cooperativa = 10
    
    cmdCampos.visible = (Text1(20).Text = "TAL" Or Text1(20).Text = "RMT") And vParamAplic.Cooperativa = 10
    cmdCampos.Enabled = (Text1(20).Text = "TAL" Or Text1(20).Text = "RMT") And vParamAplic.Cooperativa = 10
   
   
    '[Monica]15/09/2015: si es Escalona indicamos si el recibo está pagado
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
        b = vSeccion.AbrirConta
        If b Then
            Text2(3).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Text1(3).Text, "N")
        End If
    End If
    Set vSeccion = Nothing
'    MostrarCadena Text1(3), Text1(4)
    
    Modo = 2
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario

    If Err.Number <> 0 Then Err.Clear
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim I As Byte, Numreg As Byte
Dim b As Boolean
Dim b1 As Boolean

    On Error GoTo EPonerModo

    BuscaChekc = ""

    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Or hcoCodMovim <> "" Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
    
    
    'Poner Flechas de desplazamiento visibles
    Numreg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then Numreg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, Numreg
          
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    BloquearCombo Me, Modo
    
    '+++ bloqueamos el combo1(0) como si tuviera tag
    b1 = (Modo = 3 Or Modo = 4 Or Modo = 1) '06/09/2005, lleve el modo 5 per a que no es puga modificar la capçalera mentre treballe en les llínies
    
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
    
    b = (Modo <> 1)
    'Campos Nº Recibo bloqueado y en azul
    BloquearTxt Text1(0), b, True
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2)
    cmdCancelar.visible = b
    CmdAceptar.visible = b
    
    BloquearImgBuscar Me, Modo, ModificaLineas
    BloquearImgFec Me, 0, Modo
    
    b = (Modo = 1 Or Modo = 3)
    Text1(1).Enabled = b
    Text1(2).Enabled = b Or Modo = 4
    imgFec(0).Enabled = b
    imgFec(0).visible = b
    imgBuscar(0).Enabled = b Or Modo = 4
    imgBuscar(0).visible = b Or Modo = 4
    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    BloquearTxt Text1(4), (Modo <> 1)
    BloquearTxt Text1(6), (Modo <> 1)
       
    For I = 1 To 2
        imgFec(I).Enabled = (Modo = 1 Or Modo = 3 Or Modo = 4)
    Next I
        
    ' ***************************
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
'lineas
    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 2, False
        CargaGrid 0, False
        CargaGrid 1, False
    End If
    
    b = (Modo = 4) Or (Modo = 2)
    DataGridAux(2).Enabled = b
    
    b = (Modo = 5)
    For I = 1 To 3
        BloquearTxt txtAux3(I), Not b
    Next I
    b = (Modo = 5) And ModificaLineas = 2
    BloquearTxt txtAux3(1), b
    
    b = (Modo = 5)
    For I = 4 To 10
        BloquearTxt txtAux5(I), Not b
    Next I
    
    '[Monica]15/09/2015: ponemos la situacion
    Label1(13).visible = ((Modo = 2 Or Modo = 4) And vParamAplic.Cooperativa = 10)
    
    
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim b As Boolean
Dim sql As String

    On Error GoTo EDatosOK

    DatosOk = False
    
'    ComprobarDatosTotales

    'comprobamos datos OK de la tabla scafac
    b = CompForm2(Me, 1)  'Comprobar formato datos ok de la cabecera: opcion=1
    If Not b Then Exit Function
    
    If b Then
        If (Modo = 3 Or Modo = 4) And Combo1(0).ListIndex = 0 Then
            '[Monica]17/11/2014: obligamos a que si es de consumo metan el hidrante sólo para escalona y utxera
            If Text1(14).Text = "" Then
                If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
                    MsgBox "El hidrante no puede estar vacio en un recibo de consumo. Revise.", vbExclamation
                    PonerFoco Text1(14)
                    b = False
                End If
            Else
            
                '[Monica]17/11/2014: si el hidrante no existe evitamos el error de clave referencial
                If b Then
                    sql = DevuelveDesdeBDNew(cAgro, "rpozos", "hidrante", "hidrante", Text1(14).Text, "T")
                    If sql = "" Then
                        MsgBox "El Hidrante no existe. Revise.", vbExclamation
                        PonerFoco Text1(14)
                        b = False
                    End If
                End If
            
            
                ' comprobamos si insertamos o modificamos que existe el hidrante para el socio
                If b Then
                    sql = ""
                    sql = DevuelveDesdeBDNew(cAgro, "rpozos", "hidrante", "hidrante", Text1(14).Text, "T", , "codsocio", Text1(2).Text, "N")
                    
                    If sql = "" Then
                        If MsgBox("El Hidrante no es del socio introducido. " & vbCrLf & vbCrLf & "¿ Desea continuar ? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                            PonerFoco Text1(14)
                            b = False
                        Else
                            b = True
                        End If
                    End If
                End If
                    
            End If
        End If
    End If
    
    If b Then
        If Modo = 3 Or Modo = 4 Then
            If Text1(11).Text <> "" And Text1(10).Text <> "" Then
                If CDate(Text1(11).Text) > CDate(Text1(10).Text) Then
                    MsgBox "La Fecha de Lectura Anterior no puede ser superior a la de Lectura Actual. Revise.", vbExclamation
                    PonerFoco Text1(11)
                    b = False
                End If
            End If
            
            If b Then
                If Text1(8).Text <> "" And Text1(9).Text <> "" Then
                    If CLng(Text1(8).Text) > CLng(Text1(9).Text) Then
                        MsgBox "El Contador Anterior no puede ser superior al del Contador Actual. Revise.", vbExclamation
                        PonerFoco Text1(8)
                        b = False
                    End If
                End If
            End If
        End If
    End If
    
    DatosOk = b
    
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
        Case 1  'Buscar
            mnBuscar_Click
        Case 2  'Todos
            BotonVerTodos
        Case 4  'Añadir
            mnNuevo_Click
        Case 5  'Modificar
            mnModificar_Click
        Case 6  'Borrar
            mnEliminar_Click
        Case 8  ' Impresion de albaran
            mnImprimir_Click
        Case 10   'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
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

Private Function eliminar() As Boolean
Dim sql As String, LEtra As String, Sql2 As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String

    On Error GoTo FinEliminar

    If Data1.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    'Eliminar en tablas de cabecera de factura
    '------------------------------------------
    sql = " " & ObtenerWhereCP(True)
    
    conn.Execute "delete from rrecibpozos_acc " & sql
    
    conn.Execute "delete from rrecibpozos_hid " & sql
    
    conn.Execute "delete from rrecibpozos_cam " & sql
    
    
    'Cabecera de factura (rrecibpozos)
    conn.Execute "Delete from " & NombreTabla & sql
    
    
    CadenaCambio = "DELETE FROM " & NombreTabla & sql
    
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
    
    b = True
FinEliminar:
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, "Eliminar Recibo", Err.Description & " " & Mens
        b = False
    End If
    If Not b Then
        conn.RollbackTrans
        eliminar = False
    Else
        conn.CommitTrans
        eliminar = True
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
Dim sql As String

    On Error Resume Next
    
    sql = " codtipom= '" & Text1(20).Text & "'"
    sql = sql & " and numfactu = " & Text1(0).Text
    sql = sql & " and fecfactu = " & DBSet(Text1(1).Text, "F")
    '08/09/2010 : añadido a la clave primaria
    sql = sql & " and numlinea = " & DBSet(Text1(31).Text, "N")

    If conWhere Then sql = " WHERE " & sql
    ObtenerWhereCP = sql
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function



Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean, bAux As Boolean
Dim I As Integer

        b = ((Modo = 2) Or (Modo = 0)) And (hcoCodMovim = "") 'Or (Modo = 5 And ModificaLineas = 0)
        'Buscar
        Toolbar1.Buttons(1).Enabled = b
        Me.mnBuscar.Enabled = b
        'Vore Tots
        Toolbar1.Buttons(2).Enabled = b
        Me.mnVerTodos.Enabled = b
        'Añadir
        Toolbar1.Buttons(4).Enabled = b
        Me.mnModificar.Enabled = b
        
        
        b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And (hcoCodMovim = "") And Not (Check1(1).Value = 1)
        'Modificar
        Toolbar1.Buttons(5).Enabled = b
        Me.mnModificar.Enabled = b
        'eliminar
        Toolbar1.Buttons(6).Enabled = b
        Me.mnEliminar.Enabled = b
        'Impresión de albaran
        Toolbar1.Buttons(8).Enabled = (Modo = 2) And Data1.Recordset.RecordCount > 0
        Me.mnImprimir.Enabled = (Modo = 2) And Data1.Recordset.RecordCount > 0
        

    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And DatosADevolverBusqueda = "" And Check1(1).Value = 0
    For I = 0 To 2
        ToolAux(I).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.AdoAux(I).Recordset.RecordCount > 0)
        ToolAux(I).Buttons(2).Enabled = bAux
        ToolAux(I).Buttons(3).Enabled = bAux
    Next I
    ' ****************************************


End Sub


Private Sub BotonImprimir()
Dim cadFormula As String
Dim cadParam As String
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
    cadParam = ""
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
    
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
    
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
    'Cadena para seleccion Nº de recibo
    '---------------------------------------------------
    If Text1(0).Text <> "" Then
        'tipo de fichero
        devuelve = "{" & NombreTabla & ".codtipom}='" & Text1(20).Text & "'"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
        
        'Nº factura
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
            cadParam = cadParam & "pDuplicado=1|"
            numParam = numParam + 1
        End If
    End If
    
    If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub
     
    With frmImprimir
          '[Monica]06/02/2012: añadido la siguientes 3 lineas para el envio por el outlook
            .outClaveNombreArchiv = Mid(Combo1(0).Text, 1, 3) & Format(Text1(0).Text, "0000000")
            .outCodigoCliProv = Text1(2).Text
            .outTipoDocumento = 100
            
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 0
            .Titulo = "Impresión de Recibos de Socios"
            
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
Dim sql As String
Dim vSeccion As CSeccion
Dim Rs As ADODB.Recordset

    ReciboCobrado = False

    If Check1(1).Value = 0 Then Exit Function
    

    Set vSeccion = New CSeccion
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If vSeccion.AbrirConta Then
    
            sql = "SELECT count(*) FROM scobro INNER JOIN usuarios.stipom ON scobro.numserie = stipom.letraser "
            sql = sql & " WHERE stipom.codtipom = " & DBSet(TipoM, "T")
            sql = sql & " and scobro.codfaccl = " & DBSet(numfactu, "N")
            sql = sql & " and scobro.fecfaccl = " & DBSet(fecfactu, "F")
            
            Set Rs = New ADODB.Recordset
            Rs.Open sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If Not Rs.EOF Then
                If Rs.Fields(0).Value = 0 Then
                    ReciboCobrado = True
                    Exit Function
                End If
            End If
            Set Rs = Nothing
            
            
            sql = "SELECT sum(coalesce(impvenci,0) + coalesce(gastos,0) - coalesce(impcobro,0))  FROM scobro INNER JOIN usuarios.stipom ON scobro.numserie = stipom.letraser "
            sql = sql & " WHERE stipom.codtipom = " & DBSet(TipoM, "T")
            sql = sql & " and scobro.codfaccl = " & DBSet(numfactu, "N")
            sql = sql & " and scobro.fecfaccl = " & DBSet(fecfactu, "F")
            Set Rs = New ADODB.Recordset
            Rs.Open sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
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
Dim sql As String
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
    
End Sub


Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim sql As String

    On Error GoTo EInsertarCab
    
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        sql = CadenaInsertarDesdeForm(Me)
        If sql <> "" Then
            If InsertarOferta(sql, vTipoMov) Then
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
Dim sql As String
Dim NumF As String

    On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Factura
    'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
    Do
        '[Monica]12/06/2014: en el caso de escalona no hay cambio de campaña
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
    
    '[Monica]19/11/2013: añadido el log de que ha insertado
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
Dim Cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim Rs As ADODB.Recordset


    If Text1(2).Text = "" Then Exit Sub
    
    If Not (Modo = 3) Then Exit Sub

    Cad = "rpozos.codsocio = " & DBSet(Text1(2).Text, "N")
     
    Cad1 = "select count(*) from rpozos where " & Cad
     
    NumRegis = TotalRegistros(Cad1)
    
    If NumRegis = 0 Then Exit Sub
    If NumRegis = 1 Then
        Cad1 = "select hidrante, rpozos.codparti, rpartida.nomparti, rpozos.poligono from rpozos inner join rpartida on rpozos.codparti = rpartida.codparti where " & Cad
        Set Rs = New ADODB.Recordset
        Rs.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            Text1(14).Text = DBLet(Rs.Fields(0).Value)
        End If
    Else
        Set frmMens = New frmMensajes
        frmMens.cadwhere = " and " & Cad
        frmMens.campo = Text1(14).Text
        frmMens.OpcionMensaje = 23
        frmMens.Show vbModal
        Set frmMens = Nothing
    End If
    
End Sub




Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
Dim b As Boolean
Dim I As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)
    
    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    Select Case Index
        Case 2 'pozos
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;S|txtAux3(1)|T|Fases|900|;" 'codsocio,numfase
            tots = tots & "S|txtAux3(2)|T|Acciones|1200|;"
            tots = tots & "S|txtAux3(3)|T|Observaciones|5280|;"
            arregla tots, DataGridAux(Index), Me
        
'            DataGridAux(Index).Columns(3).Alignment = dbgLeft
        
            b = (Modo = 4) And ((ModificaLineas = 1) Or (ModificaLineas = 2))

        Case 0 'hidrantes
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;S|txtAux4(4)|T|Hidrante|1300|;" 'codsocio,numfase
            tots = tots & "S|txtAux4(5)|T|Hanegadas|1200|;"
            arregla tots, DataGridAux(Index), Me
        
'            DataGridAux(Index).Columns(3).Alignment = dbgLeft
        
            b = (Modo = 4) And ((ModificaLineas = 1) Or (ModificaLineas = 2))

        Case 1 'campos
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;S|txtAux5(4)|T|Campo|1300|;" 'codsocio,numfase
            tots = tots & "S|txtAux5(5)|T|Hanegadas|1200|;S|txtAux5(6)|T|Pr/Ha.Cuota|1200|;S|txtAux5(7)|T|Pr/Ha.Ordin.|1200|;"
            tots = tots & "S|txtAux5(8)|T|Poligono|1100|;S|txtAux5(9)|T|Parcela|1100|;S|txtAux5(10)|T|SParcela|850|;"
            arregla tots, DataGridAux(Index), Me
        
'            DataGridAux(Index).Columns(3).Alignment = dbgLeft
        
            b = (Modo = 4) And ((ModificaLineas = 1) Or (ModificaLineas = 2))


    End Select
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
'    ' **** si n'hi han llínies en grids i camps fora d'estos ****
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
'   Basant-se en la informació proporcionada pel vector de camps
'   crea un SQl per a executar una consulta sobre la base de datos que els
'   torne.
' Si ENLAZA -> Enlaça en el data1
'           -> Si no el carreguem sense enllaçar a cap camp
'--------------------------------------------------------------------
Dim sql As String
Dim Tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
       Case 0 ' hidrantes
            Tabla = "rrecibpozos_hid"
            sql = "SELECT codtipom,numfactu,fecfactu,numlinea,hidrante, hanegada "
            sql = sql & " FROM " & Tabla
            If enlaza Then
                sql = sql & ObtenerWhereCab(True)
            Else
                sql = sql & " WHERE numfactu = -1"
            End If
            sql = sql & " ORDER BY " & Tabla & ".hidrante "
       
       Case 2 ' pozos
            Tabla = "rrecibpozos_acc"
            sql = "SELECT codtipom,numfactu,fecfactu,numlinea,numfases, acciones,observac "
            sql = sql & " FROM " & Tabla
            If enlaza Then
                sql = sql & ObtenerWhereCab(True)
            Else
                sql = sql & " WHERE numfactu = -1"
            End If
            sql = sql & " ORDER BY " & Tabla & ".numfases "
            
            
       Case 1 ' campos
            Tabla = "rrecibpozos_cam"
            sql = "SELECT codtipom,numfactu,fecfactu,numlinea,codcampo, hanegada, precio1, precio2, poligono, parcela, subparce "
            sql = sql & " FROM " & Tabla
            If enlaza Then
                sql = sql & ObtenerWhereCab(True)
            Else
                sql = sql & " WHERE numfactu = -1"
            End If
            sql = sql & " ORDER BY " & Tabla & ".codcampo "
            
            
    End Select
    ' ********************************************************************************
    
    MontaSQLCarga = sql
End Function

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " codtipom=" & DBSet(Mid(Combo1(0).Text, 1, 3), "T") & " and numfactu = " & Text1(0).Text & _
                      " and fecfactu = " & DBSet(Text1(1).Text, "F") & " and numlinea = " & Text1(31).Text
    ' *******************************************************
    ObtenerWhereCab = vWhere
End Function

Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vTabla As String
Dim anc As Single
Dim I As Integer
    
    ModificaLineas = 1 'Posem Modo Afegir Llínia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModificaLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    ' **************************************************

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 0: vTabla = "rrecibpozos_hid"
        Case 1: vTabla = "rrecibpozos_campos"
        Case 2: vTabla = "rrecibpozos_acc"
    End Select
    ' ********************************************************
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0

            AnyadirLinea DataGridAux(Index), AdoAux(Index)
    
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

            AnyadirLinea DataGridAux(Index), AdoAux(Index)
    
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

            AnyadirLinea DataGridAux(Index), AdoAux(Index)
    
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
Dim sql As String
Dim vWhere As String
Dim eliminar As Boolean

    On Error GoTo Error2

    ModificaLineas = 3 'Posem Modo Eliminar Llínia
    
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModificaLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index

    If AdoAux(Index).Recordset.EOF Then Exit Sub
'    If Not SepuedeBorrar(Index) Then Exit Sub
    NumTabMto = Index
    eliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 0 'hidrantes
            sql = "¿Seguro que desea eliminar el registro?"
            sql = sql & vbCrLf & "Hidrante: " & AdoAux(Index).Recordset!Hidrante
            If MsgBox(sql, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                sql = "DELETE FROM rrecibpozos_hid"
                sql = sql & vWhere & " AND hidrante= " & DBLet(AdoAux(Index).Recordset!Hidrante, "T")
                
                
                CadenaCambio = sql
                '------------------------------------------------------------------------------
                '  LOG de acciones
                ValorAnterior = ""
                Set LOG = New cLOG
                LOG.InsertarCambiosRegistros 14, vUsu, "Cambio Recibos Pozos-D", "rrecibpozos_hid", ObtenerWhereCab(False)
                Set LOG = Nothing
                '-----------------------------------------------------------------------------
            End If
        
        Case 1 'campos
            sql = "¿Seguro que desea eliminar el registro?"
            sql = sql & vbCrLf & "Campos: " & AdoAux(Index).Recordset!codcampo
            If MsgBox(sql, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                sql = "DELETE FROM rrecibpozos_cam"
                sql = sql & vWhere & " AND codcampo= " & DBLet(AdoAux(Index).Recordset!codcampo, "N")
                
                CadenaCambio = sql
                '------------------------------------------------------------------------------
                '  LOG de acciones
                ValorAnterior = ""
                Set LOG = New cLOG
                LOG.InsertarCambiosRegistros 14, vUsu, "Cambio Recibos Pozos-D", "rrecibpozos_cam", ObtenerWhereCab(False)
                Set LOG = Nothing
                '-----------------------------------------------------------------------------
            End If
        
        
        Case 2 'pozos
            sql = "¿Seguro que desea eliminar el registro?"
            sql = sql & vbCrLf & "Numero Fase: " & AdoAux(Index).Recordset!numfases
            If MsgBox(sql, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                sql = "DELETE FROM rrecibpozos_acc"
                sql = sql & vWhere & " AND numfases= " & DBLet(AdoAux(Index).Recordset!numfases, "N")
            
                CadenaCambio = sql
                '------------------------------------------------------------------------------
                '  LOG de acciones
                ValorAnterior = ""
                Set LOG = New cLOG
                LOG.InsertarCambiosRegistros 14, vUsu, "Cambio Recibos Pozos-D", "rrecibpozos_acc", ObtenerWhereCab(False)
                Set LOG = Nothing
                '-----------------------------------------------------------------------------
            End If
        
    End Select

    If eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        conn.Execute sql
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        If Index <> 3 Then _
            CargaGrid Index, True
        ' ***************************************************
        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
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
    
    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub
    
    ModificaLineas = 2 'Modificar llínia
       
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModificaLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    ' *** bloqueje la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    ' *********************************
    
    Select Case Index
        Case 0, 1, 2 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
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
Dim b As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    If Index <> 3 Then DeseleccionaGrid DataGridAux(Index)
    ' ***************************************************
       
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 0 ' hidrantes
            For jj = 4 To 5
                txtAux4(jj).visible = b
                txtAux4(jj).Top = alto
            Next jj
            If xModo = 2 Then txtAux4(4).visible = False
        Case 1 ' campos
            For jj = 4 To 10
                txtAux5(jj).visible = b
                txtAux5(jj).Top = alto
            Next jj
        
        Case 2 ' pozos
            For jj = 1 To 3
                txtAux3(jj).visible = b
                txtAux3(jj).Top = alto
            Next jj
    End Select
End Sub


Private Sub InsertarLinea()
'Inserta registre en les taules de Llínies
Dim nomframe As String
Dim b As Boolean

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
            b = BLOQUEADesdeFormulario2(Me, Data1, 1)
            
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
                V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
                CargaGrid NumTabMto, True
            End If

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


Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim sql As Integer
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
    
    
    If b And NumTabMto = 2 And ModificaLineas = 1 Then
        sql = DevuelveValor("select acciones from rrecibpozos_acc where codtipom = " & DBSet(txtAux3(0).Text, "T") & " and numfactu = " & DBSet(txtAux3(4).Text, "N") & " and fecfactu = " & DBSet(txtAux3(5).Text, "F") & " and numlinea = " & DBSet(txtAux3(6).Text, "N") & " and numfase = " & DBSet(txtAux3(1).Text, "N"))
        If sql <> 0 Then
            MsgBox "El número de fase ya existe. Reintroduzca.", vbExclamation
            b = False
            PonerFoco txtAux3(1)
        End If
    End If
    
    If b And NumTabMto = 0 And ModificaLineas = 1 Then
        sql = DevuelveValor("select count(*) from rpozos where hidrante = " & DBSet(txtAux4(4).Text, "T"))
        If sql = 0 Then
            MsgBox "El hidrante no existe. Reintroduzca.", vbExclamation
            b = False
            PonerFoco txtAux4(4)
        End If
        If b Then
            sql = DevuelveValor("select count(*) from rrecibpozos_hid where codtipom = " & DBSet(txtAux4(0).Text, "T") & " and numfactu = " & DBSet(txtAux4(1).Text, "N") & " and fecfactu = " & DBSet(txtAux4(2).Text, "F") & " and numlinea = " & DBSet(txtAux4(3).Text, "N") & " and hidrante = " & DBSet(txtAux4(4).Text, "T"))
            If sql <> 0 Then
                MsgBox "El hidrante ya existe en el recibo. Revise.", vbExclamation
                b = False
                PonerFoco txtAux4(4)
            End If
        End If
    End If
    
    If b And NumTabMto = 1 And ModificaLineas = 1 Then
        sql = DevuelveValor("select count(*) from rcampos where codcampo = " & DBSet(txtAux5(4).Text, "N"))
        If sql = 0 Then
            MsgBox "El campo no existe. Reintroduzca.", vbExclamation
            b = False
            PonerFoco txtAux5(4)
        End If
        If b Then
            sql = DevuelveValor("select count(*) from rrecibpozos_cam where codtipom = " & DBSet(txtAux5(0).Text, "T") & " and numfactu = " & DBSet(txtAux5(1).Text, "N") & " and fecfactu = " & DBSet(txtAux5(2).Text, "F") & " and numlinea = " & DBSet(txtAux5(3).Text, "N") & " and codcampo = " & DBSet(txtAux5(4).Text, "T"))
            If sql <> 0 Then
                MsgBox "El campo ya existe en el recibo. Revise.", vbExclamation
                b = False
                PonerFoco txtAux5(4)
            End If
        End If
    End If
    
    
    
    
    DatosOkLlin = b

EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

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
Dim Cadena As String
    
    If Not PerderFocoGnral(txtAux4(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 5
            PonerFormatoDecimal txtAux4(Index), 11
        
            CmdAceptar.SetFocus
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
Dim Cadena As String
    
    If Not PerderFocoGnral(txtAux5(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 5, 6, 7
            PonerFormatoDecimal txtAux5(Index), 11
        
        Case 10
            CmdAceptar.SetFocus
    End Select
    ' ******************************************************************************
End Sub


Private Sub Timer1_Timer()
    Label1(13).visible = Not Label1(13).visible
End Sub

