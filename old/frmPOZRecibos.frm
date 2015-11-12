VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPOZRecibos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico de Recibos "
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   9585
   Icon            =   "frmPOZRecibos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox cmdConceptos 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   8580
      ScaleHeight     =   435
      ScaleWidth      =   525
      TabIndex        =   57
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
      TabIndex        =   41
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
      Left            =   1560
      MaxLength       =   7
      TabIndex        =   6
      Tag             =   "Consumo|N|S|||rrecibpozos|consumo|||"
      Text            =   "1234567"
      Top             =   1650
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Height          =   555
      Index           =   15
      Left            =   4170
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Tag             =   "Conceptol|T|S|||rrecibpozos|concepto|||"
      Text            =   "frmPOZRecibos.frx":000C
      Top             =   1650
      Width           =   4965
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   13
      Left            =   2970
      MaxLength       =   10
      TabIndex        =   7
      Tag             =   "Cuota|N|S|||rrecibpozos|impcuota|###,##0.00||"
      Text            =   "1234567"
      Top             =   1650
      Width           =   1065
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
      TabIndex        =   40
      Text            =   "Text2"
      Top             =   990
      Width           =   4110
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Contabilizado"
      Height          =   195
      Index           =   1
      Left            =   2910
      TabIndex        =   10
      Tag             =   "Contabilizado|N|N|0|1|rrecibpozos|contabilizado|0||"
      Top             =   2010
      Width           =   1245
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Impreso"
      Height          =   195
      Index           =   0
      Left            =   1590
      TabIndex        =   9
      Tag             =   "Impreso|N|N|0|1|rrecibpozos|impreso|0||"
      Top             =   2010
      Width           =   915
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
      TabIndex        =   29
      Top             =   3540
      Width           =   3135
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   1500
         MaxLength       =   8
         TabIndex        =   70
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
         TabIndex        =   69
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
         TabIndex        =   68
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
         TabIndex        =   71
         Tag             =   "Consumo 2|N|S|||rrecibpozos|consumo2|0000000||"
         Text            =   "m3"
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label Label14 
         Caption         =   "Precio"
         Height          =   285
         Left            =   300
         TabIndex        =   73
         Top             =   1890
         Width           =   1245
      End
      Begin VB.Label Label13 
         Caption         =   "Hasta m3."
         Height          =   285
         Left            =   300
         TabIndex        =   72
         Top             =   1500
         Width           =   1245
      End
      Begin VB.Label Label28 
         Caption         =   "Hasta m3."
         Height          =   285
         Left            =   300
         TabIndex        =   31
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label Label27 
         Caption         =   "Precio"
         Height          =   285
         Left            =   300
         TabIndex        =   30
         Top             =   870
         Width           =   1245
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
      TabIndex        =   26
      Top             =   3540
      Width           =   2925
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   1500
         MaxLength       =   10
         TabIndex        =   67
         Tag             =   "Fecha Lectura Actual|F|S|||rrecibpozos|fech_act|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   1170
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   1500
         MaxLength       =   7
         TabIndex        =   66
         Tag             =   "Contador Actual|N|S|||rrecibpozos|lect_act|0000000||"
         Text            =   "1234567"
         Top             =   630
         Width           =   1065
      End
      Begin VB.Label Label9 
         Caption         =   "Contador"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   660
         Width           =   1035
      End
      Begin VB.Label Label12 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   1200
         Width           =   705
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1140
         Picture         =   "frmPOZRecibos.frx":0014
         ToolTipText     =   "Buscar fecha"
         Top             =   1200
         Width           =   240
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
      Left            =   240
      TabIndex        =   23
      Top             =   3540
      Width           =   2895
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   1590
         MaxLength       =   7
         TabIndex        =   65
         Tag             =   "Lectura Anterior|N|S|||rrecibpozos|lect_ant|0000000||"
         Text            =   "1234567"
         Top             =   630
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   1590
         MaxLength       =   10
         TabIndex        =   64
         Tag             =   "Fecha lectura anterior|F|S|||rrecibpozos|fech_ant|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   1170
         Width           =   1065
      End
      Begin VB.Label Label23 
         Caption         =   "Contador"
         Height          =   285
         Left            =   330
         TabIndex        =   25
         Top             =   660
         Width           =   1125
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1200
         Picture         =   "frmPOZRecibos.frx":009F
         ToolTipText     =   "Buscar fecha"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label18 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   330
         TabIndex        =   24
         Top             =   1200
         Width           =   570
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   210
      TabIndex        =   14
      Top             =   5910
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
         TabIndex        =   15
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8280
      TabIndex        =   12
      Top             =   5985
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7080
      TabIndex        =   11
      Top             =   6000
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   9585
      _ExtentX        =   16907
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
      Left            =   8250
      TabIndex        =   13
      Top             =   5970
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   2880
      Top             =   6030
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
      Left            =   2850
      Top             =   6030
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
      Top             =   6060
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
      Height          =   1065
      Index           =   0
      Left            =   180
      TabIndex        =   17
      Top             =   2370
      Width           =   9105
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   5370
         MaxLength       =   10
         TabIndex        =   63
         Tag             =   "Importe Iva|N|S|||rrecibpozos|imporiva|###,##0.00||"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   1620
         MaxLength       =   2
         TabIndex        =   62
         Tag             =   "Tipo Iva|N|S|||rrecibpozos|tipoiva|00||"
         Text            =   "Text1"
         Top             =   600
         Width           =   600
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   4560
         MaxLength       =   6
         TabIndex        =   61
         Tag             =   "Porc.Iva|N|S|||rrecibpozos|porc_iva|##0.00||"
         Text            =   "123"
         Top             =   600
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
         TabIndex        =   60
         Text            =   "Text2"
         Top             =   600
         Width           =   2190
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   180
         MaxLength       =   10
         TabIndex        =   59
         Tag             =   "Base Imponible|N|N|||rrecibpozos|baseimpo|###,##0.00||"
         Top             =   600
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CAE3FD&
         Height          =   315
         Index           =   7
         Left            =   7140
         MaxLength       =   10
         TabIndex        =   58
         Tag             =   "Total Factura|N|N|||rrecibpozos|totalfact|###,##0.00||"
         Top             =   600
         Width           =   1830
      End
      Begin VB.Label Label2 
         Caption         =   "% Iva"
         Height          =   255
         Left            =   4530
         TabIndex        =   22
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Iva"
         Height          =   255
         Index           =   8
         Left            =   1620
         TabIndex        =   21
         Top             =   360
         Width           =   285
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1950
         ToolTipText     =   "Buscar Iva"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Importe IVA"
         Height          =   255
         Index           =   7
         Left            =   5400
         TabIndex        =   20
         Top             =   360
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
         TabIndex        =   19
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
         Height          =   255
         Index           =   10
         Left            =   180
         TabIndex        =   18
         Top             =   360
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
      TabIndex        =   74
      Tag             =   "Linea|N|N|||rrecibpozos|numlinea|0000000|S|"
      Text            =   "Text1"
      Top             =   1650
      Width           =   885
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
      TabIndex        =   42
      Top             =   3510
      Width           =   9105
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   30
         Left            =   7680
         MaxLength       =   100
         TabIndex        =   52
         Tag             =   "Importe Art.4|N|S|||rrecibpozos|importear4|###,##0.00||"
         Top             =   2010
         Width           =   1260
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   29
         Left            =   240
         MaxLength       =   100
         TabIndex        =   51
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
         TabIndex        =   50
         Tag             =   "Importe Art.3|N|S|||rrecibpozos|importear3|###,##0.00||"
         Top             =   1710
         Width           =   1260
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   27
         Left            =   240
         MaxLength       =   100
         TabIndex        =   49
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
         TabIndex        =   48
         Tag             =   "Importe Art.2|N|S|||rrecibpozos|importear2|###,##0.00||"
         Top             =   1410
         Width           =   1260
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   25
         Left            =   240
         MaxLength       =   100
         TabIndex        =   47
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
         TabIndex        =   46
         Tag             =   "Importe Art.1|N|S|||rrecibpozos|importear1|###,##0.00||"
         Top             =   1110
         Width           =   1260
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   23
         Left            =   240
         MaxLength       =   100
         TabIndex        =   45
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
         TabIndex        =   44
         Tag             =   "Importe MO|N|S|||rrecibpozos|importemo|###,##0.00||"
         Top             =   540
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   21
         Left            =   240
         MaxLength       =   100
         TabIndex        =   43
         Tag             =   "Concepto MO|T|S|||rrecibpozos|conceptomo|||"
         Text            =   "1234567"
         Top             =   540
         Width           =   7245
      End
      Begin VB.Label Label11 
         Caption         =   "Importe"
         Height          =   255
         Left            =   7710
         TabIndex        =   56
         Top             =   810
         Width           =   1035
      End
      Begin VB.Label Label10 
         Caption         =   "Importe"
         Height          =   255
         Left            =   7680
         TabIndex        =   55
         Top             =   270
         Width           =   1035
      End
      Begin VB.Label Label8 
         Caption         =   "Artículos"
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label7 
         Caption         =   "Mano de Obra"
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   270
         Width           =   1035
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Recibo"
      Height          =   255
      Index           =   28
      Left            =   2040
      TabIndex        =   39
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
      TabIndex        =   38
      Top             =   720
      Width           =   510
   End
   Begin VB.Image imgFec 
      Height          =   240
      Index           =   0
      Left            =   3840
      Picture         =   "frmPOZRecibos.frx":012A
      ToolTipText     =   "Buscar fecha"
      Top             =   690
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha "
      Height          =   255
      Index           =   29
      Left            =   3000
      TabIndex        =   37
      Top             =   720
      Width           =   780
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo Recibo"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   36
      Top             =   720
      Width           =   945
   End
   Begin VB.Label Label3 
      Caption         =   "Consumo"
      Height          =   255
      Left            =   1590
      TabIndex        =   35
      Top             =   1380
      Width           =   1035
   End
   Begin VB.Label Label4 
      Caption         =   "Cuota"
      Height          =   255
      Left            =   3000
      TabIndex        =   34
      Top             =   1380
      Width           =   1035
   End
   Begin VB.Label Label5 
      Caption         =   "Hidrante"
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   1380
      Width           =   1035
   End
   Begin VB.Label Label6 
      Caption         =   "Concepto"
      Height          =   255
      Left            =   4200
      TabIndex        =   32
      Top             =   1380
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
Dim i As Integer

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
                    TerminaBloquear
                    PosicionarData
                End If
            End If
        ' *** si n'hi han llínies ***
        
        Case 5 'LLÍNIES
        ' **************************
            If NumTabMto = 1 Then
'                If Not vSeccion Is Nothing Then
'                    vSeccion.CerrarConta
'                    Set vSeccion = Nothing
'                End If
            End If
    
    End Select
    Screen.MousePointer = vbDefault
    

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
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
    
    If Data1.Recordset.Fields(0).Value = "RCP" Then
        cad = "Recibo de Consumo." & vbCrLf
    Else
        cad = "Recibo de Mantenimiento." & vbCrLf
    End If
    
    cad = cad & "-------------------------------------      " & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar el Recibo del Socio:            "
    cad = cad & vbCrLf & "Nº Recibo:  " & Format(Text1(0).Text, "0000000")
    cad = cad & vbCrLf & "Fecha:  " & Format(Text1(1).Text, "dd/mm/yyyy")

    cad = cad & vbCrLf & vbCrLf & " ¿Desea Eliminarlo? "

    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
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
        Frame3.visible = Not FrameConceptos.visible
        Frame3.Enabled = Not FrameConceptos.visible
        Frame4.visible = Not FrameConceptos.visible
        Frame4.Enabled = Not FrameConceptos.visible
        Frame6.visible = Not FrameConceptos.visible
        Frame6.Enabled = Not FrameConceptos.visible
       
        Me.cmdConceptos.Picture = frmPpal.imgListPpal.ListImages(24).Picture
    End If
End Sub

Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String

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
Dim i As Integer
    If Combo1(Index).BackColor = vbYellow Then Combo1(Index).BackColor = vbWhite
    
    i = Combo1(Index).ListIndex
    Text1(20).Text = Mid(Trim(Combo1(Index).List(i)), 1, 3)
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
Dim i As Integer

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
    
'    ' ******* si n'hi han llínies *******
'    'ICONETS DE LES BARRES ALS TABS DE LLÍNIA
'    For kCampo = 0 To ToolAux.Count - 1
'        With Me.ToolAux(kCampo)
'            .HotImageList = frmPpal.imgListComun_OM16
'            .DisabledImageList = frmPpal.imgListComun_BN16
'            .ImageList = frmPpal.imgListComun16
'            .Buttons(1).Image = 3   'Insertar
'            .Buttons(2).Image = 4   'Modificar
'            .Buttons(3).Image = 5   'Borrar
'        End With
'    Next kCampo
'   ' ***********************************
    Me.cmdConceptos.Picture = frmPpal.imgListPpal.ListImages(24).Picture

    Me.FrameConceptos.visible = False
    Me.FrameConceptos.Enabled = False
    Me.cmdConceptos.visible = False
    Me.cmdConceptos.Enabled = False

    LimpiarCampos   'Limpia los campos TextBox
    CargaCombo
    
    CodTipoMov = "RCP"
    VieneDeBuscar = False
    
    If vParamAplic.Cooperativa = 8 Then
        Label28.Caption = "Consumo m3."
        Label13.Caption = "Consumo m3."
        Label14.Caption = "Precio Celador"
    End If
    
    '## A mano
    NombreTabla = "rrecibpozos"
    Ordenacion = " ORDER BY rrecibpozos.numfactu"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    CadenaConsulta = "select * from rrecibpozos "
    If hcoCodMovim <> "" Then
        'Se llama desde Dobleclick en frmManSocios
        CadenaConsulta = CadenaConsulta & " WHERE tipofichero=" & hcoCodTipoM & " AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
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
Dim i As Integer

    On Error Resume Next
    
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Combo1(0).ListIndex = -1
    For i = 0 To Check1.Count - 1
        Me.Check1(i).Value = 0
    Next i
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
        Aux = ValorDevueltoFormGrid(Combo1(0), CadenaDevuelta, 1)
        CadB = CadB & Aux
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 2)
        CadB = CadB & " and " & Aux
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 3)
        CadB = CadB & " and " & Aux
        Aux = ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 4)
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
Dim Sql As String
Dim vSeccion As CSeccion
Dim vSocio As CSocio

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
        
        Case 8, 9, 16, 18 'contadores
            PonerFormatoEntero Text1(Index)
            
        Case 10, 11 'fechas
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
                        
        Case 17, 19 ' precios aplicados
            If Modo <> 1 Then PonerFormatoDecimal Text1(Index), 5
            
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
Dim Tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
    'Llamamos a al form
    '##A mano
    cad = ""
'    Cad = Cad & "Tipo|if(rfactsoc.codtipom='FAA','Anticipo','Liquidación') as a|T||10·"
    cad = cad & ParaGrid(Combo1(0), 0, "Tipo")
    cad = cad & "Tipo Fichero|case rrecibpozos.codtipom when ""RCP"" then ""RCP-Consumo"" when ""RMP"" then ""RMP-Mantenim."" end as tipo|N||22·"
    cad = cad & "Nº.Factura|rrecibpozos.numfactu|N||12·"
    cad = cad & "Fecha|rrecibpozos.fecfactu|F||15·"
    cad = cad & "Código|rrecibpozos.codsocio|N|000000|12·"
    cad = cad & "Socio|rsocios.nomsocio|N||38·"
    
    Tabla = NombreTabla & " inner join rsocios on rrecibpozos.codsocio = rsocios.codsocio "
    Titulo = "Recibos de Contadores"
    devuelve = "0|2|3|4|"
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = Tabla
        frmB.vSQL = CadB
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
    
'    cmdConceptos_Click
    
    cmdConceptos.visible = (Text1(20).Text = "RVP")
    cmdConceptos.Enabled = (Text1(20).Text = "RVP")
  
   
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
Dim i As Byte, Numreg As Byte
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
    
    For i = 0 To Check1.Count - 1
        Me.Check1(i).Enabled = (Modo = 1)
    Next i
    
    b = (Modo <> 1)
    'Campos Nº Recibo bloqueado y en azul
    BloquearTxt Text1(0), b, True
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    BloquearImgBuscar Me, Modo, ModificaLineas
    BloquearImgFec Me, 0, Modo
    
    b = (Modo = 1 Or Modo = 3)
    Text1(1).Enabled = b
    Text1(2).Enabled = b
    imgFec(0).Enabled = b
    imgFec(0).visible = b
    imgBuscar(0).Enabled = b
    imgBuscar(0).visible = b
    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    BloquearTxt Text1(4), (Modo <> 1)
    BloquearTxt Text1(6), (Modo <> 1)
       
    For i = 1 To 2
        imgFec(i).Enabled = (Modo = 1 Or Modo = 3 Or Modo = 4)
    Next i
        
    ' ***************************
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
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
Dim Sql As String

    On Error GoTo EDatosOK

    DatosOk = False
    
'    ComprobarDatosTotales

    'comprobamos datos OK de la tabla scafac
    b = CompForm2(Me, 1)  'Comprobar formato datos ok de la cabecera: opcion=1
    If Not b Then Exit Function
    
    If b Then
        ' comprobamos si insertamos o modificamos que existe el hidrante para el socio
        If (Modo = 3 Or Modo = 4) And Combo1(0).ListIndex = 0 Then
            Sql = ""
            Sql = DevuelveDesdeBDNew(cAgro, "rpozos", "hidrante", "hidrante", Text1(14).Text, "T", , "codsocio", Text1(2).Text, "N")
            
            If Sql = "" Then
                MsgBox "El Hidrante no existe o no es del socio introducido. Revise.", vbExclamation
                PonerFoco Text1(14)
                b = False
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
Dim Sql As String, LEtra As String, Sql2 As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String

    On Error GoTo FinEliminar

    If Data1.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    'Eliminar en tablas de cabecera de factura
    '------------------------------------------
    Sql = " " & ObtenerWhereCP(True)
    
    'Cabecera de factura (rcabfactalmz)
    conn.Execute "Delete from " & NombreTabla & Sql
    
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
Dim Sql As String

    On Error Resume Next
    
    Sql = " codtipom= '" & Text1(20).Text & "'"
    Sql = Sql & " and numfactu = " & Text1(0).Text
    Sql = Sql & " and fecfactu = " & DBSet(Text1(1).Text, "F")
    '08/09/2010 : añadido a la clave primaria
    Sql = Sql & " and numlinea = " & DBSet(Text1(31).Text, "N")

    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function



Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean, bAux As Boolean
Dim i As Integer

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
'++monica: si insertamos lo he quitado
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
'    b = (Modo = 2) And Not Check1(1).Value = 1
'    For I = 0 To ToolAux.Count - 1
'        ToolAux(I).Buttons(1).Enabled = b
'
'        If b Then
'            Select Case I
'              Case 0
'                bAux = (b And Me.Data3.Recordset.RecordCount > 0)
'            End Select
'        End If
'        ToolAux(I).Buttons(2).Enabled = bAux
'        ToolAux(I).Buttons(3).Enabled = bAux
'    Next I


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
    End Select
    
    
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
      
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
    
    cadParam = cadParam & "pDuplicado=1|"
    numParam = numParam + 1
    
    
    If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub
     
    With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 0
            .Titulo = "Impresión de Recibos de Socios"
            .ConSubInforme = True
            .Show vbModal
    End With

    If frmVisReport.EstaImpreso Then
        ActualizarRegistros "rrecibpozos", cadSelect
    End If
End Sub

Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim i As Byte
    
    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 0 To Combo1.Count - 1
        Combo1(i).Clear
    Next i
    
    'tipo de fichero
    Combo1(0).AddItem "RCP-Consumo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "RMP-Mantenimiento"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "RVP-Contadores"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    
    
End Sub


Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String

    On Error GoTo EInsertarCab
    
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        Sql = CadenaInsertarDesdeForm(Me)
        If Sql <> "" Then
            If InsertarOferta(Sql, vTipoMov) Then
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
Dim Sql As String
Dim Numf As String

    On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Factura
    'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
    Do
        devuelve = DevuelveDesdeBDNew(cAgro, NombreTabla, "numfactu", "numfactu", Text1(0).Text, "N", , "codtipom", Text1(20).Text, "T")
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
Dim ImpIVA As Currency
Dim BaseReten As Currency
Dim BaseAFO As Currency
Dim PorRet As Currency
Dim ImpRet As Currency
Dim PorAFO As Currency
Dim ImpAFO As Currency
Dim TotFac As Currency

    Base = CCur(ComprobarCero(Text1(5).Text))
    PorIva = CCur(ComprobarCero(Text1(4).Text))
    ImpIVA = Round2(Base * PorIva / 100, 2)
    
    
    TotFac = Base + ImpIVA

    If ImpIVA = 0 Then
        Text1(6).Text = "0"
    Else
        Text1(6).Text = Format(ImpIVA, "###,##0.00")
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
        Cad1 = "select hidrante, rpozos.codparti, rpartida.nomparti, rpozos.poligono from rpozos where " & cad
        Set Rs = New ADODB.Recordset
        Rs.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            Text1(14).Text = DBLet(Rs.Fields(0).Value)
        End If
    Else
        Set frmMens = New frmMensajes
        frmMens.cadWhere = " and " & cad
        frmMens.campo = Text1(14).Text
        frmMens.OpcionMensaje = 23
        frmMens.Show vbModal
        Set frmMens = Nothing
    End If
    
End Sub

