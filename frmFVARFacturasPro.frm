VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmFVARFacturasPro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturas Varias Proveedor"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12180
   Icon            =   "frmFVARFacturasPro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   12180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Retenci�n"
      ForeColor       =   &H00972E0B&
      Height          =   645
      Left            =   240
      TabIndex        =   75
      Top             =   3750
      Width           =   11850
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   26
         Left            =   855
         MaxLength       =   6
         TabIndex        =   27
         Tag             =   "% Ret|N|S|0|100.00|fvarcabfactpro|retfaccl|##0.00|N|"
         Text            =   "99.99"
         Top             =   225
         Width           =   645
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   27
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   28
         Tag             =   "Cta.Contable|T|S|||fvarcabfactpro|cuereten|||"
         Text            =   "1234567890"
         Top             =   225
         Width           =   1035
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   27
         Left            =   4320
         TabIndex        =   76
         Top             =   225
         Width           =   2895
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   28
         Left            =   8775
         MaxLength       =   15
         TabIndex        =   29
         Tag             =   "Importe Retenci�n|N|S|||fvarcabfactpro|trefaccl|#,###,###,##0.00|N|"
         Top             =   225
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "% Ret."
         Height          =   255
         Index           =   12
         Left            =   345
         TabIndex        =   79
         Top             =   225
         Width           =   615
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   2970
         Tag             =   "-1"
         ToolTipText     =   "Buscar Cta Contable"
         Top             =   225
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cta.Contable"
         Height          =   255
         Index           =   17
         Left            =   2025
         TabIndex        =   78
         Top             =   225
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Retenci�n"
         Height          =   255
         Index           =   18
         Left            =   7335
         TabIndex        =   77
         Top             =   225
         Width           =   1545
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1545
      Index           =   0
      Left            =   240
      TabIndex        =   44
      Top             =   540
      Width           =   11835
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   25
         Left            =   2610
         TabIndex        =   73
         Top             =   1110
         Width           =   3480
      End
      Begin VB.TextBox text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   25
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "Forma de Pago|N|N|||fvarcabfactpro|codforpa|000||"
         Top             =   1110
         Width           =   1035
      End
      Begin VB.TextBox text1 
         Height          =   615
         Index           =   5
         Left            =   6300
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Tag             =   "Observaciones|T|S|||fvarcabfactpro|observac|||"
         Top             =   780
         Width           =   5235
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   1
         Left            =   11225
         TabIndex        =   67
         Tag             =   "Contabilizada|N|N|0|1|fvarcabfactpro|intconta|||"
         Top             =   285
         Width           =   255
      End
      Begin VB.TextBox text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "Cod.Socio|N|S|||fvarcabfactpro|codsocio|000000||"
         Text            =   "123456"
         Top             =   765
         Width           =   1035
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   1
         Left            =   4020
         MaxLength       =   7
         TabIndex        =   2
         Tag             =   "N� de Factura|N|S|0|9999999|fvarcabfactpro|numfactu|0000000|S|"
         Top             =   420
         Width           =   795
      End
      Begin VB.TextBox text1 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   3
         Left            =   120
         MaxLength       =   3
         TabIndex        =   1
         Tag             =   "Seccion|N|N|0|999|fvarcabfactpro|codsecci|000||"
         Top             =   420
         Width           =   900
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   1080
         TabIndex        =   49
         Top             =   420
         Width           =   2070
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   2610
         TabIndex        =   48
         Top             =   765
         Width           =   3480
      End
      Begin VB.TextBox text1 
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   0
         Left            =   3255
         MaxLength       =   3
         TabIndex        =   7
         Tag             =   "Tipo Movimiento|T|N|||fvarcabfactpro|codtipom||S|"
         Top             =   420
         Width           =   705
      End
      Begin VB.TextBox text1 
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   2
         Left            =   4905
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Fecha Factura|F|N|||fvarcabfactpro|fecfactu|dd/mm/yyyy|S|"
         Top             =   420
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "Forma Pago"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   74
         Top             =   1155
         Width           =   945
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1125
         Tag             =   "-1"
         ToolTipText     =   "Buscar Forma de Pago"
         Top             =   1110
         Width           =   240
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   7425
         ToolTipText     =   "Zoom descripci�n"
         Top             =   435
         Width           =   240
      End
      Begin VB.Label Label29 
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   6300
         TabIndex        =   71
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Contabilizada"
         Height          =   255
         Index           =   7
         Left            =   10170
         TabIndex        =   68
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Secci�n"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   66
         Top             =   195
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "N� Factura"
         Height          =   255
         Index           =   4
         Left            =   4020
         TabIndex        =   50
         Top             =   180
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   750
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Buscar Secci�n"
         Top             =   180
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   5805
         Picture         =   "frmFVARFacturasPro.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   135
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1125
         Tag             =   "-1"
         ToolTipText     =   "Buscar Socio"
         Top             =   765
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "TipoMov"
         Height          =   255
         Index           =   2
         Left            =   3180
         TabIndex        =   47
         Top             =   180
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Fact."
         Height          =   255
         Index           =   1
         Left            =   4950
         TabIndex        =   46
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Socio"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   45
         Top             =   810
         Width           =   945
      End
   End
   Begin VB.Frame FrameAux0 
      Caption         =   "Lineas Factura"
      ForeColor       =   &H00972E0B&
      Height          =   2760
      Left            =   225
      TabIndex        =   57
      Top             =   4440
      Width           =   11865
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   10
         Left            =   9180
         MaxLength       =   15
         TabIndex        =   38
         Tag             =   "Precio|N|S|||fvarlinfactpro|precio|###,##0.0000||"
         Text            =   "precio"
         Top             =   1920
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   9
         Left            =   8370
         MaxLength       =   15
         TabIndex        =   37
         Tag             =   "Cantidad|N|S|||fvarlinfactpro|cantidad|###,##0.00||"
         Text            =   "cantidad"
         Top             =   1920
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   10770
         TabIndex        =   72
         Tag             =   "Iva|N|N|0|99|fvarlinfactpro|tipoiva|00||"
         Top             =   1920
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   6
         Left            =   5850
         MaxLength       =   50
         TabIndex        =   36
         Tag             =   "Ampliaci�n|T|S|||fvarlinfactpro|ampliaci|||"
         Text            =   "Ampliacion"
         Top             =   1920
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   0
         Left            =   135
         MaxLength       =   10
         TabIndex        =   40
         Text            =   "Seccion"
         Top             =   1920
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   300
         Index           =   0
         Left            =   3420
         MaskColor       =   &H00000000&
         TabIndex        =   62
         ToolTipText     =   "Buscar Concepto"
         Top             =   1920
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   5
         Left            =   2940
         MaxLength       =   6
         TabIndex        =   35
         Tag             =   "Concepto|N|N|0|999|fvarlinfactpro|codconce|000||"
         Text            =   "Concep"
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   1
         Left            =   1020
         MaxLength       =   3
         TabIndex        =   0
         Tag             =   "Tipo Movimiento|T|N|||fvarlinfactpro|codtipom||S|"
         Text            =   "L"
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   4
         Left            =   2580
         MaxLength       =   2
         TabIndex        =   34
         Tag             =   "N�mero de l�nea|N|N|1|99|fvarlinfactpro|numlinea|00|S|"
         Text            =   "li"
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   2
         Left            =   1380
         MaxLength       =   7
         TabIndex        =   32
         Tag             =   "N� Factura|N|N|0|9999999|fvarlinfactpro|numfactu|0000000|S|"
         Text            =   "Fac"
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   3
         Left            =   1980
         MaxLength       =   10
         TabIndex        =   33
         Tag             =   "Fecha Factura|F|N|||fvarlinfactpro|fecfactu|dd/mm/yyyy|S|"
         Text            =   "fecfactu"
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   7
         Left            =   9900
         MaxLength       =   15
         TabIndex        =   39
         Tag             =   "Importe|N|N|||fvarlinfactpro|importe|##,###,##0.00||"
         Text            =   "Importe"
         Top             =   1920
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3645
         TabIndex        =   58
         Top             =   1935
         Visible         =   0   'False
         Width           =   2115
      End
      Begin MSAdodcLib.Adodc AdoAux 
         Height          =   375
         Index           =   0
         Left            =   4560
         Top             =   240
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
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   0
         Left            =   240
         TabIndex        =   59
         Top             =   270
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
         Begin VB.CheckBox Check2 
            Caption         =   "Vista previa"
            Height          =   195
            Index           =   1
            Left            =   8400
            TabIndex        =   60
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin MSDataGridLib.DataGrid DataGridAux 
         Height          =   1905
         Index           =   0
         Left            =   240
         TabIndex        =   61
         Top             =   735
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   3360
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
   Begin VB.Frame FrameTotFactu 
      Caption         =   "Total Factura"
      ForeColor       =   &H00972E0B&
      Height          =   1575
      Left            =   240
      TabIndex        =   51
      Top             =   2130
      Width           =   11835
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   22
         Left            =   6000
         MaxLength       =   6
         TabIndex        =   24
         Tag             =   "% REC 3|N|S|0|100.00|fvarcabfactpro|porcrec3|##0.00|N|"
         Top             =   1185
         Width           =   645
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   6000
         MaxLength       =   6
         TabIndex        =   18
         Tag             =   "% REC 2|N|S|0|100.00|fvarcabfactpro|porcrec2|##0.00|N|"
         Top             =   840
         Width           =   645
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   6000
         MaxLength       =   6
         TabIndex        =   12
         Tag             =   "% REC 1|N|S|0|100.00|fvarcabfactpro|porcrec1|##0.00|N|"
         Text            =   "99.99"
         Top             =   495
         Width           =   645
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   23
         Left            =   6810
         MaxLength       =   15
         TabIndex        =   25
         Tag             =   "Importe REC 3|N|S|||fvarcabfactpro|imporec3|#,###,###,##0.00|N|"
         Top             =   1185
         Width           =   1635
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   6810
         MaxLength       =   15
         TabIndex        =   19
         Tag             =   "Importe REC 2|N|S|||fvarcabfactpro|imporec2|#,###,###,##0.00|N|"
         Top             =   840
         Width           =   1635
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   6810
         MaxLength       =   15
         TabIndex        =   13
         Tag             =   "Importe Rec 1|N|S|||fvarcabfactpro|imporec1|#,###,###,##0.00|N|"
         Top             =   495
         Width           =   1635
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00CAE3FD&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   24
         Left            =   8790
         MaxLength       =   15
         TabIndex        =   26
         Tag             =   "Total Factura|N|S|||fvarcabfactpro|totalfac|#,###,###,##0.00|N|"
         Top             =   495
         Width           =   2280
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   2430
         MaxLength       =   2
         TabIndex        =   9
         Tag             =   "Tipo IVA 1|N|S|0|99|fvarcabfactpro|tipoiva1|00||"
         Text            =   "12"
         Top             =   510
         Width           =   525
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   2445
         MaxLength       =   2
         TabIndex        =   15
         Tag             =   "Tipo IVA 2|N|S|0|99|fvarcabfactpro|tipoiva2|00||"
         Top             =   840
         Width           =   525
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   2445
         MaxLength       =   2
         TabIndex        =   21
         Tag             =   "Tipo IVA 3|N|S|0|99|fvarcabfactpro|tipoiva3|00||"
         Top             =   1185
         Width           =   525
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   3180
         MaxLength       =   6
         TabIndex        =   10
         Tag             =   "% IVA 1|N|S|0|100.00|fvarcabfactpro|porciva1|##0.00|N|"
         Text            =   "99.99"
         Top             =   510
         Width           =   645
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   3180
         MaxLength       =   6
         TabIndex        =   16
         Tag             =   "% IVA 2|N|S|0|100.00|fvarcabfactpro|porciva2|##0.00|N|"
         Top             =   840
         Width           =   645
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   20
         Left            =   3180
         MaxLength       =   6
         TabIndex        =   22
         Tag             =   "% IVA 3|N|S|0|100.00|fvarcabfactpro|porciva3|##0.00|N|"
         Top             =   1185
         Width           =   645
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   4005
         MaxLength       =   15
         TabIndex        =   11
         Tag             =   "Importe IVA 1|N|S|||fvarcabfactpro|impoiva1|#,###,###,##0.00|N|"
         Top             =   495
         Width           =   1605
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   4005
         MaxLength       =   15
         TabIndex        =   17
         Tag             =   "Importe IVA 2|N|S|||fvarcabfactpro|impoiva2|#,###,###,##0.00|N|"
         Top             =   840
         Width           =   1605
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   21
         Left            =   4005
         MaxLength       =   15
         TabIndex        =   23
         Tag             =   "Importe IVA 3|N|S|||fvarcabfactpro|impoiva3|#,###,###,##0.00|N|"
         Top             =   1185
         Width           =   1605
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   240
         MaxLength       =   15
         TabIndex        =   8
         Tag             =   "Base IVA 1|N|S|||fvarcabfactpro|baseiva1|#,###,###,##0.00|N|"
         Text            =   "575757575757557"
         Top             =   495
         Width           =   1605
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   240
         MaxLength       =   15
         TabIndex        =   14
         Tag             =   "Base IVA 2|N|S|||fvarcabfactpro|baseiva2|#,###,###,##0.00|N|"
         Top             =   840
         Width           =   1605
      End
      Begin VB.TextBox text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   240
         MaxLength       =   15
         TabIndex        =   20
         Tag             =   "Base IVA 3|N|S|||fvarcabfactpro|baseiva3|#,###,###,##0.00|N|"
         Top             =   1185
         Width           =   1605
      End
      Begin VB.Label Label1 
         Caption         =   "% Rec."
         Height          =   255
         Index           =   8
         Left            =   6030
         TabIndex        =   70
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Recargo"
         Height          =   255
         Index           =   0
         Left            =   6810
         TabIndex        =   69
         Top             =   270
         Width           =   1545
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   2145
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Buscar tipo de IVA"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   2145
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Buscar tipo de IVA"
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   2145
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Buscar tipo de IVA"
         Top             =   510
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Total Factura"
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
         Index           =   11
         Left            =   8790
         TabIndex        =   56
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo IVA"
         Height          =   255
         Index           =   14
         Left            =   2445
         TabIndex        =   55
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "% IVA"
         Height          =   255
         Index           =   15
         Left            =   3210
         TabIndex        =   54
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Importe IVA"
         Height          =   255
         Index           =   16
         Left            =   4005
         TabIndex        =   53
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   52
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   42
      Top             =   7200
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
         TabIndex        =   43
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   11010
      TabIndex        =   31
      Top             =   7350
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9750
      TabIndex        =   30
      Top             =   7350
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   4200
      Top             =   7320
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
      Height          =   375
      Left            =   11010
      TabIndex        =   41
      Top             =   7350
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   64
      Top             =   0
      Width           =   12180
      _ExtentX        =   21484
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar Total Factura"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "�ltimo"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Index           =   0
         Left            =   8520
         TabIndex        =   65
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Empresa"
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   63
      Top             =   720
      Width           =   615
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
      Begin VB.Menu mn_ModTotales 
         Caption         =   "&Mod.Totales"
         Enabled         =   0   'False
         Shortcut        =   ^T
         Visible         =   0   'False
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
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
Attribute VB_Name = "frmFVARFacturasPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MONICA  +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public DatosADevolverBusqueda As String    'Tindr� el n� de text que vol que torne, empipat
Public Event DatoSeleccionado(CadenaSeleccion As String)


Public numfactu As Long
Public LetraSerie As String
Public Tipo As Byte ' 0 schfac normal
                    ' 1 schfacr ajena para el Regaixo

Private Modo As Byte
'*************** MODOS ********************
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la b�squeda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edici� del camp
'   3.-  Inserci� de nou registre
'   4.-  Modificar
'   5.-  Manteniment Llinies
'   6.-  Modificar totales
'***Variables comuns a tots els formularis*****

Dim ModoLineas As Byte
'1.- A�adir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim NumTabMto As Integer 'Indica quin n� de Tab est� en modo Mantenimient
Dim TituloLinea As String 'Descripci� de la ll�nia que est� en Mantenimient
Dim PrimeraVez As Boolean
Dim ModificarTotales As Boolean

Private CadenaConsulta As String 'SQL de la taula principal del formulari
Private Ordenacion As String
Private NomTabla As String  'Nom de la taula

Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnPrimero As Byte 'Variable que indica el n� del Bot� PrimerRegistro en la Toolbar1
'Dim CadAncho() As Boolean  'array, per a quan cridem al form de ll�nies
Dim Indice As Integer 'Indice del text1 donde se ponen los datos devueltos desde otros Formularios de Mtos
Dim indCodigo As Byte

Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1

Private WithEvents frmSec As frmManSeccion
Attribute frmSec.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmFpa As frmComFpa 'formas de pago de comercial
Attribute frmFpa.VB_VarHelpID = -1

Private WithEvents frmCon As frmFVARConceptos
Attribute frmCon.VB_VarHelpID = -1
Private WithEvents frmCtas As frmCtasConta
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmTipIVA As frmTipIVAConta  'Tipos de IVA de la contabilidad
Attribute frmTipIVA.VB_VarHelpID = -1

Dim CtaAnt As String
Dim FormaPagoAnt As String
Dim ModoModificar As Boolean
Dim ModificaImportes As Boolean ' variable que me indica q hay que modificar lineas de la factura de contabilidad
                                ' y cobros en la tesoreria
Dim BdConta As Integer
Dim BdConta1 As Integer

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadparam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim TipForpa As String
Dim TipForpaAnt As String

' utilizado para buscar por checks
Private BuscaChekc As String

Dim CadenaBorrado As String

Dim Seguir As Boolean
Dim vSeccion As CSeccion
Dim CodTipoMov As String


Private Sub btnBuscar_Click(Index As Integer)
    ' els formularis als que crida son d'una atra BDA
    TerminaBloquear
    
    Select Case Index
        Case 0 'Conceptos
            Set frmCon = New frmFVARConceptos
            frmCon.DatosADevolverBusqueda = "0|1|2|4|"
            frmCon.CodigoActual = txtAux(5).Text
            frmCon.Show vbModal
            Set frmCon = Nothing
            
    End Select
    
    PonerFoco txtAux(5)
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

Private Sub Check1_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "check1(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "check1(" & Index & ")|"
    End If
End Sub


Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cmdAceptar_Click()
Dim b As Boolean
Dim vSec As CSeccion 'Clase Seccion
Dim vTabla As String
Dim CtaClie As String
Dim Cad As String

' variables para el recalculo de iva y totales
    Dim I As Integer
    Dim Imptot(2)
    Dim Tipiva(2)
    Dim Impbas(2) As Currency
    Dim ImpIVA(2) As Currency
    Dim PorIva(2) As Currency
    Dim ImpREC(2) As Currency
    Dim PorRec(2) As Currency
    Dim TotFac As Currency

'retencion
    Dim PorRet As Currency
    Dim ImpRet As Currency
    
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    ModoModificar = False
    b = True
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then InsertarCabecera

        Case 4  'MODIFICAR
            If Not DatosOk Then
                ModoLineas = 0
                Screen.MousePointer = vbDefault
                Exit Sub
            Else
                ModoModificar = True
                conn.BeginTrans
                
                Set vSeccion = New CSeccion
                If vSeccion.LeerDatos(text1(3)) Then
                    Text2(3).Text = vSeccion.Nombre
                    If vSeccion.AbrirConta Then
                        PorRet = 0
                        If text1(26).Text <> "" Then PorRet = CCur(ImporteSinFormato(text1(26).Text))
                        AdoAux(0).Recordset.MoveFirst
                        RecalculoBasesIvaFactura AdoAux(0).Recordset, Imptot, Tipiva, Impbas, ImpIVA, PorIva, TotFac, ImpREC, PorRec, PorRet, ImpRet, text1(4).Text, text1(0).Text

                        text1(28).Text = ""
                        If ImpRet <> 0 Then text1(28).Text = Format(ImpRet, "#,###,###,##0.00")
                        text1(24).Text = Format(TotFac, "#,###,###,##0.00")

                        If text1(8).Text = "" Then text1(8).Text = "0,00"
                        If text1(9).Text = "" Then text1(9).Text = "0,00"
                    End If
                End If
                
                If CadenaBorrado <> "" Then
                    conn.Execute CadenaBorrado
                    CadenaBorrado = ""
                    EliminarLinea
                End If
                
                If ModificaDesdeFormulario2(Me, 1) Then
                    If Check1(1).Value = 1 Then
                        MsgBox "Los cambios realizados recuerde hacerlos en la Contabilidad y Cartera correspondiente.", vbExclamation
                        
'12/02/2008: lo he quitado porque los cambios los haran ellos en la contabilidad
'                        'solo en el caso de que este contabilizada
'                        If Val(CtaAnt) <> Val(text1(4).Text) Then
''                            CtaClie = ""
''                            CtaClie = DevuelveDesdeBDNew(cPTours, "ssocio", "codmacta", "codsocio", text1(3).Text, "N")
'                            b = ModificaCtaClienteFacturaContabilidad(text1(0).Text, text1(1).Text, text1(2).Text, text1(4).Text)
'                        End If
'' 09022007 ya no dejo modificar la forma de pago
''                        If Val(FormaPagoAnt) <> Val(Text1(5).Text) Then _
''                            ModificaFormaPagoTesoreria Text1(0).Text, Text1(1).Text, Text1(2).Text, Text1(5).Text, FormaPagoAnt, TipForpa, TipForpaAnt
'
'                        If ModificaImportes And b Then
'                            BorrarTMPErrFact
'                            vTabla = "fvarcabfactpro"
'' cuando aclare temas de contabilizacion en tesoreria se tiene que realizar esta funcion
''                            b = ModificaImportesFacturaContabilidad(text1(0).Text, text1(1).Text, text1(2).Text, text1(18).Text, text1(5).Text, vTabla)
'                            ModificaImportes = False
'                        End If
                    End If
                    TerminaBloquear
                    PosicionarData "codtipom = '" & Trim(text1(0).Text) & "' and numfactu = " & text1(1).Text & " and fecfactu = " & DBSet(text1(2).Text, "F")
                End If
            
                If Not vSeccion Is Nothing Then
                    vSeccion.CerrarConta
                    Set vSeccion = Nothing
                End If
            End If
        
        Case 5 'LLINIES
            Select Case ModoLineas
                Case 1 'afegir llinia
                    InsertarLinea
                Case 2 'modificar llinies
                    ModificarLinea
                    PosicionarData "codtipom = '" & Trim(text1(0).Text) & "' and numfactu = " & text1(1).Text & " and fecfactu = " & DBSet(text1(2).Text, "F")
                    Screen.MousePointer = vbDefault
                    Exit Sub
            End Select
            
            
        Case 6  'MODIFICAR TOTALES
            If Not DatosOk Then
                ModoLineas = 0
                Screen.MousePointer = vbDefault
                Exit Sub
            Else
                ModoModificar = True
                conn.BeginTrans
                
                If ModificaDesdeFormulario2(Me, 1) Then
                    If Check1(1).Value = 1 Then
                        MsgBox "Los cambios realizados recuerde hacerlos en la Contabilidad y Cartera correspondiente.", vbExclamation
                        
                    End If
                    TerminaBloquear
                    PosicionarData "codtipom = '" & Trim(text1(0).Text) & "' and numfactu = " & text1(1).Text & " and fecfactu = " & DBSet(text1(2).Text, "F")
                End If
            End If
            
            
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Or Not b Then
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
        If ModoModificar Then
            conn.RollbackTrans
'            ConnContaFac.RollbackTrans
            ModoModificar = False
        End If
    Else
        If ModoModificar Then
            conn.CommitTrans
'            ConnContaFac.CommitTrans
            ModoModificar = False
        End If
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
Dim Sql2 As String

    PrimeraVez = True

    ' ICONITOS DE LA BARRA
    btnPrimero = 16 'index del bot� "primero"
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'el 1 i el 2 son separadors
        .Buttons(3).Image = 1   'Buscar
        .Buttons(4).Image = 2   'Todos
        'el 5 i el 6 son separadors
        .Buttons(7).Image = 3   'Insertar
        .Buttons(8).Image = 4   'Modificar
        .Buttons(9).Image = 5   'Borrar
        .Buttons(10).Image = 13 ' Modificar totales
        .Buttons(12).Image = 10  'Imprimir
        .Buttons(13).Image = 11  'Salir
        'el 14 i el 15 son separadors
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 '�ltimo
    End With
    
    'ICONITOS DE LAS BARRAS EN LOS TABS DE LINEA
    For I = 0 To ToolAux.Count - 1
        With Me.ToolAux(I)
            '.ImageList = frmPpal.imgListComun_VELL
            '  ### [Monica] 02/10/2006 acabo de comentarlo
            '.HotImageList = frmPpal.imgListComun_OM16
            '.DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next I
    
    'cargar IMAGES de busqueda
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
   
    'IMAGES para zoom
    For I = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(I).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next I
    
    LimpiarCampos   'Limpia los campos TextBox
    For I = 0 To DataGridAux.Count - 1 'neteje tots els grids de llinies
        DataGridAux(I).ClearFields
    Next I
    
    '## A mano
    NomTabla = "fvarcabfactpro"
    Ordenacion = " ORDER BY codtipom, numfactu, fecfactu "
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Sql2 = "Select * from " & NomTabla & " where numfactu = -1"
    Data1.RecordSource = Sql2
    Data1.Refresh
        
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
        text1(0).BackColor = vbYellow 'letraser
    End If
    
    ModoLineas = 0
    
    For I = 0 To DataGridAux.Count - 1
        CargaGrid I, (Modo = 2) 'carregue els datagrids de llinies
    Next I
    
    If LetraSerie <> "" Then
        text1(0).Text = Trim(LetraSerie)
        text1(1).Text = numfactu
        PonerModo 1
        cmdAceptar_Click
    End If

    CodTipoMov = "FVP"



End Sub

Private Sub LimpiarCampos()
    On Error Resume Next

    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    
    Me.Check1(1).Value = 0
    
    'Aqui va el especifico de cada form es
    '### a mano
    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim I As Integer, NumReg As Byte
Dim b As Boolean
Dim b1 As Boolean
On Error GoTo EPonerModo
 
    Modo = Kmodo
    BuscaChekc = ""
    
    PonerIndicador lblIndicador, Modo, ModoLineas
       
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    
    '=======================================
    b = (Modo = 2)
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    

    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    '---------------------------------------------
    
    'Bloquea los campos Text1 si no estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    BloquearText1 Me, Modo
    BloquearChecks Me, Modo
    
    BloquearImgBuscar Me, Modo, ModoLineas
       
    'Bloquear los campos de clave primaria, NO se puede modificar
    b = Not (Modo = 1) 'solo al insertar/buscar estar� activo
    For I = 0 To 1
        BloquearTxt text1(I), b, True
        text1(I).Enabled = Not b
    Next I
    b = (Modo = 4) Or (Modo = 0) Or (Modo = 2) Or (Modo = 5)
    For I = 2 To 3
        BloquearTxt text1(I), b, True
        text1(I).Enabled = Not b
    Next I
    
    
    For I = 6 To 24
        BloquearTxt text1(I), Not (Modo = 1 Or (Modo = 4 And ModificarTotales))
    Next I
    
    ' el importe de retencion solo se puede consultar
    BloquearTxt text1(28), Not (Modo = 1 Or (Modo = 4 And ModificarTotales))
    text1(28).Enabled = (Modo = 1 Or (Modo = 4 And ModificarTotales))
    
'    'Los % de IVA siempre bloqueados
'    BloquearTxt text1(8), True
'    BloquearTxt text1(14), True
'    BloquearTxt text1(20), True
'    'Los % de REC siempre bloqueados
'    BloquearTxt text1(10), True
'    BloquearTxt text1(16), True
'    BloquearTxt text1(22), True
    'El total de la factura siempre bloqueado
'    BloquearTxt text1(24), True
    
    '09/02/2007 no dejo modificar la forma de pago
    b = ((Modo = 4) And Me.Check1(1).Value = 1) Or (Modo = 0) Or (Modo = 2) Or (Modo = 5)
    BloquearTxt text1(25), b
    
    text1(24).BackColor = &HCAE3FD
    
    Me.FrameTotFactu.Enabled = (Modo = 1)
    
    ' **** si n'hi han imagens de buscar en la cap�alera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
    ' ********************************************************

    b = (Modo = 3) Or (Modo = 1)
    Me.imgBuscar(0).Enabled = b
    Me.imgBuscar(0).visible = b
    
    b = (Modo = 3) Or (Modo = 1) Or (Modo = 4 And Me.Check1(1).Value = 0)
    Me.imgBuscar(5).Enabled = b
    Me.imgBuscar(5).visible = b
    
    ' ayuda de socio
    imgBuscar(1).Enabled = (Modo = 1) Or (Modo = 3) Or (Modo = 4)
    imgBuscar(1).visible = (Modo = 1) Or (Modo = 3) Or (Modo = 4)
    
    'Imagen Calendario fechas
    b = (Modo = 3 Or Modo = 4 Or Modo = 1 Or (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2)))
    Me.imgFec(2).Enabled = (Modo = 3 Or Modo = 1) 'es clave, solo al insertar o buscar
    Me.imgFec(2).visible = (Modo = 3 Or Modo = 1) 'es clave, solo al insertar o buscar
            
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor.
    PonerLongCampos
                          
    If (Modo < 2) Or (Modo = 3) Then
        For I = 0 To DataGridAux.Count - 1
            CargaGrid I, False
        Next I
    End If
    
    b = (Modo = 4) Or (Modo = 2)
    For I = 0 To DataGridAux.Count - 1
        DataGridAux(I).Enabled = b
    Next I
    
    ' solo podremos tocar el campo de contabilizado si estamos buscando
    Check1(1).Enabled = (Modo = 1)
    
    'b = (Modo = 4)
    b = (Modo = 1) Or (Modo = 4 And ModificarTotales)
    FrameTotFactu.Enabled = b
    
    Frame2(0).Enabled = (Modo = 4 And Not ModificarTotales) Or (Modo <> 4)
    
    b = (Modo = 5)
    Me.FrameAux0.Enabled = (Modo = 2) Or (Modo = 5)
    
    PonerModoOpcionesMenu (Modo) 'Activar opciones menu seg�n modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario

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
'Activas unas Opciones de Menu y Toolbar seg�n el nivel de usuario
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub

Private Sub PonerModoOpcionesMenu(Modo)
'Activas unas Opciones de Menu y Toolbar seg�n el modo en que estemos
Dim b As Boolean, bAux As Boolean
Dim I As Byte

    '-----  TOOLBAR DE LA CABECERA
    '------------------------------------------
    'b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    b = (Modo = 2 Or Modo = 0)
    'Buscar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(4).Enabled = b
    Me.mnVerTodos.Enabled = b
    'Insertar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnNuevo.Enabled = b
    
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And (Check1(1).Value = 0)
    'Modificar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(9).Enabled = b
    Me.mnEliminar.Enabled = b
    'modificar totals
    Toolbar1.Buttons(10).Enabled = b
    Me.mnEliminar.Enabled = b
    
    'Imprimir
    'VRS:2.0.1(3)
    Toolbar1.Buttons(12).Enabled = (Modo = 2)
    Me.mnImprimir.Enabled = (Modo = 2)
    
    '-----------  LINEAS
    ' *** MEU: botons de les ll�nies de cuentas bancarias,
    ' nom�s es poden gastar quan inserte o modifique clients ****
    'b = (Modo = 3 Or Modo = 4)
    b = (Modo = 3 Or (Modo = 4 And Not ModificarTotales) Or Modo = 2) 'And (Check1(1).Value = 0)
    For I = 0 To ToolAux.Count - 1
        ToolAux(I).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.AdoAux(I).Recordset.RecordCount > 0)
        ToolAux(I).Buttons(2).Enabled = bAux
        ToolAux(I).Buttons(3).Enabled = bAux
    Next I
    'Imprimir en pesta�a Comisiones de Productos
'    ToolAux(2).Buttons(6).Enabled = (Modo = 2) Or (Modo = 3) Or (Modo = 4) Or (Modo = 5 And ModoLineas = 0)
    ' ************************************************************
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index
    PonerCampos
End Sub

Private Function MontaSQLCarga(Index As Integer, enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Bas�ndose en la informaci�n proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim Sql As String
Dim Tabla As String
    
    Select Case Index
        Case 0 'Lineas de factura
                Tabla = "fvarlinfactpro"
                Sql = "SELECT codtipom,numfactu,fecfactu,numlinea,fvarlinfactpro.codconce,fvarconce.nomconce, fvarlinfactpro.tipoiva, ampliaci,"
                Sql = Sql & "cantidad, precio, importe"
                Sql = Sql & " FROM fvarlinfactpro, fvarconce "
                Sql = Sql & " WHERE fvarlinfactpro.codconce = fvarconce.codconce "
    
                If enlaza Then
                    Sql = Sql & " AND " & Replace(ObtenerWhereCab(False), "fvarcabfactpro", "fvarlinfactpro")
                Else
                    Sql = Sql & " AND fvarlinfactpro.codtipom is null"
                End If
                Sql = Sql & " ORDER BY " & Tabla & ".numlinea "
    End Select
    MontaSQLCarga = Sql
End Function

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim CadB As String
    Dim Aux As String
    
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(text1(3), CadenaDevuelta, 1) 'codsecci
        CadB = Aux
        Aux = ValorDevueltoFormGrid(text1(0), CadenaDevuelta, 2) 'letraser
        CadB = CadB & " AND " & Aux
        Aux = ValorDevueltoFormGrid(text1(1), CadenaDevuelta, 3) 'numfactu
        CadB = CadB & " AND " & Aux
        Aux = ValorDevueltoFormGrid(text1(2), CadenaDevuelta, 4) 'fecfactu
        CadB = CadB & " AND " & Aux
        '   Como la clave principal es unica, con poner el sql apuntando
        '   al valor devuelto sobre la clave ppal es suficiente
        CadenaConsulta = "select * from " & NomTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
    'Fecha
    text1(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
    text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nommacta
End Sub

Private Sub frmFPa_DatoSeleccionado(CadenaSeleccion As String)
Dim Cad As String
    text1(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codforpa
    Text2(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nomsecci
End Sub

Private Sub frmCon_DatoSeleccionado(CadenaSeleccion As String)
'Conceptos
Dim BdConta As String
Dim Tipiva As String

    txtAux(5).Text = RecuperaValor(CadenaSeleccion, 1) 'codconce
    txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'nomartic
    BdConta = RecuperaValor(CadenaSeleccion, 3) 'base de datos de conta
    Tipiva = RecuperaValor(CadenaSeleccion, 4) 'tipo de iva
End Sub

Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
Dim Cad As String
    text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codsecci
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'nomsecci
    
    Cad = RecuperaValor(CadenaSeleccion, 3)  'numconta
    If Cad <> "" Then BdConta = CInt(Cad)  'numero de conta
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codsocio
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nomsocio
End Sub

Private Sub frmTipIVA_DatoSeleccionado(CadenaSeleccion As String)
'Tipos de IVA (de la Contabilidad)
    text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigiva
    FormateaCampo text1(Indice)
    text1(Indice + 1).Text = RecuperaValor(CadenaSeleccion, 3) '% iva
    If Modo <> 1 Then
        text1(Indice + 3).Text = RecuperaValor(CadenaSeleccion, 4) '% rec
    End If
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     text1(Indice).Text = vCampo
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   'Screen.MousePointer = vbHourglass
    TerminaBloquear
    
    Select Case Index
        Case 0 'Seccion
            Indice = 3
            Set frmSec = New frmManSeccion
            frmSec.DatosADevolverBusqueda = "0|1|2|3|4|"
            frmSec.CodigoActual = text1(3).Text
            frmSec.Show vbModal
            Set frmSec = Nothing
            
        Case 1 'codigo de socio
            Indice = 4
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
'            frmSoc.CodigoActual = Text1(4).Text
            frmSoc.Show vbModal
            Set frmSoc = Nothing
        
        Case 6 'Cuenta Contable
            If text1(3).Text = "" Then Exit Sub
            
            Indice = 27
            
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(text1(3)) Then
                If vSeccion.AbrirConta Then
                    Set frmCtas = New frmCtasConta
                    
                    frmCtas.NumDigit = 0
                    frmCtas.DatosADevolverBusqueda = "0|1|"
                    frmCtas.CodigoActual = text1(Indice).Text
                    frmCtas.Show vbModal
                    Set frmCtas = Nothing
                
                    Set frmCtas = Nothing
                End If
            End If
            If Not vSeccion Is Nothing Then
                vSeccion.CerrarConta
                Set vSeccion = Nothing
            End If
                        
        Case 5 'forma de pago
            Indice = Index + 20
            AbrirFrmForpa Indice
            
        Case 2, 3, 4 'tipos de IVA (de la contabilidad)
            If text1(3).Text = "" Then Exit Sub
            
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(text1(3)) Then
                If vSeccion.AbrirConta Then
                    If Index = 2 Then Let Indice = 7
                    If Index = 3 Then Let Indice = 13
                    If Index = 4 Then Let Indice = 19
                    
                    Set frmTipIVA = New frmTipIVAConta
                    
                    frmTipIVA.DatosADevolverBusqueda = "0|1|"
                    frmTipIVA.CodigoActual = text1(Indice).Text
                    frmTipIVA.Show vbModal
                    
                    Set frmTipIVA = Nothing
                End If
            End If
            If Not vSeccion Is Nothing Then
                vSeccion.CerrarConta
                Set vSeccion = Nothing
            End If

    End Select
    
    PonerFoco text1(Indice)
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
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

    ' es desplega baix i cap a la dreta
    'frmC.Left = esq + imgFec(Index).Parent.Left + 30
    'frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40
    
    ' es desplega dalt i cap a la esquerra
    frmC.Left = esq + imgFec(Index).Parent.Left - frmC.Width + imgFec(Index).Width + 40
    frmC.Top = dalt + imgFec(Index).Parent.Top - frmC.Height + menu - 25
       
    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
    imgFec(2).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If text1(Index).Text <> "" Then frmC.NovaData = text1(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco text1(CByte(imgFec(2).Tag))
    ' ***************************
End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        Indice = 5
        frmZ.pTitulo = "Observaciones de la Factura"
        frmZ.pValor = text1(Indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco text1(Indice)
    End If
End Sub

Private Sub mn_ModTotales_Click()

    'Comprobaciones
    '--------------
    If Data1.Recordset.EOF Then Exit Sub
    If Data1.Recordset.RecordCount < 1 Then Exit Sub
    
    'El registre de codi 0 no es pot Modificar ni Eliminar
    ' ### [Monica] 27/09/2006
    ' quitamos el control de no poder modificar ni eliminar si es 0
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    
    ' ### [Monica] 27/09/2006
    ' solo podemos modificar en el caso de que haya contabilidad si la factura es modificable
'12/02/2008: lo he quitado porque lo modificaran ellos manualmente en la contabilidad
'    If vParamAplic.NumeroConta <> 0 And Not FacturaModificable(text1(0).Text, text1(1).Text, text1(2).Text, Check1(1).Value) Then Exit Sub
    
    'Preparar para modificar
    '-----------------------
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificarTotales
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
    Me.Check1(1).Value = 0
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    'VRS:2.0.1(3): a�adido el boton de imprimir
    cadTitulo = "Impresi�n Facturas Varias Proveedor"

    ' ### [Monica] 11/09/2006
    '****************************
    Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
    Dim nomDocu As String 'Nombre de Informe rpt de crystal

    indRPT = 91 'Facturas Varias Proveedor


    cadparam = "|pEmpresa=" & vEmpresa.nomempre & "|" '& "|pCodigoISO="11112"|pCodigoRev="01"|

    If Not PonerParamRPT(indRPT, cadparam, numParam, nomDocu) Then Exit Sub
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
    ' he a�adido estas dos lineas para que llame al rpt correspondiente

    cadNombreRPT = nomDocu  ' "rFactgas.rpt"
    cadFormula = "({" & NomTabla & ".codtipom} = """ & Trim(text1(0).Text) & """) AND ({" & NomTabla & ".numfactu} = " & text1(1).Text & ") and ({" & NomTabla & ".fecfactu} = cdate(""" & text1(2).Text & """)) "
    
    '23022007 Monica: la separacion de la bonificacion solo la quieren en Alzira
'    If vParamAplic.Cooperativa = 1 Then cadFormula = cadFormula & " and {slhfac.numalbar} <> 'BONIFICA'" ' AND ({ssocio.impfactu}<=1)"
    
    
    LlamarImprimir
End Sub

Private Sub mnModificar_Click()

    'Comprobaciones
    '--------------
    If Data1.Recordset.EOF Then Exit Sub
    If Data1.Recordset.RecordCount < 1 Then Exit Sub
    
    'El registre de codi 0 no es pot Modificar ni Eliminar
    ' ### [Monica] 27/09/2006
    ' quitamos el control de no poder modificar ni eliminar si es 0
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    
    ' ### [Monica] 27/09/2006
    ' solo podemos modificar en el caso de que haya contabilidad si la factura es modificable
'12/02/2008: lo he quitado porque lo modificaran ellos manualmente en la contabilidad
'    If vParamAplic.NumeroConta <> 0 And Not FacturaModificable(text1(0).Text, text1(1).Text, text1(2).Text, Check1(1).Value) Then Exit Sub
    
    'Preparar para modificar
    '-----------------------
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
End Sub



Private Sub mnNuevo_Click()
     BotonAnyadir
End Sub

'Private Sub mnRectificar_Click()
'
'    'Comprobaciones
'    '--------------
'    If Data1.Recordset.EOF Then Exit Sub
'    If Data1.Recordset.RecordCount < 1 Then Exit Sub
'
'    'El registre de codi 0 no es pot Modificar ni Eliminar
'    ' ### [Monica] 27/09/2006
'    ' quitamos el control de no poder modificar ni eliminar si es 0
'    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
'
'    ' ### [Monica] 27/09/2006
'    ' solo podemos modificar en el caso de que haya contabilidad si la factura es modificable
'    If vParamAplic.NumeroConta <> 0 And Not FacturaModificable(Text1(0).Text, Text1(1).Text, Text1(2).Text, Check1(1).Value) Then Exit Sub
'
'    'Preparar para modificar
'    '-----------------------
'    If Not FacturaModificable(Text1(0).Text, Text1(1).Text, Text1(2).Text, Check1(1).Value) Then Exit Sub
'    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonRectificar
'End Sub



Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Cad As String
    
    
    Select Case Button.Index
        Case 3  'Buscar
           mnBuscar_Click
        Case 4  'Todos
            mnVerTodos_Click
        Case 7  'Nuevo
            mnNuevo_Click
        Case 8  'Modificar
            '++monica:12/02/2008
            If CByte(Data1.Recordset!intconta) = 1 Then
               Cad = "   Se dispone a realizar cambios en los datos de la Factura.     " & vbCrLf & vbCrLf & _
                     "Recuerde modificar la Contabilidad y Tesoreria correspondiente!!!"
               MsgBox Cad, vbExclamation
            End If
            '++
            mnModificar_Click
        Case 9  'Borrar
            '++monica:12/02/2008
            If CByte(Data1.Recordset!intconta) = 1 Then
               Cad = "No se permite eliminar una Factura Contabilizada!!!"
               MsgBox Cad, vbExclamation
            Else
            '++
                mnEliminar_Click
            End If
        Case 10 'Rectificativa
            mn_ModTotales_Click
        Case 12 'Imprimir
            mnImprimir_Click
        Case 13    'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonBuscar()
    'Buscar
    Seguir = True
    
    If Modo <> 1 Then
        BdConta = 0
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        'LimpiarDataGrids
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco text1(3)
        text1(3).BackColor = vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            text1(kCampo).Text = ""
            text1(kCampo).BackColor = vbYellow
            PonerFoco text1(kCampo)
        End If
    End If
End Sub

Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
    If chkVistaPrevia(0) = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NomTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        PonerFoco text1(0)
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
Dim Cad As String
        'Llamamos a al form
        '##A mano
        Cad = ""
        Cad = Cad & "Secci�n|" & NomTabla & ".codsecci|N|" & FormatoCampo(text1(3)) & "|10�"
        Cad = Cad & "Nom. Secci�n|nomsecci|T||28�"
        Cad = Cad & "Tipo|" & NomTabla & ".codtipom|T|" & text1(0) & "|6�"
        Cad = Cad & "N� Fact.|" & NomTabla & ".numfactu|N|" & FormatoCampo(text1(1)) & "|10�"
        Cad = Cad & ParaGrid(text1(2), 14, "Fecha")
        Cad = Cad & "Codigo|" & NomTabla & ".codsocio|N|" & FormatoCampo(text1(4)) & "|9�"
        Cad = Cad & "Socio|rsocios.nomsocio|N||23�"
        
        If Cad <> "" Then
            Screen.MousePointer = vbHourglass
            Set frmB = New frmBuscaGrid
            frmB.vCampos = Cad
            frmB.vTabla = "(" & NomTabla & " INNER JOIN rseccion ON " & NomTabla & ".codsecci=rseccion.codsecci) INNER JOIN rsocios ON " & NomTabla & ".codsocio = rsocios.codsocio"
            frmB.vSQL = CadB
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|2|3|4|"
            frmB.vTitulo = "Facturas Varias Proveedor"
            frmB.vSelElem = 0
            '#
            frmB.Show vbModal
            Set frmB = Nothing
            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
            'tendremos que cerrar el form lanzando el evento
            If HaDevueltoDatos Then
                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                    cmdRegresar_Click
            Else   'de ha devuelto datos, es decir NO ha devuelto datos
                PonerFoco text1(kCampo)
            End If
        End If
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String
Dim Aux As String
Dim I As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ning�n registro devuelto.", vbExclamation
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
            Cad = Cad & text1(J).Text & "|"
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
        MsgBox "No hay ning�n registro en la tabla " & NomTabla, vbInformation
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
'Ver todos
Dim I As Integer

    LimpiarCampos 'Limpia los Text1
    
    For I = 0 To DataGridAux.Count - 1 'Limpias los DataGrid
        CargaGrid I, False
    Next I
    
    If chkVistaPrevia(0).Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NomTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub BotonAnyadir()
'A�adir registro en tabla de expedientes individuales: expincab (Cabecera)

    LimpiarCampos 'Vac�a los TextBox
    'Poner los grid sin apuntar a nada
'    LimpiarDataGrids


    Seguir = True
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    
    text1(0).Text = "FVP"
    'Quan afegixc pose en Fecha
    text1(2).Text = Format(Now, "dd/mm/yyyy")

    'Total Factura (por defecto=0)
'    text1(18).Text = "0"
'    text1(19).Text = "0"

    'em posicione en el 1r tab
    PonerFoco text1(3)
End Sub

Private Sub BotonModificar()
Dim vSec As CSeccion
    Seguir = True

    'A�adiremos el boton de aceptar y demas objetos para insertar
    ModificarTotales = False
    PonerModo 4
    
    ' cargamos la base de datos a la que apunta la seccion
    BdConta = 0
    
'    Set vSec = New CSeccion
'    If vSec.Leer(text1(3).Text) Then
'        BdConta = vSec.BdConta
'    End If
'    Set vSec = Nothing
    
    ' ### [Monica] 27/09/2006
    ' me guardo los valores anteriores de cuenta contable
    CtaAnt = text1(4).Text
    
    'Quan modifique pose en la F.Modificaci�n la data actual
    PonerFoco text1(4)
End Sub


Private Sub BotonModificarTotales()
Dim vSec As CSeccion
    Seguir = True

    'A�adiremos el boton de aceptar y demas objetos para insertar
    ModificarTotales = True
    PonerModo 4
    
    ' cargamos la base de datos a la que apunta la seccion
    BdConta = 0
'    Set vSec = New CSeccion
'    If vSec.Leer(text1(3).Text) Then
'        BdConta = vSec.BdConta
'    End If
'    Set vSec = Nothing
    
    
    'Quan modifique pose en la F.Modificaci�n la data actual
    PonerFoco text1(4)
End Sub




'Private Sub BotonRectificar()
'
'    Set frmList = New frmListado
'    'A�adiremos el boton de aceptar y demas objetos para insertar
'    frmList.CadTag = Text1(0).Text & "|" & Text1(1).Text & "|" & Text1(2).Text & "|" & Text1(3).Text & "|" & Text2(3).Text & "|" & Format(Check1(1).Value, "0") & "|"
'    frmList.OpcionListado = 12
'    frmList.Show vbModal
'
'End Sub

Private Sub BotonEliminar()
Dim Cad As String
Dim vSec As CSeccion
Dim NumFacElim As Long 'Numero de la Factura que se ha Eliminado
Dim NumSecElim As Integer 'Numero de la Seccion que se ha eliminado

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
'    'El registre de codi 0 no es pot Modificar ni Eliminar
'    If EsCodigoCero(CStr(Data1.Recordset.Fields(1).Value), FormatoCampo(text1(1))) Then Exit Sub

    Cad = "�Seguro que desea eliminar la factura?"
    Cad = Cad & vbCrLf & "Tipo: " & Format(Data1.Recordset!CodTipom, FormatoCampo(text1(0)))
    Cad = Cad & vbCrLf & "N�: " & Format(Data1.Recordset!numfactu, FormatoCampo(text1(1)))
    Cad = Cad & vbCrLf & "Fecha: " & Data1.Recordset.Fields("fecfactu")
    
    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        NumSecElim = Data1.Recordset.Fields(0)
        NumFacElim = Data1.Recordset.Fields(2)
        If Not eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            If SituarDataTrasEliminar(Data1, NumRegElim) Then
                PonerCampos
            Else
                LimpiarCampos
                'Poner los grid sin apuntar a nada
                'LimpiarDataGrids
                PonerModo 0
            End If
            'Devolvemos contador, si no estamos actualizando
'            Set vSec = New CSeccion
'            vSec.DevolverContador CStr(NumSecElim), NumFacElim
            Set vSec = Nothing
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Factura", Err.Description
End Sub

Private Sub PonerCampos()
Dim I As Integer
Dim vSec As CSeccion
    
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: pone el formato o los campos de la cabecera
    
    For I = 0 To DataGridAux.Count - 1
        CargaGrid I, True
    Next I
    
    'Recuperar Descripciones de los campos de Codigo
    '--------------------------------------------------
    Text2(4).Text = PonerNombreDeCod(text1(4), "rsocios", "nomsocio")
    Text2(25).Text = PonerNombreDeCod(text1(25), "forpago", "nomforpa")
    
    BdConta = DevuelveDesdeBDNew(cAgro, "rseccion", "empresa_conta", "codsecci", text1(3).Text, "N")
    
    Text2(27).Text = ""
    
    Set vSeccion = New CSeccion
    If vSeccion.LeerDatos(text1(3)) Then
        Text2(3).Text = vSeccion.Nombre
        If vSeccion.AbrirConta Then
            If text1(27).Text <> "" Then
                Text2(27).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", text1(27).Text, "T")
            End If
        End If
    End If
    If Not vSeccion Is Nothing Then
        vSeccion.CerrarConta
        Set vSeccion = Nothing
    End If
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    PonerModoOpcionesMenu (Modo)
End Sub

Private Sub cmdCancelar_Click()
Dim I As Integer
Dim V
    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
                LimpiarCampos
                If Data1.Recordset.EOF Then
                    PonerModo 0
                Else
                    PonerModo 2
                    PonerCampos
                End If
                PonerFoco text1(0)

        Case 4  'Modificar
                TerminaBloquear
                PonerModo 2
                PonerCampos
                PonerFoco text1(0)
        
        Case 5 'LINEAS
            Select Case ModoLineas
                Case 1 'afegir llinia
                    ModoLineas = 0
                    DataGridAux(NumTabMto).AllowAddNew = False
'                    SituarTab (NumTabMto)
                    'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar  'Modificar
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                    'If DataGridAux(NumTabMto).Enabled Then DataGridAux(NumTabMto).SetFocus
                    DataGridAux(NumTabMto).Enabled = True
                    DataGridAux(NumTabMto).SetFocus

                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        AdoAux(NumTabMto).Recordset.MoveFirst
                    End If

                Case 2 'modificar llinies
                    ModoLineas = 0
                    PonerModo 4
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        V = AdoAux(NumTabMto).Recordset.Fields(3) 'el 1 es el n� de llinia
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
                    End If
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
            End Select
            
            PosicionarData "codtipom = '" & Trim(text1(0).Text) & "' and numfactu = " & text1(1).Text & " and fecfactu = " & DBSet(text1(2).Text, "F")
            
'            If Not AdoAux(NumTabMto).Recordset.EOF Then
'                DataGridAux_RowColChange NumTabMto, 1, 1
'            Else
'                LimpiarCamposFrame NumTabMto
'            End If
    End Select
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Datos As String
Dim Sql As String
Dim UltNiv As Integer

    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm2(Me, 1) 'Comprobar formato datos ok
    If Not b Then Exit Function
    
    'cuenta contable
    If b And text1(27).Text <> "" Then
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(text1(3)) Then
            Text2(3).Text = vSeccion.Nombre
            If vSeccion.AbrirConta Then
                If text1(27).Text <> "" Then
                    Text2(27).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", text1(27).Text, "T")
                    If Text2(27) = "" Then
                        MsgBox "No existe la cuenta contable de Retenci�n en la contabilidad asociada a la secci�n", vbExclamation
                        b = False
                    End If
                End If
            End If
        End If
        If Not vSeccion Is Nothing Then
            vSeccion.CerrarConta
            Set vSeccion = Nothing
        End If
    End If
    
    'si hay porcentaje de retencion debe de haber cuenta de retencion e
    If b And text1(26).Text <> "" And text1(27).Text = "" Then
        If CInt(text1(26).Text) <> 0 Then
            MsgBox "Si hay porcentaje de retenci�n debe introducir una cuenta contable asociada. Revise.", vbExclamation
            b = False
        End If
    End If
    
    
'--monica: lo he quitado pq ha de recalcular
'    'Comprobamos que la suma de importes de las lineas es igual al total de la factura
'    If b And Modo <> 3 Then
'        Datos = SumaLineas("")
'
'        If CCur(Datos) > CCur(TransformaPuntosComas(DBSet(text1(6).Text, "N"))) + CCur(TransformaPuntosComas(DBSet(text1(12).Text, "N"))) + CCur(TransformaPuntosComas(DBSet(text1(18).Text, "N"))) Then
'            MsgBox "La suma de los importes de lineas es mayor que el total de la factura!!!", vbExclamation
'            b = False
'        ElseIf CCur(Datos) < CCur(TransformaPuntosComas(DBSet(text1(6).Text, "N"))) + CCur(TransformaPuntosComas(DBSet(text1(12).Text, "N"))) + CCur(TransformaPuntosComas(DBSet(text1(18).Text, "N"))) Then
'            MsgBox "La suma de los importes de lineas es menor que el total de la factura!!!", vbExclamation
'            b = False
'        End If
'    End If
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub PosicionarData(Cad As String)
'Dim cad As String
Dim Indicador As String
    
  '  cad = ""
    If SituarDataMULTI(Data1, Cad, Indicador) Then
        If ModoLineas <> 1 Then
            PonerModo 2
        End If
       
       lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       'Poner los grid sin apuntar a nada
       'LimpiarDataGrids
       PonerModo 0
    End If
End Sub

Private Function eliminar() As Boolean
Dim vWhere As String
Dim vTipoMov As CTiposMov

    On Error GoTo FinEliminar
        
    conn.BeginTrans
    vWhere = ObtenerWhereCab(True)

    'Eliminar las Lineas de facturas de proveedor
    conn.Execute "DELETE FROM fvarlinfactpro " & Replace(vWhere, "fvarcabfactpro", "fvarlinfactpro")
    
    'Eliminar la CABECERA
    conn.Execute "Delete from " & NomTabla & vWhere
               
    'Decrementar contador si borramos el ultima factura
    Set vTipoMov = New CTiposMov
    vTipoMov.DevolverContador CodTipoMov, Val(text1(1).Text)
    Set vTipoMov = Nothing
                 
               
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        conn.RollbackTrans
        eliminar = False
    Else
        conn.CommitTrans
        eliminar = True
    End If
End Function

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco text1(Index), Modo
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim Cad As String, Datos As String
Dim Suma As Currency
Dim I As Integer

    If Not PerderFocoGnral(text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 1 'N� factura
            If text1(Index).Text <> "" Then FormateaCampo text1(Index)
                        
        Case 2 'Fecha
            If text1(Index).Text <> "" Then PonerFormatoFecha text1(Index)
            
        Case 3 'Seccion
            If text1(Index).Text <> "" Then
                If PonerFormatoEntero(text1(3)) Then
                    Text2(Index).Text = PonerNombreDeCod(text1(Index), "rseccion", "nomsecci", "codsecci", "N")
                    If Text2(Index).Text = "" Then
                        Cad = "No existe la Secci�n: " & text1(Index).Text & vbCrLf
                        Cad = Cad & "�Desea crearla?" & vbCrLf
                        If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
                            Set frmSec = New frmManSeccion
                            frmSec.DatosADevolverBusqueda = "0|1|"
                            text1(Index).Text = ""
                            TerminaBloquear
                            frmSec.Show vbModal
                            Set frmSec = Nothing
                            If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                        Else
                            text1(Index).Text = ""
                        End If
                        PonerFoco text1(Index)
                    Else
                        'recuperar el numero de contabilidad
                        BdConta = DevuelveDesdeBDNew(cAgro, "rseccion", "empresa_conta", "codsecci", text1(3).Text, "N")
                        If DBLet(BdConta, "N") = 0 Then
                            MsgBox "Esta seccion no est� asociada a ninguna contabilidad. Revise.", vbExclamation
                            text1(Index).Text = ""
                            PonerFoco text1(Index)
                        Else
                            
                        End If
                    End If
                Else
                    Text2(Index).Text = ""
                End If
            End If
            
        
        Case 4 ' Socio
            If text1(Index).Text <> "" Then
                If PonerFormatoEntero(text1(Index)) Then
                    Text2(Index).Text = PonerNombreDeCod(text1(Index), "rsocios", "nomsocio", "codsocio", "N")
                    If Text2(Index).Text = "" Then
                        Cad = "No existe el Socio: " & text1(Index).Text & vbCrLf
                        Cad = Cad & "�Desea crearlo?" & vbCrLf
                        If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
                            Set frmSoc = New frmManSocios
                            frmSoc.DatosADevolverBusqueda = "0|1|"
                            text1(Index).Text = ""
                            TerminaBloquear
                            frmSoc.Show vbModal
                            Set frmSoc = Nothing
                            If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                        Else
                            text1(Index).Text = ""
                        End If
                        PonerFoco text1(Index)
                    End If
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        Case 25 'Forma pago
            If text1(Index).Text = "" Then Exit Sub
            
            Text2(25).Text = DevuelveDesdeBDNew(cAgro, "forpago", "nomforpa", "codforpa", text1(25).Text, "N")
            If Text2(25).Text = "" Then
                MsgBox "No existe la Forma de Pago. Reintroduzca.", vbExclamation
                Seguir = False
                PonerFoco text1(Index)
            Else
                Seguir = True
            End If

        Case 26 'porcentaje de retencion
            If Modo = 1 Then Exit Sub
            PonerFormatoDecimal text1(Index), 7
            
        Case 8, 10, 14, 16, 20, 22, 24
            If Modo = 1 Then Exit Sub
            PonerFormatoDecimal text1(Index), 7
            
        Case 5 'despues de las observaciones si estamos insertando despues he de ir al campo de retencion
            If Modo = 3 And Seguir Then PonerFoco text1(26)
            
        Case 6, 9, 11, 12, 15, 17, 18, 21, 23    'IMPORTES Base, IVA
            PonerFormatoDecimal text1(Index), 1
            
        Case 7, 13, 19 'cod. IVA
           If text1(Index).Text = "" Then
              text1(Index + 1).Text = ""
           Else
                Set vSeccion = New CSeccion
                If vSeccion.LeerDatos(text1(3).Text) Then
                    If vSeccion.AbrirConta Then
                        text1(Index + 1).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", text1(Index).Text, "N")
                    End If
                End If
                If Not vSeccion Is Nothing Then
                    vSeccion.CerrarConta
                    Set vSeccion = Nothing
                End If
           End If
              
        Case 27 'cuenta de retencion
            Text2(Index).Text = ""
            If text1(Index).Text = "" Then Exit Sub
            
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(text1(3).Text) Then
                If vSeccion.AbrirConta Then
                    Text2(27) = PonerNombreCuenta(text1(27), Modo)
                    If Text2(Index).Text = "" Then
                        PonerFoco text1(Index)
                    End If
                End If
            End If
            If Not vSeccion Is Nothing Then
                vSeccion.CerrarConta
                Set vSeccion = Nothing
            End If
            
    End Select

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 2: KEYFecha KeyAscii, 2
'                Case 3: KEYBusqueda KeyAscii, 0
'                Case 4: KEYBusqueda KeyAscii, 1
'                Case 5: KEYBusqueda KeyAscii, 2
'                Case 7: KEYBusqueda KeyAscii, 3
'                Case 11: KEYBusqueda KeyAscii, 4
'                Case 15: KEYBusqueda KeyAscii, 5
'               ' Case 1: KEYFecha KeyAscii, 1
            End Select
        End If
    Else
        If Not text1(Index).MultiLine Then
            KEYpress KeyAscii
        End If
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
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

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFec_Click (Indice)
End Sub

Private Sub KEYBusquedaLin(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    btnBuscar_Click (Indice)
End Sub

'************* LLINIES: ****************************
Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
Dim Cad As String
'12/02/2008: lo he quitado porque lo modificaran ellos manualmente en la contabilidad
'    If vParamAplic.NumeroConta <> 0 And _
'       Not FacturaModificable(text1(0).Text, text1(1).Text, text1(2).Text, Check1(1).Value) Then Exit Sub
    '++monica:12/02/2008
     If CByte(Data1.Recordset!intconta) = 1 Then
        Cad = "   Se dispone a realizar cambios en los datos de la Factura.     " & vbCrLf & vbCrLf & _
              "Recuerde modificar la Contabilidad y Tesoreria correspondiente!!!"
        MsgBox Cad, vbExclamation
     End If
    '++
    
    
     Select Case Button.Index
        Case 1
'            TerminaBloquear
            BotonAnyadirLinea Index
        Case 2
'            TerminaBloquear
            BotonModificarLinea Index
        Case 3
'            TerminaBloquear
            BotonEliminarLinea Index
            If Modo = 4 Then
                If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
            End If
        Case 6 'Imprimir
'            BotonImprimirLinea Index
    End Select
End Sub

Private Sub BotonEliminarLinea(Index As Integer)
Dim Sql As String
Dim eliminar As Boolean

    On Error GoTo Error2

    ModoLineas = 3 'Posem Modo Eliminar Ll�nia

    If Modo = 4 Then 'Modificar Cap�alera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If

    NumTabMto = Index
    PonerModo 5

'    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar(Index) Then Exit Sub
    
    If AdoAux(Index).Recordset.RecordCount = 1 Then
        MsgBox "No se puede borrar un �nica l�nea de factura, elimine la factura completa", vbExclamation
        PonerModo 2
        Exit Sub
    End If
    
    
    eliminar = False

    Select Case Index
        Case 0 'lineas de factura
            Sql = "�Seguro que desea eliminar la l�nea?"
            Sql = Sql & vbCrLf & "N� l�nea: " & Format(DBLet(AdoAux(Index).Recordset!numlinea), FormatoCampo(txtAux(4)))
            Sql = Sql & vbCrLf & "Concepto: " & DBLet(AdoAux(Index).Recordset!codConce) '& "  " & txtAux(4).Text
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
                eliminar = True
                Sql = "DELETE FROM fvarlinfactpro"
                Sql = Sql & Replace(ObtenerWhereCab(True), "fvarcabfactpro", "fvarlinfactpro") & " AND numlinea= " & AdoAux(Index).Recordset!numlinea
            End If
    End Select

    If eliminar Then
        TerminaBloquear
'        conn.Execute Sql
        CadenaBorrado = Sql
        '16022007
        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
                ModificaImportes = True
                BotonModificar
                cmdAceptar_Click
        End If
'        EliminarLinea
        
        
        'antes estaba debajo de situardata
        CargaGrid Index, True
        SituarDataTrasEliminar AdoAux(Index), NumRegElim, True
        
        
        
    End If

    ModoLineas = 0
    PosicionarData "codtipom = '" & Trim(text1(0).Text) & "' and numfactu = " & text1(1).Text & " and fecfactu = " & DBSet(text1(2).Text, "F")

    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub

Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vTabla As String
Dim anc As Single
Dim I As Integer
Dim SumLin As Currency
Dim vSec As CSeccion

    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    'If ModificaLineas = 2 Then Exit Sub
    ModoLineas = 1 'Ponemos Modo A�adir Linea

    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modifcar Cabecera
        cmdAceptar_Click
        'No se ha insertado la cabecera
        If ModoLineas = 0 Then Exit Sub
    End If

    NumTabMto = Index
    PonerModo 5
'    If b Then BloquearText1 Me, 4 'Si viene de Insertar Cabecera no bloquear los Text1


    'Obtener el numero de linea ha insertar
    Select Case Index
        Case 0: vTabla = "fvarlinfactpro"
    End Select
    'Obtener el sig. n� de linea a insertar
    vWhere = Replace(ObtenerWhereCab(False), "fvarcabfactpro", vTabla)
    NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", vWhere)

    'Situamos el grid al final
    AnyadirLinea DataGridAux(Index), AdoAux(Index)

    anc = DataGridAux(Index).Top
    If DataGridAux(Index).Row < 0 Then
        anc = anc + 210
    Else
        anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
    End If

    LLamaLineas Index, ModoLineas, anc

    Select Case Index
        Case 0 'lineas factura
            txtAux(0).Text = text1(3).Text 'seccion
            txtAux(1).Text = text1(0).Text 'tipo de movimiento
            txtAux(2).Text = text1(1).Text 'factura
            txtAux(3).Text = text1(2).Text 'fecha
            txtAux(4).Text = NumF 'numlinea
'            FormateaCampo txtAux(3)
            For I = 5 To txtAux.Count - 1
                txtAux(I).Text = ""
            Next I
            txtAux2(0).Text = ""

            'desbloquear la linea (se bloquea al a�adir)
'            BloquearTxt txtAux(3), False
            PonerFoco txtAux(5)
    End Select
End Sub

Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim I As Integer
    Dim J As Integer
    Dim vSec As CSeccion
    
    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub
    
    ModoLineas = 2 'Modificar ll�nia
    
    If Modo = 4 Then 'Modificar Cabecera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
    
    ' cargamos la base de datos a la que apunta la seccion
    BdConta = 0
    
    Set vSeccion = New CSeccion
    If vSeccion.LeerDatos(text1(3)) Then
        Text2(3).Text = vSeccion.Nombre
        If vSeccion.AbrirConta Then
        
        End If
    End If
    
    
    NumTabMto = Index
    PonerModo 5
    
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

    Select Case Index
        Case 0 'lineas de factura
            For J = 1 To 5
                txtAux(J).Text = DataGridAux(Index).Columns(J - 1).Text
            Next J
            txtAux2(0).Text = DataGridAux(Index).Columns(5).Text 'DevuelveDesdeBDNew(cPTours, "concefact", "nomconce", "codconce", DataGridAux(Index).Columns(5).Text, "N")
            txtAux(8).Text = DataGridAux(Index).Columns(6).Text 'DevuelveDesdeBDNew(cPTours, "concefact", "tipoiva", "codconce", DataGridAux(Index).Columns(5).Text, "N")
            txtAux(6).Text = DataGridAux(Index).Columns(7).Text    ' ampliacion
            txtAux(7).Text = DataGridAux(Index).Columns(10).Text   ' importe
            txtAux(9).Text = DataGridAux(Index).Columns(8).Text    ' cantidad
            txtAux(10).Text = DataGridAux(Index).Columns(9).Text  ' precio
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    Select Case Index
        Case 0 'lineas de factura
            PonerFoco txtAux(5)
    End Select
End Sub

Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    On Error GoTo ELLamaLin

    DeseleccionaGrid DataGridAux(Index)
    
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Lineas
    Select Case Index
        Case 0 'lineas de factura
            For jj = 5 To txtAux.Count - 1
                txtAux(jj).Top = alto
                txtAux(jj).visible = b
            Next jj
            txtAux(8).visible = False
            txtAux(8).Enabled = False
            
            txtAux2(0).Top = alto
            txtAux2(0).visible = b
            Me.btnBuscar(0).Top = alto
            Me.btnBuscar(0).visible = b
    End Select
    
ELLamaLin:
    Err.Clear
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) Then
            Select Case Index
                Case 5: KEYBusquedaLin KeyAscii, 0
                Case 6: KEYBusquedaLin KeyAscii, 1
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Sql As String
    txtAux(Index).Text = Trim(txtAux(Index).Text)

    Select Case Index
        Case 6 ' Ampliacion
            txtAux(Index).Text = UCase(txtAux(Index).Text)

        Case 5 ' Concepto
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(0).Text = PonerNombreDeCod(txtAux(Index), "fvarconce", "nomconce", "codconce", "N")
                txtAux(8).Text = PonerNombreDeCod(txtAux(Index), "fvarconce", "tipoiva", "codconce", "N")
                
                If txtAux2(0).Text = "" Then
                    cadMen = "No existe el Concepto: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCon = New frmFVARConceptos
                        frmCon.DatosADevolverBusqueda = "0|1|"
                        frmCon.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmCon.Show vbModal
                        Set frmCon = Nothing
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                Else
                    cadMen = DevuelveDesdeBDNew(cAgro, "fvarconce", "codsecci", "codconce", txtAux(Index), "N")
                    If CInt(ComprobarCero(cadMen)) <> CInt(text1(3).Text) Then
                        MsgBox "El concepto ha de ser de la misma secci�n. Reintroduzca.", vbExclamation
                        txtAux(Index).Text = ""
                        PonerFoco txtAux(5)
                    End If
                End If
            Else
                txtAux2(0).Text = ""
            End If

        Case 9 ' cantidad
            If PonerFormatoDecimal(txtAux(Index), 3) Then
                txtAux(7).Text = Round2(CCur(txtAux(9).Text) * CCur(ComprobarCero(txtAux(10).Text)), 2)
                PonerFormatoDecimal txtAux(7), 3
            Else
                txtAux(Index).Text = ""
            End If
            
        Case 10 ' precio
            If PonerFormatoDecimal(txtAux(Index), 11) Then
                txtAux(7).Text = Round2(CCur(ComprobarCero(txtAux(9).Text)) * CCur(txtAux(10).Text), 2)
                PonerFormatoDecimal txtAux(7), 3
            Else
                txtAux(Index).Text = ""
            End If
        
        Case 7 'Importe
'           If Not EsNumerico(txtAux(Index).Text) Then
'                MsgBox "El Importe debe ser num�rico.", vbExclamation
'                On Error Resume Next
'                txtAux(Index).Text = ""
'                PonerFoco txtAux(Index)
'                Exit Sub
'            End If
            'Es numerico
            If PonerFormatoDecimal(txtAux(Index), 3) Then
                PonerFocoBtn Me.cmdAceptar
            Else
                txtAux(Index).Text = ""
            End If
    End Select
    
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ' si vamos a insertar el importe miramos si podemos calcularlo y no entrar en importe
    If Index = 7 And (txtAux(9).Text <> "" Or txtAux(10).Text <> "") And txtAux(Index).Text = "" Then
        txtAux(Index).Text = Round2(ComprobarCero(txtAux(9).Text) * ComprobarCero(txtAux(10).Text), 2)
'        cmdAceptar.SetFocus
        Exit Sub
    End If
    
    ConseguirFocoLin txtAux(Index)
End Sub

Private Function DatosOkLlin(nomframe As String) As Boolean
Dim b As Boolean
Dim SumLin As Currency
    
    On Error GoTo EDatosOKLlin

    DatosOkLlin = False
        
    b = CompForm2(Me, 2, nomframe) 'Comprobar formato datos ok
    If Not b Then Exit Function
    
' ### [Monica] 29/09/2006
' he quitado la parte de comprobar la suma de lineas
'    'Comprobar que el Importe del total de las lineas suma el total o menos de la factura
'    SumLin = CCur(SumaLineas(txtAux(4).Text))
'
'    'Le a�adimos el importe de linea que vamos a insertar
'    SumLin = SumLin + CCur(txtAux(7).Text)
'
'    'comprobamos que no sobrepase el total de la factura
'    If SumLin > CCur(Text1(18).Text) Then
'        MsgBox "La suma del importe de las lineas no puede ser superior al total de la factura.", vbExclamation
'        b = False
'    End If
    
    DatosOkLlin = b
    
EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function SepuedeBorrar(ByRef Index As Integer) As Boolean

    SepuedeBorrar = False
    If AdoAux(Index).Recordset.EOF Then Exit Function

    SepuedeBorrar = True
End Function

Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
Dim b As Boolean
Dim tots As String

    On Error GoTo ECarga

    'b = DataGridAux(Index).Enabled
    'DataGridAux(Index).Enabled = False
    
    tots = MontaSQLCarga(Index, enlaza)
    
    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    Select Case Index
        Case 0 'lineas de factura
            'si es visible|control|tipo campo|nombre campo|ancho control|formato campo|
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux(5)|T|C�digo|700|;S|btnBuscar(0)|B|||;S|txtAux2(0)|T|Concepto|2100|;S|txtAux(8)|T|T.Iva|550|;"
            tots = tots & "S|txtAux(6)|T|Ampliaci�n|4300|;S|txtAux(9)|T|Cantidad|1000|;S|txtAux(10)|T|Precio|1000|;S|txtAux(7)|T|Importe|1200|;"
            arregla tots, DataGridAux(Index), Me
'           DataGridAux(Index).Columns(6).Alignment = dbgCenter
'           DataGridAux(Index).Columns(9).Alignment = dbgRight
    End Select
    
    DataGridAux(Index).ScrollBars = dbgAutomatic
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub

Private Sub InsertarLinea()
'Inserta registro en las tablas de Lineas: provbanc, provdpto
Dim nomframe As String
Dim b As Boolean
Dim V As Integer

' variables para el recalculo de iva y totales
    Dim I As Integer
    Dim Imptot(2)
    Dim Tipiva(2)
    Dim Impbas(2) As Currency
    Dim ImpIVA(2) As Currency
    Dim PorIva(2) As Currency
    Dim ImpREC(2) As Currency
    Dim PorRec(2) As Currency
    Dim TotFac As Currency

'retencion
    Dim PorRet As Currency
    Dim ImpRet As Currency

    On Error Resume Next

    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'lineas de factura
    End Select

    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomframe) Then
            b = BLOQUEADesdeFormulario2(Me, Data1, 1)
            CargaGrid NumTabMto, True
            V = AdoAux(NumTabMto).Recordset.Fields(4) 'el 2 es el n� de llinia
            ModoLineas = 0
            CargaGrid NumTabMto, True
'            SituarTab (NumTabMto)
            DataGridAux(NumTabMto).SetFocus
            AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(4).Name & " =" & V)
            
'            ' ### [Monica] 29/09/2006
'            ' a�adido el tema de de recalculo de bases
'            If BdConta = 0 Then Exit Sub
            
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(text1(3)) Then
                Text2(3).Text = vSeccion.Nombre
                If vSeccion.AbrirConta Then
                    PorRet = 0
                    If text1(26).Text <> "" Then PorRet = CCur(ImporteSinFormato(text1(26).Text))
                    
                    RecalculoBasesIvaFactura AdoAux(0).Recordset, Imptot, Tipiva, Impbas, ImpIVA, PorIva, TotFac, ImpREC, PorRec, PorRet, ImpRet
                End If
            End If
            If Not vSeccion Is Nothing Then
                vSeccion.CerrarConta
                Set vSeccion = Nothing
            End If

            '13/02/2007 iniacializo los txt
            For I = 0 To 2
                text1(6 + (6 * I)).Text = ""
                text1(7 + (6 * I)).Text = ""
                text1(8 + (6 * I)).Text = ""
                text1(9 + (6 * I)).Text = ""
                text1(10 + (6 * I)).Text = ""
                text1(11 + (6 * I)).Text = ""
            Next I
            text1(26).Text = ""
            text1(28).Text = ""
            
            '13/02/2007 he a�adido las condiciones del for antes solo estaban las sentencias
            For I = 0 To 2
                 If Tipiva(I) <> 0 Then
                    text1(6 + (6 * I)).Text = Impbas(I)
                    text1(7 + (6 * I)).Text = Tipiva(I)
                    text1(8 + (6 * I)).Text = PorIva(I)
                    text1(9 + (6 * I)).Text = ImpIVA(I)
                    If PorRec(I) <> 0 Then text1(10 + (6 * I)).Text = PorRec(I)
                    If ImpREC(I) <> 0 Then text1(11 + (6 * I)).Text = ImpREC(I)
                 End If
'12/03/2007
'                 If Impbas(i) <> 0 Then text1(6 + (6 * i)).Text = Impbas(i)
'                 If PorIva(i) <> 0 Then text1(8 + (6 * i)).Text = PorIva(i)
'                 If impiva(i) <> 0 Then text1(9 + (6 * i)).Text = impiva(i)
'                 If PorRec(i) <> 0 Then text1(10 + (6 * i)).Text = PorRec(i)
'                 If ImpRec(i) <> 0 Then text1(11 + (6 * i)).Text = ImpRec(i)

                 'TotFac = Impbas(i) + impiva(i)
            Next I
            If PorRet <> 0 Then text1(26).Text = PorRet
            If ImpRet <> 0 Then text1(28).Text = ImpRet
            text1(24).Text = TotFac

            If text1(8).Text = "" Then text1(8).Text = "0,00"
            If text1(9).Text = "" Then text1(9).Text = "0,00"
            
            
'++monica: 10/03/2009
            PonerFormatos
'++
            
            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
'                Modo = 4
'                PonerModo Modo
'                ClienteAnt = Text1(3).Text
'                FormaPagoAnt = Text1(5).Text
                ModificaImportes = True
                BotonModificar
                cmdAceptar_Click

            End If

            LLamaLineas NumTabMto, 0
            
            If b Then BotonAnyadirLinea NumTabMto
        End If
    End If
End Sub

Private Sub ModificarLinea()
'Modifica registro en las tablas de Lineas: provbanc, provdpto
Dim nomframe As String
Dim V As Currency

' variables para el recalculo de iva y totales
    Dim I As Integer
    Dim Imptot(2)
    Dim Tipiva(2)
    Dim Impbas(2) As Currency
    Dim ImpIVA(2) As Currency
    Dim PorIva(2) As Currency
    Dim ImpREC(2) As Currency
    Dim PorRec(2) As Currency
    Dim TotFac As Currency
    
    'retencion
    Dim PorRet As Currency
    Dim ImpRet As Currency


    On Error GoTo EModificarLin

    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'lineas de factura
    End Select
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
'        conn.BeginTrans
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
            
            ' ### [Monica] 29/09/2006
            ' he quitado el boton modificar para recalcular bases e iva
            
            'BotonModificar
                

                
            End If
            V = AdoAux(NumTabMto).Recordset.Fields(4) 'el 2 es el n� de llinia
            ModoLineas = 0
            CargaGrid NumTabMto, True
'            SituarTab (NumTabMto)
            DataGridAux(NumTabMto).SetFocus
'            AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
            
'            ' ### [Monica] 29/09/2006
'            ' a�adido el tema de de recalculo de bases
'            If BdConta = 0 Then Exit Sub
            
            
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(text1(3)) Then
                If vSeccion.AbrirConta Then
                
                    PorRet = 0
                    If text1(26).Text <> "" Then PorRet = CCur(ImporteSinFormato(text1(26).Text))
        
                    RecalculoBasesIvaFactura AdoAux(0).Recordset, Imptot, Tipiva, Impbas, ImpIVA, PorIva, TotFac, ImpREC, PorRec, PorRet, ImpRet
                    
                End If
            End If
            
            If Not vSeccion Is Nothing Then
                vSeccion.CerrarConta
                Set vSeccion = Nothing
            End If

            '13/02/2007 iniacializo los txt
            For I = 0 To 2
                text1(6 + (6 * I)).Text = ""
                text1(7 + (6 * I)).Text = ""
                text1(8 + (6 * I)).Text = ""
                text1(9 + (6 * I)).Text = ""
                text1(10 + (6 * I)).Text = ""
                text1(11 + (6 * I)).Text = ""
            Next I

            '13/02/2007 he a�adido las condiciones del for antes solo estaban las sentencias
            For I = 0 To 2
                 If Impbas(I) <> 0 Then text1(6 + (6 * I)).Text = Impbas(I)
                 If Tipiva(I) <> 0 Then text1(7 + (6 * I)).Text = Tipiva(I)
                 If PorIva(I) <> 0 Then text1(8 + (6 * I)).Text = PorIva(I)
                 If ImpIVA(I) <> 0 Then text1(9 + (6 * I)).Text = ImpIVA(I)
                 If PorRec(I) <> 0 Then text1(10 + (6 * I)).Text = PorRec(I)
                 If ImpREC(I) <> 0 Then text1(11 + (6 * I)).Text = ImpREC(I)

                 'TotFac = Impbas(i) + impiva(i)
            Next I
            text1(24).Text = TotFac
            If ImpRet <> 0 Then text1(28).Text = ImpRet
            
            If text1(8).Text = "" Then text1(8).Text = "0,00"
            If text1(9).Text = "" Then text1(9).Text = "0,00"
            
'++monica: 10/03/2009
            PonerFormatos
'++
            
            
            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
'                Modo = 4
'                PonerModo Modo
'                ClienteAnt = Text1(3).Text
'                FormaPagoAnt = Text1(5).Text
                ModificaImportes = True
'--monica:10/03/2009
'                PonerCamposForma Me, Me.Data1
                BotonModificar
                cmdAceptar_Click

            End If

            LLamaLineas NumTabMto, 0
        End If
    End If
    Exit Sub
    
EModificarLin:
    If Err.Number <> 0 Then MuestraError Err.Number, "Modificar Linea", Err.Description
End Sub

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    vWhere = ""
    If conW Then vWhere = " WHERE "
    vWhere = vWhere & "codtipom='" & Trim(text1(0).Text) & "'"
    vWhere = vWhere & " AND numfactu= " & text1(1).Text & " AND fecfactu= '" & Format(text1(2).Text, FormatoFecha) & "'"
    ObtenerWhereCab = vWhere
End Function



Private Function SumaLineas(NumLin As String) As String
'Al Insertar o Modificar linea sumamos todas las lineas excepto la que estamos
'Insertando o modificando que su valor sera el del txtaux(4).text
'En el DatosOK de la factura sumamos todas las lineas
Dim Sql As String
Dim RS As ADODB.Recordset
Dim SumLin As Currency

    SumLin = 0
    Sql = "SELECT SUM(importe) FROM fvarlinfactpro "
    Sql = Sql & Replace(ObtenerWhereCab(True), "fvarcabfactpro", "fvarlinfactpro")
    If NumLin <> "" Then Sql = Sql & " AND numlinea<>" & DBSet(txtAux(4).Text, "N") 'numlinea
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        'En SumLin tenemos la suma de las lineas ya insertadas
        SumLin = CCur(DBLet(RS.Fields(0), "N"))
    End If
    RS.Close
    Set RS = Nothing
    SumaLineas = CStr(SumLin)
End Function


Private Function FacturaModificable(letraser As String, numfactu As String, fecfactu As String, Contabil As String) As Boolean

    FacturaModificable = False
    
    If Contabil = 0 Then
        FacturaModificable = True
    Else
        ' si la factura esta contabilizada tenemos que ver si en la contabilidad esta contabilizada y
        ' si en la tesoreria esta remesada o cobrada en estos casos la factura no puede ser modificada
        If FacturaContabilizada(letraser, numfactu, Year(CDate(fecfactu))) Then
            MsgBox "Factura contabilizada en la Contabilidad, no puede modificarse ni eliminarse."
            Exit Function
        End If
        
        If FacturaRemesada(letraser, numfactu, fecfactu) Then
            MsgBox "Factura Remesada, no puede modificarse ni eliminarse."
            Exit Function
        End If
        
        If FacturaCobrada(letraser, numfactu, fecfactu) Then
            MsgBox "Factura Cobrada, no puede modificarse ni eliminarse."
            Exit Function
        End If
           
        FacturaModificable = True
    End If

End Function

'VRS:2.0.1(3)
Private Sub LlamarImprimir()
    With frmImprimir
        'Nuevo. Febrero 2010
        .outClaveNombreArchiv = text1(0).Text & Format(text1(1).Text, "0000000")
        .outCodigoCliProv = text1(4).Text
        .outTipoDocumento = 1
    
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadparam
        .NumeroParametros = 2
        .SoloImprimir = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .EnvioEMail = False
        .Opcion = 1
        .Show vbModal
    End With
End Sub


Private Sub ActivarFrameCobros()
Dim obj As Object

For Each obj In Me
    If TypeOf obj Is Frame Then
        If obj.Name = "FrameCobros" Then
            
            
        End If
        
    End If
Next obj

End Sub


Private Sub EliminarLinea()
Dim nomframe As String
Dim V As Currency
Dim Sql As String

    
 
' variables para el recalculo de iva y totales
    Dim I As Integer
    Dim Imptot(2)
    Dim Tipiva(2)
    Dim Impbas(2) As Currency
    Dim ImpIVA(2) As Currency
    Dim PorIva(2) As Currency
    Dim ImpREC(2) As Currency
    Dim PorRec(2) As Currency
    Dim TotFac As Currency

    'retencion
    Dim PorRet As Currency
    Dim ImpRet As Currency


    On Error GoTo EEliminarLin

    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'lineas de factura
    End Select
    

    TerminaBloquear
'        conn.BeginTrans
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then

            ' ### [Monica] 29/09/2006
            ' he quitado el boton modificar para recalcular bases e iva

            'BotonModificar

            End If
            ModoLineas = 0
'            V = AdoAux(NumTabMto).Recordset.Fields(4) 'el 2 es el n� de llinia
            CargaGrid NumTabMto, True

'            SituarTab (NumTabMto)

' [Monica] 25/01/2010 Daba error cuando elimina linea he quitado el setfocus
'            DataGridAux(NumTabMto).SetFocus

'            AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)

'            ' ### [Monica] 29/09/2006
'            ' a�adido el tema de de recalculo de bases
            PorRet = 0
            If text1(26).Text <> "" Then PorRet = CCur(ImporteSinFormato(text1(26).Text))

            RecalculoBasesIvaFactura AdoAux(0).Recordset, Imptot, Tipiva, Impbas, ImpIVA, PorIva, TotFac, ImpREC, PorRec, PorRet, ImpRet


            '13/02/2007 iniacializo los txt
            For I = 0 To 2
                text1(6 + (6 * I)).Text = ""
                text1(7 + (6 * I)).Text = ""
                text1(8 + (6 * I)).Text = ""
                text1(9 + (6 * I)).Text = ""
                text1(10 + (6 * I)).Text = ""
                text1(11 + (6 * I)).Text = ""
            Next I

            '13/02/2007 he a�adido las condiciones del for antes solo estaban las sentencias
            For I = 0 To 2
                 If Impbas(I) <> 0 Then text1(6 + (6 * I)).Text = Impbas(I)
                 If Tipiva(I) <> 0 Then text1(7 + (6 * I)).Text = Tipiva(I)
                 If PorIva(I) <> 0 Then text1(8 + (6 * I)).Text = PorIva(I)
                 If ImpIVA(I) <> 0 Then text1(9 + (6 * I)).Text = ImpIVA(I)
                 If PorRec(I) <> 0 Then text1(10 + (6 * I)).Text = PorRec(I)
                 If ImpREC(I) <> 0 Then text1(11 + (6 * I)).Text = ImpREC(I)

                 'TotFac = Impbas(i) + impiva(i)
            Next I
            
            text1(24).Text = TotFac
            If ImpRet <> 0 Then text1(28).Text = ImpRet
            
            If text1(8).Text = "" Then text1(8).Text = "0,00"
            If text1(9).Text = "" Then text1(9).Text = "0,00"
            
            
'            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
'                ModificaImportes = True
'                BotonModificar
'                cmdAceptar_Click
'            End If

'++monica: 10/03/2009
            PonerFormatos
'++
            LLamaLineas NumTabMto, 0
    Exit Sub
    
EEliminarLin:
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Linea", Err.Description
End Sub

Private Sub PonerFormatos()
Dim mTag As CTag
Dim I As Integer

    Set mTag = New CTag
    For I = 6 To 24
        mTag.Cargar text1(I)
        If mTag.Formato <> "" And CStr(text1(I).Text) <> "" Then
             text1(I).Text = Format(text1(I).Text, mTag.Formato)
        End If
    Next I
    Set mTag = Nothing

End Sub

Private Sub AbrirFrmForpa(Indice As Integer)
    indCodigo = Indice
    Set frmFpa = New frmComFpa
    frmFpa.DatosADevolverBusqueda = "0|1|"
    frmFpa.CodigoActual = text1(indCodigo)
    frmFpa.Show vbModal
    Set frmFpa = Nothing
End Sub



Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String

    On Error GoTo EInsertarCab
    
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        text1(1).Text = vTipoMov.ConseguirContador(CodTipoMov)
        Sql = CadenaInsertarDesdeForm(Me)
        If Sql <> "" Then
            If InsertarOferta(Sql, vTipoMov) Then
                CadenaConsulta = "Select * from fvarcabfactpro " & ObtenerWhereCab(True) & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
                'Ponerse en Modo Insertar Lineas
                BotonAnyadirLinea 0
                
'                CalcularTotales
            End If
        End If
        text1(0).Text = Format(text1(0).Text, "0000000")
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

    On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Factura
    'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
    Do
        devuelve = DevuelveDesdeBDNew(cAgro, "fvarcabfactpro", "numfactu", "numfactu", text1(1).Text, "N", , "fecfactu", text1(2).Text, "F", "codtipom", text1(0).Text, "T")
        If devuelve <> "" Then
            'Ya existe el contador incrementarlo
            Existe = True
            vTipoMov.IncrementarContador (CodTipoMov)
            text1(1).Text = vTipoMov.ConseguirContador(CodTipoMov)
            cambiaSQL = True
        Else
            Existe = False
        End If
    Loop Until Not Existe
    If cambiaSQL Then vSQL = CadenaInsertarDesdeForm(Me)
    
    
    'Aqui empieza transaccion
    conn.BeginTrans
    MenError = "Error al insertar en la tabla Cabecera de Factura (fvarcabfactpro)."
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



