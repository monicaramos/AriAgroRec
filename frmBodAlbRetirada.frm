VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBodAlbRetirada 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Albaranes de Retirada de Vino y Aceite"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   12360
   Icon            =   "frmBodAlbRetirada.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBodAlbRetirada.frx":000C
   ScaleHeight     =   8160
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameFactura 
      Caption         =   "Totales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2715
      Left            =   6360
      TabIndex        =   38
      Top             =   660
      Width           =   5940
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   4410
         MaxLength       =   15
         TabIndex        =   60
         Text            =   "Text1 7"
         Top             =   570
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   21
         Left            =   90
         MaxLength       =   5
         TabIndex        =   59
         Text            =   "Text1 7"
         Top             =   1875
         Width           =   435
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   90
         MaxLength       =   5
         TabIndex        =   58
         Text            =   "Text1 7"
         Top             =   1545
         Width           =   435
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   90
         MaxLength       =   5
         TabIndex        =   57
         Text            =   "Text1 7"
         Top             =   1230
         Width           =   435
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   540
         MaxLength       =   15
         TabIndex        =   56
         Text            =   "Text1 7"
         Top             =   1230
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   1875
         MaxLength       =   5
         TabIndex        =   55
         Text            =   "Text1 7"
         Top             =   1230
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   16
         Left            =   2400
         MaxLength       =   15
         TabIndex        =   54
         Text            =   "Text1 7"
         Top             =   1230
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   540
         MaxLength       =   15
         TabIndex        =   53
         Text            =   "Text1 7"
         Top             =   1560
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   1875
         MaxLength       =   5
         TabIndex        =   52
         Text            =   "Text1 7"
         Top             =   1545
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   20
         Left            =   2400
         MaxLength       =   15
         TabIndex        =   51
         Text            =   "Text1 7"
         Top             =   1545
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   23
         Left            =   540
         MaxLength       =   15
         TabIndex        =   50
         Text            =   "Text1 7"
         Top             =   1875
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   22
         Left            =   1890
         MaxLength       =   5
         TabIndex        =   49
         Text            =   "Text1 7"
         Top             =   1875
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   24
         Left            =   2400
         MaxLength       =   15
         TabIndex        =   48
         Text            =   "Text1 7"
         Top             =   1875
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Index           =   25
         Left            =   3450
         MaxLength       =   15
         TabIndex        =   47
         Text            =   "Text1 7"
         Top             =   2340
         Width           =   2325
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   26
         Left            =   4395
         MaxLength       =   15
         TabIndex        =   46
         Text            =   "Text1 7"
         Top             =   1875
         Width           =   1365
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   27
         Left            =   3825
         MaxLength       =   5
         TabIndex        =   45
         Text            =   "Text1 7"
         Top             =   1875
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   28
         Left            =   4395
         MaxLength       =   15
         TabIndex        =   44
         Text            =   "Text1 7"
         Top             =   1545
         Width           =   1365
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   29
         Left            =   3810
         MaxLength       =   5
         TabIndex        =   43
         Text            =   "Text1 7"
         Top             =   1545
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   30
         Left            =   4395
         MaxLength       =   15
         TabIndex        =   42
         Text            =   "Text1 7"
         Top             =   1230
         Width           =   1365
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   31
         Left            =   3810
         MaxLength       =   5
         TabIndex        =   41
         Text            =   "Text1 7"
         Top             =   1230
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   4440
         MaxLength       =   15
         TabIndex        =   40
         Text            =   "Text1 7"
         Top             =   570
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   12
         Left            =   4410
         MaxLength       =   15
         TabIndex        =   39
         Text            =   "Text1 7"
         Top             =   570
         Width           =   1335
      End
      Begin VB.Line Line1 
         X1              =   4410
         X2              =   5730
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Bruto"
         Height          =   255
         Index           =   7
         Left            =   4410
         TabIndex        =   70
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Cod."
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   69
         Top             =   990
         Width           =   405
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
         Height          =   255
         Index           =   13
         Left            =   555
         TabIndex        =   68
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Importe IVA"
         Height          =   255
         Index           =   33
         Left            =   2490
         TabIndex        =   67
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   37
         Left            =   4110
         TabIndex        =   66
         Top             =   600
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   36
         Left            =   11880
         TabIndex        =   65
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "TOTAL ALBARAN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   39
         Left            =   1620
         TabIndex        =   64
         Top             =   2400
         Width           =   1530
      End
      Begin VB.Label Label1 
         Caption         =   "% IVA"
         Height          =   255
         Index           =   41
         Left            =   1875
         TabIndex        =   63
         Top             =   1020
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "% Rec"
         Height          =   255
         Index           =   15
         Left            =   3810
         TabIndex        =   62
         Top             =   1020
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Recargo"
         Height          =   255
         Index           =   16
         Left            =   4395
         TabIndex        =   61
         Top             =   1020
         Width           =   1335
      End
   End
   Begin VB.Frame FrameAux1 
      Caption         =   "Variedades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3990
      Left            =   90
      TabIndex        =   19
      Top             =   3450
      Width           =   12195
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   300
         Index           =   0
         Left            =   3480
         MaskColor       =   &H00000000&
         TabIndex        =   31
         ToolTipText     =   "Buscar Variedad"
         Top             =   2250
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   3705
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   30
         Text            =   "Nombre variedad"
         Top             =   2250
         Width           =   1200
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   990
         MaxLength       =   12
         TabIndex        =   29
         Tag             =   "Num.Albaran|N|N|||rbodalbaran_variedad|numalbar|0000000|S|"
         Text            =   "NumAlb"
         Top             =   2250
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   5
         Left            =   2445
         MaxLength       =   16
         TabIndex        =   20
         Tag             =   "Variedad|T|N|||rbodalbaran_variedad|codvarie||N|"
         Text            =   "variedad"
         Top             =   2250
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   6
         Left            =   6255
         MaxLength       =   10
         TabIndex        =   22
         Tag             =   "Cantidad|N|N|||rbodalbaran_variedad|cantidad|###,##0.00||"
         Text            =   "cantidad"
         Top             =   2250
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   1740
         MaxLength       =   12
         TabIndex        =   28
         Tag             =   "Num.Linea|N|N|||rbodalbaran_variedad|numlinea|000|S|"
         Text            =   "Linea"
         Top             =   2250
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   7
         Left            =   7155
         MaxLength       =   12
         TabIndex        =   23
         Tag             =   "Precio|N|N|||rbodalbaran_variedad|precioar|###,##0.0000||"
         Text            =   "precio"
         Top             =   2250
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   8
         Left            =   8010
         MaxLength       =   5
         TabIndex        =   24
         Tag             =   "Dto.Linea|N|N|||rbodalbaran_variedad|dtolinea|#0.00||"
         Text            =   "Dto.Linea"
         Top             =   2250
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   9
         Left            =   8820
         MaxLength       =   12
         TabIndex        =   25
         Tag             =   "Importe|N|N|||rbodalbaran_variedad|importel|##,###,##0.00||"
         Text            =   "importe"
         Top             =   2250
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   10
         Left            =   9630
         MaxLength       =   2
         TabIndex        =   27
         Tag             =   "CodIva|N|N|||rbodalbaran_variedad|codigiva|00||"
         Text            =   "Codiva"
         Top             =   2250
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   16
         Left            =   1755
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   26
         Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
         Top             =   3540
         Width           =   8430
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   4
         Left            =   5520
         MaxLength       =   10
         TabIndex        =   21
         Tag             =   "Unidades|N|N|||rbodalbaran_variedad|unidades|###,##0.00||"
         Text            =   "Uds"
         Top             =   2250
         Visible         =   0   'False
         Width           =   420
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   1
         Left            =   240
         TabIndex        =   32
         Top             =   315
         Width           =   1110
         _ExtentX        =   1958
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
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "frmBodAlbRetirada.frx":0A0E
         Height          =   2640
         Left            =   240
         TabIndex        =   33
         Top             =   810
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   4657
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adoaux 
         Height          =   330
         Index           =   1
         Left            =   1455
         Top             =   315
         Visible         =   0   'False
         Width           =   1425
         _ExtentX        =   2514
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
      Begin VB.Label Label1 
         Caption         =   "Ampliación Línea"
         Height          =   255
         Index           =   35
         Left            =   405
         TabIndex        =   34
         Top             =   3585
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2685
      Left            =   90
      TabIndex        =   11
      Top             =   690
      Width           =   6225
      Begin VB.TextBox Text1 
         Height          =   990
         Index           =   2
         Left            =   225
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Tag             =   "Observaciones|T|S|||rbodalbaran|observac|||"
         Top             =   1560
         Width           =   5835
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   72
         Top             =   1770
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   2340
         MaxLength       =   10
         TabIndex        =   71
         Top             =   1770
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1155
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Albaran|F|N|||rbodalbaran|fechaalb|dd/mm/yyyy|N|"
         Top             =   390
         Width           =   1065
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   2010
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   750
         Width           =   4065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   1155
         MaxLength       =   6
         TabIndex        =   2
         Tag             =   "Cod. Socio|N|N|0|999999|rbodalbaran|codsocio|000000||"
         Text            =   "Text1"
         Top             =   750
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         Height          =   315
         Index           =   0
         Left            =   210
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "NºAlbarán|N|S|||rbodalbaran|numalbar|0000000|S|"
         Text            =   "Text1 7"
         Top             =   390
         Width           =   885
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   2940
         MaxLength       =   10
         TabIndex        =   35
         Top             =   1800
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   7
         Left            =   3330
         MaxLength       =   10
         TabIndex        =   36
         Top             =   1800
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   8
         Left            =   4500
         MaxLength       =   10
         TabIndex        =   37
         Top             =   1800
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha "
         Height          =   255
         Index           =   29
         Left            =   1170
         TabIndex        =   16
         Top             =   150
         Width           =   585
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1980
         Picture         =   "frmBodAlbRetirada.frx":0A23
         ToolTipText     =   "Buscar fecha"
         Top             =   120
         Width           =   240
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   1350
         ToolTipText     =   "Zoom descripción"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label29 
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   225
         TabIndex        =   14
         Top             =   1200
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Socio"
         Height          =   255
         Index           =   0
         Left            =   225
         TabIndex        =   13
         Top             =   795
         Width           =   540
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   885
         ToolTipText     =   "Buscar Socio"
         Top             =   795
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "NºAlbarán"
         Height          =   255
         Index           =   28
         Left            =   210
         TabIndex        =   12
         Top             =   150
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   7560
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
         Left            =   120
         TabIndex        =   8
         Top             =   180
         Width           =   1785
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   11085
      TabIndex        =   5
      Top             =   7650
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9870
      TabIndex        =   4
      Top             =   7650
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   12360
      _ExtentX        =   21802
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
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
            Object.ToolTipText     =   "Asignación Precios"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   10
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   11070
      TabIndex        =   6
      Top             =   7620
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   3000
      Top             =   1080
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   240
      Top             =   7650
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
      Left            =   240
      Top             =   7710
      Visible         =   0   'False
      Width           =   1425
      _ExtentX        =   2514
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
      Height          =   285
      Index           =   11
      Left            =   6705
      MaxLength       =   15
      TabIndex        =   17
      Text            =   "Text1 7"
      Top             =   1710
      Width           =   1485
   End
   Begin VB.Label Label1 
      Caption         =   "Imp.Descuento 2"
      Height          =   255
      Index           =   10
      Left            =   6705
      TabIndex        =   18
      Top             =   1440
      Width           =   1215
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
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         HelpContextID   =   2
         Shortcut        =   ^I
      End
      Begin VB.Menu mnAsignacion 
         Caption         =   "&Asignación Precios"
         Shortcut        =   ^A
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
Attribute VB_Name = "frmBodAlbRetirada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public Albaran As String  ' venimos de albaranes para ver las facturas donde aparece el albaran

'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmVar As frmComVar 'Form variedades comercial
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes ' ayuda para ver que se ha llevado el socio
Attribute frmMens.VB_VarHelpID = -1

Private WithEvents frmSoc As frmManSocios 'Form Mto de socios
Attribute frmSoc.VB_VarHelpID = -1

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

Dim TipoFactura As Byte
Dim cantidad As Currency
Private BuscaChekc As String

Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 0 'Variedades
            Set frmVar = New frmComVar
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.CodigoActual = txtAux(5).Text
            frmVar.Condicion = "productos.codgrupo in (5,6)"

            frmVar.Show vbModal
            Set frmVar = Nothing
    
            PonerFoco txtAux(5)
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1

End Sub

Private Sub Check1_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "check1(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "check1(" & Index & ")|"
    End If
End Sub

Private Sub chkAbonos_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim i As Integer

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda

        Case 3  'AÑADIR
            If DatosOk Then InsertarCabecera

        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaCabecera Then
                    espera 0.2
                    TerminaBloquear
                    PosicionarData
                    PonerCampos
                    PonerCamposLineas
                End If
            End If
            
         Case 5 'InsertarModificar LINEAS
            Select Case ModificaLineas
                Case 1 'afegir llínia
                    InsertarLinea
                Case 2 'modificar llínies
                    If ModificarLinea Then PosicionarData
            End Select
    End Select
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 0, 1, 3 'Busqueda, Insertar
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
            LLamaLineas Modo, 0, "DataGrid2"
            PonerFoco Text1(3)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(3)
            
        Case 5 'Lineas Detalle
            TerminaBloquear
            If ModificaLineas = 1 Then 'INSERTAR
                ModificaLineas = 0
                DataGrid3.AllowAddNew = False
                If Not Adoaux(1).Recordset.EOF Then Adoaux(1).Recordset.MoveFirst
            End If
            ModificaLineas = 0
            LLamaLineas Modo, 0, "DataGrid3"
            PonerModo 2
            DataGrid3.Enabled = True
            If Not Data1.Recordset.EOF Then _
                Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
            'Habilitar las opciones correctas del menu segun Modo
            PonerModoOpcionesMenu (Modo)
            PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
            DataGrid3.Enabled = True
            PonerFocoGrid DataGrid3
    
    End Select
End Sub
Private Sub BotonAnyadir()

    LimpiarCampos 'Huida els TextBox
    
'    TipoFactura = 1
    
    PonerModo 3
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la capçalera *******
'    Text1(0).Text = SugerirCodigoSiguienteStr("palets", "numpalet")
'    FormateaCampo Text1(0)
'    cmbAux(0).ListIndex = -1
    
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
        
    LimpiarDataGrids
    
    If vParamAplic.AlbRetiradaManual Then
        PonerFoco Text1(0)
    Else
        PonerFoco Text1(3) '*** 1r camp visible que siga PK ***
    End If
    
    ' *** si n'hi han camps de descripció a la capçalera ***
    'PosarDescripcions
End Sub


Private Sub BotonBuscar()
Dim anc As Single

    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PonerModo 1
        
'        'poner los txtaux para buscar por lineas de albaran
'        anc = DataGrid2.Top
'        If DataGrid2.Row < 0 Then
'            anc = anc + 440
'        Else
'            anc = anc + DataGrid2.RowTop(DataGrid2.Row) + 20
'        End If
'        LLamaLineas Modo, anc, "DataGrid2"
        
        
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
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
        LimpiarDataGrids
        CadenaConsulta = "Select rbodalbaran.* "
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
'    If Check1(0).Value = 1 Or Check1(1).Value = 1 Or Check1(2).Value = 1 Then
'        MsgBox "Esta factura no podemos modificarla", vbExclamation
'        TerminaBloquear
'        Exit Sub
'    End If
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFoco Text1(1) '*** 1r camp visible que siga PK ***
        
End Sub


Private Sub BotonModificarLinea(Index As Integer)
'Modificar una linea
Dim vWhere As String
Dim anc As Single
Dim J As Byte

    On Error GoTo eModificarLinea


    ModificaLineas = 2 'Modificar

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then  '1= Insertar
        TerminaBloquear
        Exit Sub
    End If
    
'--monica
'    If Data2.Recordset.EOF Then
'        TerminaBloquear
'        Exit Sub
'    End If
    
    
    PonerModo 5, Index
 

    
    vWhere = ObtenerWhereCP(False)
    vWhere = vWhere & " and numlinea=" & Adoaux(1).Recordset!numlinea
    If Not BloqueaRegistro("rbodalbaran_variedad", vWhere) Then
        TerminaBloquear
        Exit Sub
    End If
    If DataGrid3.Bookmark < DataGrid3.FirstRow Or DataGrid3.Bookmark > (DataGrid3.FirstRow + DataGrid3.VisibleRows - 1) Then
        J = DataGrid3.Bookmark - DataGrid3.FirstRow
        DataGrid3.Scroll 0, J
        DataGrid3.Refresh
    End If
    
'    anc = ObtenerAlto(Me.DataGrid1)
    anc = DataGrid3.Top
    If DataGrid3.Row < 0 Then
        anc = anc + 210
    Else
        anc = anc + DataGrid3.RowTop(DataGrid3.Row) + 10
    End If

    txtAux(1).Text = DataGrid3.Columns(0).Text ' albaran
    txtAux(3).Text = DataGrid3.Columns(1).Text ' linea
    txtAux(5).Text = DataGrid3.Columns(2).Text ' variedad
    Text2(0).Text = DataGrid3.Columns(3).Text ' nombre de la variedad
    txtAux(4).Text = DataGrid3.Columns(4).Text ' unidades
    txtAux(6).Text = DataGrid3.Columns(5).Text ' cantidad
    txtAux(7).Text = DataGrid3.Columns(6).Text ' precio
    txtAux(8).Text = DataGrid3.Columns(7).Text ' dtolinea
    txtAux(9).Text = DataGrid3.Columns(8).Text ' importe
    Text2(16).Text = DataGrid3.Columns(9).Text ' ampliacion
    txtAux(10).Text = DataGrid3.Columns(10).Text ' codigo de iva
    
    cantidad = DevuelveDesdeBDNew(cAgro, "variedades", "kgscajon", "codvarie", txtAux(5), "N")
    
    BloquearTxt txtAux(5), True
'    BloquearTxt txtAux(7), True
    BloquearTxt txtAux(6), True
    BloquearTxt txtAux(9), True
    txtAux(5).Enabled = False
'    txtAux(7).Enabled = False
    txtAux(6).Enabled = False
    txtAux(9).Enabled = False
    
    BloquearTxt txtAux(4), False
    BloquearTxt txtAux(8), False
    BloquearTxt txtAux(7), False
    
    BloquearBtn Me.btnBuscar(0), True
    
    LLamaLineas ModificaLineas, anc, "DataGrid3"
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
    DataGrid3.Enabled = True
    
    PonerFoco txtAux(4)
    Me.DataGrid3.Enabled = False


eModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub LLamaLineas(xModo As Byte, Optional alto As Single, Optional grid As String)
Dim jj As Integer
Dim b As Boolean
    
    Select Case grid
        Case "DataGrid3"
            DeseleccionaGrid Me.DataGrid3
            b = (xModo = 1 Or xModo = 2)
            For jj = 4 To 9
                txtAux(jj).Height = DataGrid3.RowHeight - 10
                txtAux(jj).Top = alto + 5
                txtAux(jj).visible = b
                txtAux(jj).Enabled = b
            Next jj
            
            txtAux(9).Enabled = False
            txtAux(6).Enabled = False
            
            Text2(0).Height = DataGrid3.RowHeight - 10
            Text2(0).Top = alto + 5
            Text2(0).visible = b
           
            btnBuscar(0).Height = DataGrid3.RowHeight - 10
            btnBuscar(0).Top = alto + 5
            btnBuscar(0).visible = b
            
    End Select
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
    
'    'solo se puede modificar la factura si no esta contabilizada
'    If FactContabilizada Then Exit Sub
    
    Cad = "Cabecera de Albaranes." & vbCrLf
    Cad = Cad & "-------------------------------------      " & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar el Albarán:            "
    Cad = Cad & vbCrLf & "Nº Albarán:  " & Format(Text1(0).Text, "0000000")
    Cad = Cad & vbCrLf & "Fecha:  " & Format(Text1(1).Text, "dd/mm/yyyy")

    Cad = Cad & vbCrLf & vbCrLf & " ¿Desea Eliminarlo? "

    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
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
            LimpiarDataGrids
            PonerModo 0
        End If
        
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEliminar:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminar Albaran", Err.Description
End Sub

Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim Cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        DataGrid3.Enabled = True
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


Private Sub DataGrid3_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Adoaux(1).Recordset.EOF Then
        Text2(16).Text = DBLet(Adoaux(1).Recordset!ampliaci, "T")
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Load()
Dim i As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
     'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next kCampo

    ' ICONITOS DE LA BARRA
    btnPrimero = 14
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(4).Image = 3   'Insertar
        .Buttons(5).Image = 4   'Modificar
        .Buttons(6).Image = 5   'Borrar
        .Buttons(8).Image = 10  'Impresión de albaran
        .Buttons(10).Image = 13  'Asignación masiva de precios
        .Buttons(11).Image = 11  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    
    ' ******* si n'hi han llínies *******
    'ICONETS DE LES BARRES ALS TABS DE LLÍNIA
    For kCampo = 1 To 1
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
   'IMAGES para zoom
    For i = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(i).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next i
    
    LimpiarCampos   'Limpia los campos TextBox

    CodTipoMov = "ALB"
    VieneDeBuscar = False
    
        
    '## A mano
    NombreTabla = "rbodalbaran"
    NomTablaLineas = "rbodalbaran_variedad" 'Tabla lineas de lineas de albaranes de variedades
    Ordenacion = " ORDER BY rbodalbaran.numalbar"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    CadenaConsulta = "select * from rbodalbaran "
    If Albaran <> "" Then
        CadenaConsulta = CadenaConsulta & " where numalbar = " & Albaran
    Else
        CadenaConsulta = CadenaConsulta & " where numalbar = -1"
    End If
    
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
'        CargaGrid DataGrid1, Data2, False
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
'    SSTab1.Tab = 0
    
'    If DatosADevolverBusqueda <> "" Then
'        Text1(0).Text = DatosADevolverBusqueda
'        HacerBusqueda
'    Else
'        PonerModo 0
'    End If
    
    If DatosADevolverBusqueda = "" Then
        If Albaran = "" Then
            PonerModo 0
        Else
            HacerBusqueda
'            SSTab1.Tab = 0
        End If
    Else
        BotonBuscar
    End If
    
    PrimeraVez = False
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Modo = 4 Then TerminaBloquear
End Sub

Private Sub frmAlb_DatoSeleccionado(CadenaSeleccion As String)
'    txtAux3(4).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000") 'Numero albaran
'    txtAux3(5).Text = Format(RecuperaValor(CadenaSeleccion, 2), "00") 'Numero linea
'    CargarDatosAlbaran txtAux3(4).Text, txtAux3(5).Text
End Sub


Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Articulos
    txtAux(5).Text = RecuperaValor(CadenaSeleccion, 1) 'codartic
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
    If txtAux(5) <> "" Then
        txtAux(7) = DevuelveDesdeBDNew(cAgro, "sartic", "preciove", "codartic", txtAux(5), "T")
        txtAux(10) = DevuelveDesdeBDNew(cAgro, "sartic", "codigiva", "codartic", txtAux(5), "T")
    End If
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = Aux
        
        
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    If imgFec(0).Tag < 2 Then
        Text1(CByte(imgFec(0).Tag) + 1).Text = Format(vFecha, "dd/mm/yyyy") '<===
    Else
        Text1(CByte(imgFec(0).Tag) + 8).Text = Format(vFecha, "dd/mm/yyyy") '<===
    End If
    ' ********************************************
End Sub


Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000") 'Cod Socio
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre del socio
    PonerFoco Text1(indice)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(5).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
End Sub

Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. de Socio
            indice = 3
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

    If Index < 2 Then
        imgFec(0).Tag = Index '<===
        ' *** repasar si el camp es txtAux o Text1 ***
        If Text1(Index + 1).Text <> "" Then frmC.NovaData = Text1(Index + 1).Text
    Else
        imgFec(0).Tag = Index '<===
        ' *** repasar si el camp es txtAux o Text1 ***
        If Text1(Index + 8).Text <> "" Then frmC.NovaData = Text1(Index + 8).Text
    End If
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    If Index < 2 Then
        PonerFoco Text1(CByte(imgFec(0).Tag) + 1) '<===
    Else
        PonerFoco Text1(CByte(imgFec(0).Tag) + 8) '<===
    End If
    ' ********************************************
End Sub


Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        indice = 2
        frmZ.pTitulo = "Observaciones del Albarán"
        frmZ.pValor = Text1(indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco Text1(indice)
    End If
End Sub

Private Sub mnAsignacion_Click()
    AbrirListadoBodEntradas (3)
    
    TerminaBloquear
    If Not Data1.Recordset.EOF Then
        PosicionarData
        PonerCampos
        PonerCamposLineas
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
    If Modo = 5 Then 'Modificar lineas
        'bloquea la tabla cabecera de albaranes: scaalb
        If BLOQUEADesdeFormulario(Me) Then
            If BloqueaLineasAlb Then BotonModificarLinea (1)
        End If
         
    Else   'Modificar albaran
        'bloquea la tabla cabecera de albaranes: scaalb
        If BLOQUEADesdeFormulario(Me) Then
            BotonModificar
        End If
    End If
End Sub


Private Function BloqueaLineasAlb() As Boolean
'bloquea todas las lineas de la factura
Dim Sql As String

    On Error GoTo EBloqueaLin

    BloqueaLineasAlb = False
    'bloquear cabecera albaranes
    Sql = "select * FROM slialb "
    Sql = Sql & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute Sql, , adCmdText
    BloqueaLineasAlb = True

EBloqueaLin:
    If Err.Number <> 0 Then BloqueaLineasAlb = False
End Function

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
    If Index = 9 Then HaCambiadoCP = True 'Cod. Postal
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
    If Index <> 2 Or (Index = 2 And Text1(2).Text = "") Then KEYpress KeyAscii
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
Dim Nregs As Long

        
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 0
            If Modo = 3 And Text1(0).Text = "" And vParamAplic.AlbRetiradaManual Then
                MsgBox "Debe introducir un número de albarán.", vbExclamation
                PonerFoco Text1(Index)
            End If
        Case 1 'Fecha albaran
            '[Monica]28/08/2013: comprobamos que la fecha esté en la campaña
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index), True
            
        Case 3 'Socio
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 1 Then
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rsocios", "nomsocio")
                Else
                    PonerDatosSocio (Text1(Index).Text)
                    If Text2(Index).Text = "" Then
'                        cadMen = "No existe el Socio: " & Text1(Index).Text & vbCrLf
'                        cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
'                        If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
'                            Set frmSoc = New frmManSocios
'                            frmSoc.DatosADevolverBusqueda = "0|1|"
'                            Text1(Index).Text = ""
'                            TerminaBloquear
'                            frmSoc.Show vbModal
'                            Set frmSoc = Nothing
'                            If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
'                        Else
'                            Text1(Index).Text = ""
'                        End If
                        PonerFoco Text1(Index)
                    Else
                        If Modo = 3 Or Modo = 4 Then
                            Set frmMens = New frmMensajes
                            frmMens.cadWHERE = "and rbodalbaran.codsocio = " & Text1(3).Text
                            frmMens.OpcionMensaje = 5
                            frmMens.Show vbModal
                            
                            Set frmMens = Nothing
                        End If
                    End If
               End If
            End If
                
            
        Case 4 ' Forma de Pago
            
         Case 6 ' destino
         
         Case 7, 8 'descuentos
            If Modo = 1 Then Exit Sub
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 4
            
'         Case 5 ' importe de descuento
'            If Modo = 1 Then Exit Sub
'            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 3
    

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
    
'--monica
'    CadB = ObtenerBusqueda(Me)
'++monica
    If Albaran = "" Then
        CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    Else
        CadB = "numalbar = " & Albaran & " "
    End If

    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select rbodalbaran.* from " & NombreTabla
        CadenaConsulta = CadenaConsulta & " WHERE " & CadB & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim Tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
    'Llamamos a al form
    '##A mano
    Cad = ""
    Cad = Cad & "Nº.Albarán|rbodalbaran.numalbar|N||15·"
    Cad = Cad & "Socio|rbodalbaran.codsocio|N||10·" 'ParaGrid(Text1(3), 10, "Socios")
    Cad = Cad & "Nombre Socio|rsocios.nomsocio|N||60·"
    Cad = Cad & ParaGrid(Text1(1), 15, "F.Albarán")
    Tabla = NombreTabla & " INNER JOIN rsocios ON rbodalbaran.codsocio=rsocios.codsocio "
    
    Titulo = "Albaranes de Retirada Vino y Aceite"
    devuelve = "0|"
           
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vtabla = Tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|4|"
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
        '--monica
        'LLamaLineas Modo, 0, "DataGrid2"
        PonerCampos
    End If


    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCamposLineas()
'Carga el grid de los AlbaranesxFactura, es decir, la tabla scafac1 de la factura seleccionada
Dim b As Boolean
Dim b2 As Boolean
Dim i As Integer

    On Error GoTo EPonerLineas

    If Data1.Recordset.EOF Then Exit Sub

    Screen.MousePointer = vbHourglass
    
    If Data1.Recordset.RecordCount > 0 Then
        CargaGrid DataGrid3, Adoaux(1), True
    Else
        CargaGrid DataGrid3, Adoaux(1), False
    End If
    If Not Adoaux(1).Recordset.EOF Then
        Text2(16).Text = DBLet(Adoaux(1).Recordset!ampliaci, "T")
    Else
        Text2(16).Text = ""
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim BrutoFac As Single
Dim b As Boolean

    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    
    b = PonerCamposForma2(Me, Data1, 2, "Frame2")
    
    'poner descripcion campos
    Modo = 4
    
    
'    PosicionarCombo Combo1(0), DevuelveDesdeBDNew(cAgro, "clientes", "tipoiva", "codclien", Text1(3).Text, "N")
    
    Text2(3).Text = PonerNombreDeCod(Text1(3), "rsocios", "nomsocio", "codsocio", "N") 'socio
    
'    Text2(18).Text = PonerNombreDeCod(Text1(16), "salmpr", "nomalmac", "codalmac", "N") 'almacen
    
    Modo = 2
    
    CalcularDatosAlbaran
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    
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
Dim i As Byte, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar
    '## No tiene el boton modificar y no utiliza la funcion general
'    ActualizarToolbar Modo, Kmodo
    BuscaChekc = ""
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Or Albaran <> "" Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
        
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
          
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    
    For i = 9 To 31
        BloquearTxt Text1(i), Not (Modo = 1)
        Text1(i).Enabled = (Modo = 1)
    Next i
    
    b = (Modo <> 1)
    
    'Campos Nº Albarán bloqueado y en azul
    BloquearTxt Text1(0), b, True
    
    If vParamAplic.AlbRetiradaManual And Modo = 3 Then
        Text1(0).Locked = False
        Text1(0).BackColor = vbWhite
    End If
    
    
    Me.imgZoom(0).Enabled = Not (Modo = 0)
    
    'Si no es modo lineas Boquear los TxtAux
    For i = 1 To 1
        txtAux(i).visible = False
        BloquearTxt txtAux(i), True
    Next i
    For i = 3 To 10
        txtAux(i).visible = False
        BloquearTxt txtAux(i), True
    Next i
    
    For i = 0 To 0
        Text2(i).visible = ((Modo = 5) And (indFrame = 1))
        Text2(i).Enabled = False
    Next i
    
    BloquearTxt Text2(16), (Modo <> 5)
    
    BloquearBtn Me.btnBuscar(0), True
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    BloquearImgBuscar Me, Modo, ModificaLineas
    BloquearImgFec Me, 0, Modo
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    ' *** si n'hi han llínies i algún tab que no te datagrid ***
    BloquearFrameAux Me, "FrameAux1", Modo, 1
    
        
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
Dim Serie As String
Dim Sql As String

    On Error GoTo EDatosOK

    DatosOk = False
    
'    ComprobarDatosTotales

    'comprobamos datos OK de la tabla scaalb
    b = CompForm2(Me, 2, "Frame2") ' , 1) 'Comprobar formato datos ok de la cabecera: opcion=1
    If Not b Then Exit Function
    
    If b And Modo = 3 And vParamAplic.AlbRetiradaManual Then
        If ExisteCP(Text1(0)) Then b = False
    End If
    
    
    DatosOk = b
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
Dim b As Boolean
Dim i As Byte

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True

    For i = 4 To 7
        If i = 4 Or i = 6 Or i = 7 Then
            If txtAux(i).Text = "" Then
                MsgBox "El campo " & txtAux(i).Tag & " no puede ser nulo", vbExclamation
                b = False
                PonerFoco txtAux(i)
                Exit Function
            End If
        End If
    Next i
            
    DatosOkLinea = b
    
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 16 And KeyCode = 40 Then 'campo Amliacion Linea y Flecha hacia abajo
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 16 And KeyAscii = 13 Then 'campo Amliacion Linea y ENTER
        PonerFocoBtn Me.cmdAceptar
    End If
End Sub

Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)

    If BloqueaRegistro(NombreTabla, "numalbar = " & Data1.Recordset!numalbar) Then
'    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
        Select Case Button.Index
            Case 1
                BotonAnyadirLinea Index
            Case 2
                BotonModificarLinea Index
            Case 3
                BotonEliminarLinea Index
            Case Else
        End Select
    End If

End Sub


Private Sub BotonEliminarLinea(Index As Integer)
Dim Cad As String
Dim Sql As String
Dim Mens As String
Dim b As Boolean

    On Error GoTo EEliminarLinea

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    ' ***************************************************************************

    b = True

    ' *************** canviar la pregunta ****************
    Cad = "¿Seguro que desea eliminar la Variedad?"
    Cad = Cad & vbCrLf & "Albarán: " & Adoaux(1).Recordset.Fields(0)
    Cad = Cad & vbCrLf & "Variedad: " & Adoaux(1).Recordset.Fields(3) & " - " & Adoaux(1).Recordset.Fields(4)
    
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminarLinea
        Screen.MousePointer = vbHourglass
        NumRegElim = Adoaux(1).Recordset.AbsolutePosition
        
        If Not EliminarLinea Then
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            CalcularDatosAlbaran
            If SituarDataTrasEliminar(Adoaux(1), NumRegElim) Then
                PonerCampos
            Else
                PonerCampos
'                        LimpiarCampos
'                        PonerModo 0
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
       
EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Or Not b Then MuestraError Err.Number, "Eliminar Linea de Factura", Err.Description

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
        Case 10 ' Asignacion de precios masiva
            mnAsignacion_Click
        Case 11    'Salir
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


Private Sub ActualizarToolbar(Modo As Byte, Kmodo As Byte)
'Modo: Modo antiguo
'Kmodo: Modo que se va a poner

    If (Modo = 5) And (Kmodo <> 5) Then
        'El modo antigu era modificando las lineas
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
'        Toolbar1.Buttons(5).Image = 3
'        Toolbar1.Buttons(5).ToolTipText = "Nuevo Albaran"
        '-- Modificar
        Toolbar1.Buttons(5).Image = 4
        Toolbar1.Buttons(5).ToolTipText = "Modificar Factura"
        '-- eliminar
        Toolbar1.Buttons(6).Image = 5
        Toolbar1.Buttons(6).ToolTipText = "Eliminar Factura"
    End If
    If Kmodo = 5 Then
        'Ponemos nuevos dibujitos y tal y tal
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
'        Toolbar1.Buttons(5).Image = 12
'        Toolbar1.Buttons(5).ToolTipText = "Nueva linea"
        '-- Modificar
        Toolbar1.Buttons(5).Image = 13
        Toolbar1.Buttons(5).ToolTipText = "Modificar linea factura"
        '-- eliminar
        Toolbar1.Buttons(6).Image = 14
        Toolbar1.Buttons(6).ToolTipText = "Eliminar linea factura"
    End If
End Sub
    


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim b As Boolean
Dim Opcion As Byte
Dim Sql As String

    On Error GoTo ECargaGRid

    b = DataGrid3.Enabled
    Select Case vDataGrid.Name
        Case "DataGrid3" 'envases
            Opcion = 1
    End Select
    
    Sql = MontaSQLCarga(enlaza, Opcion)
    CargaGridGnral vDataGrid, vData, Sql, PrimeraVez
    
    vDataGrid.RowHeight = 270
    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
    
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
    vDataGrid.Enabled = Not b
    
    Exit Sub
    
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim tots As String
    
    On Error GoTo ECargaGRid

    Select Case vDataGrid.Name
         Case "DataGrid3" 'rbodalbaran_variedad lineas de albaran de bodega
'select codtipom,numfactu,fecfactu,numlinea,codartic,sartic.nomartic,cantidad,precioar,dtolinea,importel,ampliaci,codigiva
            tots = "N||||0|;N||||0|;"
            tots = tots & "S|txtAux(5)|T|Variedad|1500|;S|btnBuscar(0)|B|||;"
            tots = tots & "S|Text2(0)|T|Nombre|3200|;S|txtAux(4)|T|Unidades|1200|;S|txtAux(6)|T|Cantidad|1200|;"
            tots = tots & "S|txtAux(7)|T|Precio|1200|;S|txtAux(8)|T|Dto|800|;S|txtAux(9)|T|Importe|1500|;N||||0|;N||||0|;"
            arregla tots, DataGrid3, Me
     
    End Select
    
    vDataGrid.HoldFields
    Exit Sub
    
ECargaGRid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
  
    ConseguirFocoLin txtAux(Index)
End Sub



Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Sql As String
Dim devuelve As String
Dim b As Boolean
Dim TipoDto As Byte


    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    Select Case Index
        Case 5 'variedad
            If txtAux(Index).Text = "" Then
                Exit Sub
            End If
        
            Text2(0).Text = PonerNombreDeCod(txtAux(5), "variedades", "nomvarie", "codvarie", "N")
            If Text2(0).Text = "" Then
                MsgBox "Variedad no existe. Reintroduzca.", vbExclamation
                PonerFoco txtAux(Index)
            Else
                If Not (EsVariedadGrupo5(txtAux(5)) Or EsVariedadGrupo6(txtAux(5))) Then
                    MsgBox "Esta variedad no es del grupo de Bodega ni de Almazara. Reintroduzca.", vbExclamation
                    PonerFoco txtAux(Index)
                Else
                    'precio de venta
                    If txtAux(7).Text = "" Then
                        txtAux(7) = DevuelveDesdeBDNew(cAgro, "variedades", "minkgcaj", "codvarie", txtAux(5), "N")
                    End If
                    'iva
                    txtAux(10).Text = DevuelveDesdeBDNew(cAgro, "variedades", "codigiva", "codvarie", txtAux(5), "N")
                    'unidades
                    cantidad = 0
                    Sql = DevuelveDesdeBDNew(cAgro, "variedades", "kgscajon", "codvarie", txtAux(5), "N")
                    If Sql <> "" Then cantidad = CCur(Sql)
                    PonerFoco txtAux(4)
                End If
            End If
            
        
        Case 4 'unidades
            If txtAux(Index).Text <> "" Then
                PonerFormatoDecimal txtAux(Index), 1
                txtAux(6).Text = Round2(cantidad * ImporteSinFormato(txtAux(Index).Text), 2)
                PonerFormatoDecimal txtAux(6), 1
            End If
        
'        Case 6 ' Cantidad
'            PonerFormatoDecimal txtAux(Index), 1
'
        Case 7 ' Precio
            PonerFormatoDecimal txtAux(Index), 11 '[Monica] 29/10/2010 antes era 1 (con 2 decimales)
            
        Case 8  'Descuentos
            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
            
        Case 9 'Importe Linea
            PonerFormatoDecimal txtAux(Index), 3 'Tipo 3: Decimal(10,2)
            
    End Select
     If (Index = 4 Or Index = 7 Or Index = 8 Or Index = 9) Then 'Cant., Precio, Dto1, Dto2
'        txtAux(6).Text = Round2(Cantidad * ImporteSinFormato(txtAux(4).Text), 2)
        If txtAux(8).Text = "" Then txtAux(8).Text = 0

'        TipoDto = DevuelveDesdeBDNew(cAgro, "clientes", "tipodtos", "codclien", Text1(3).Text, "N")
        TipoDto = 0
        txtAux(9).Text = CalcularImporte(txtAux(6).Text, txtAux(7).Text, txtAux(8).Text, 0, TipoDto, 0)
        PonerFormatoDecimal txtAux(9), 3
    End If
    
End Sub

Private Function Eliminar() As Boolean
Dim Sql As String, LEtra As String, Sql2 As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String

    On Error GoTo FinEliminar

    If Data1.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    'Eliminar en tablas de cabecera de albaran
    '------------------------------------------
    Sql = " " & ObtenerWhereCP(True)
    
    'Lineas de variedades (rbodalbaran_variedad)
    conn.Execute "Delete from rbodalbaran_variedad " & Sql
    
    'Cabecera de factura
    conn.Execute "Delete from " & NombreTabla & Sql
    
    'Decrementar contador si borramos el ult. albaran de retirada de bodega
    If Not vParamAplic.AlbRetiradaManual Then
        Set vTipoMov = New CTiposMov
        vTipoMov.DevolverContador CodTipoMov, Val(Text1(0).Text)
        Set vTipoMov = Nothing
    End If
    
    
    b = True
FinEliminar:
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, "Eliminar Albarán de Retirada de Bodega", Err.Description & " " & Mens
        b = False
    End If
    If Not b Then
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function

Private Function EliminarLinea() As Boolean
Dim Sql As String, LEtra As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String

    On Error GoTo FinEliminar

    If Adoaux(1).Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    Mens = ""
    
    'Eliminar en tablas de slialb
    '------------------------------------------
    Sql = " where numalbar = " & Adoaux(1).Recordset.Fields(0)
    Sql = Sql & " and numlinea = " & Adoaux(1).Recordset.Fields(1)

    'Lineas de variedades
    conn.Execute "Delete from rbodalbaran_variedad " & Sql
    
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Variedades del Albarán de Retirada", Err.Description & " " & Mens
        conn.RollbackTrans
        EliminarLinea = False
    Else
        conn.CommitTrans
        EliminarLinea = True
    End If
End Function

Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next

    CargaGrid DataGrid3, Me.Adoaux(1), False 'variedades
    
    If Err.Number <> 0 Then Err.Clear
End Sub


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
             LimpiarDataGrids
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
    
    Sql = "numalbar= " & DBSet(Text1(0).Text, "N")
    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function


Private Function MontaSQLCarga(enlaza As Boolean, Opcion As Byte) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim Sql As String
    
    Select Case Opcion
        Case 1  'variedades
'select codtipom,numfactu,fecfactu,numlinea,codalmac,codartic,sartic.nomartic,cantidad,precioar,dtolinea,importel,ampliaci,codigiva
            Sql = "SELECT rbodalbaran_variedad.numalbar,rbodalbaran_variedad.numlinea,rbodalbaran_variedad.codvarie,variedades.nomvarie,rbodalbaran_variedad.unidades, rbodalbaran_variedad.cantidad,"
            Sql = Sql & "precioar,dtolinea,importel,ampliaci,rbodalbaran_variedad.codigiva"
            Sql = Sql & " FROM rbodalbaran_variedad, variedades "
            Sql = Sql & " WHERE rbodalbaran_variedad.codvarie = variedades.codvarie "
    End Select
    
    If enlaza Then
        Sql = Sql & " and " & ObtenerWhereCP(False)
    Else
        Sql = Sql & " and numalbar = -1"
    End If
    Sql = Sql & " ORDER BY numalbar,numlinea"
    MontaSQLCarga = Sql
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean, bAux As Boolean
Dim i As Integer

        b = ((Modo = 2) Or (Modo = 0)) And (Albaran = "") 'Or (Modo = 5 And ModificaLineas = 0)
        'Buscar
        Toolbar1.Buttons(1).Enabled = b
        Me.mnBuscar.Enabled = b
        'Vore Tots
        Toolbar1.Buttons(2).Enabled = b
        Me.mnVerTodos.Enabled = b
        'Añadir
        Toolbar1.Buttons(4).Enabled = b
        Me.mnModificar.Enabled = b
        
        
        b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And (Albaran = "")
        'Modificar
        Toolbar1.Buttons(5).Enabled = b
        Me.mnModificar.Enabled = b
        'eliminar
        Toolbar1.Buttons(6).Enabled = b
        Me.mnEliminar.Enabled = b
        'Impresión de factura
        Toolbar1.Buttons(8).Enabled = (Modo = 2) Or (Albaran <> "")
        Me.mnImprimir.Enabled = (Modo = 2) Or (Albaran <> "")
'        'Orden de Carga
'        Toolbar1.Buttons(9).Enabled = b
'        Me.mnOrdenCarga.Enabled = b
'        'Generar CMR
'        Toolbar1.Buttons(10).Enabled = b
'        Me.mnCMR.Enabled = b
        

    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
'++monica: si insertamos lo he quitado
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    b = (Modo = 2) And (Albaran = "")
    For i = 1 To 1
        ToolAux(i).Buttons(1).Enabled = b
        
        If b Then
            Select Case i
              Case 0
                bAux = (b And Me.Adoaux(0).Recordset.RecordCount > 0)
              Case 1
                bAux = (b And Me.Adoaux(1).Recordset.RecordCount > 0)
            End Select
        End If
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i


End Sub


Private Sub BotonImprimir()
Dim cadFormula As String
Dim CadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

    If Text1(0).Text = "" Then
        MsgBox "Debe seleccionar un Albarán para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    CadParam = ""
    cadSelect = ""
    numParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    indRPT = 34 'Impresion de albaran de retirada de bodega
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
      
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de factura
    '---------------------------------------------------
    If Text1(0).Text <> "" Then
        'Nº Albaran
        devuelve = "{" & NombreTabla & ".numalbar}=" & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        devuelve = "numalbar = " & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
    End If
    
    If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub
     
     With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = CadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 0
            .Titulo = "Impresión de Albarán Retirada de Bodega"
            .ConSubInforme = False
            .NroCopias = 2
            .Show vbModal
    End With
End Sub


Private Function ModificaCabecera() As Boolean
Dim b As Boolean
Dim MenError As String

    On Error GoTo EModificarCab

    conn.BeginTrans
    
    b = ModificaDesdeFormulario2(Me, 2, "Frame2")
    
    MenError = "Recalculando Importes Netos de lineas"
    CalcularDatosAlbaran

EModificarCab:
    If Err.Number <> 0 Then
        MenError = "Modificando Albaran Envases." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        b = False
    End If
    If b Then
        ModificaCabecera = True
        conn.CommitTrans
    Else
        ModificaCabecera = False
        conn.RollbackTrans
    End If
End Function


Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String

    On Error GoTo EInsertarCab
    
'    CodTipoMov = Text1(6).Text
    
'    If TipoFactura = 0 Then
    If Not vParamAplic.AlbRetiradaManual Then
        Set vTipoMov = New CTiposMov
        If vTipoMov.Leer(CodTipoMov) Then
            Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
            Sql = CadenaInsertarDesdeForm(Me)
            If Sql <> "" Then
                If InsertarOferta(Sql, vTipoMov) Then
                    CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                    PonerCadenaBusqueda
                    PonerModo 2
                    'Ponerse en Modo Insertar Lineas
    '                BotonMtoLineas 0, "Variedades"
                    BotonAnyadirLinea 0
                End If
            End If
            Text1(0).Text = Format(Text1(0).Text, "0000000")
        End If
        Set vTipoMov = Nothing
    Else
            Sql = CadenaInsertarDesdeForm(Me)
            conn.Execute Sql

            CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
            PonerCadenaBusqueda
            PonerModo 2
            'Ponerse en Modo Insertar Lineas
'                BotonMtoLineas 0, "Variedades"
            BotonAnyadirLinea 0
            Text1(0).Text = Format(Text1(0).Text, "0000000")

    End If
    
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
    'Comprobar si mientras tanto se incremento el contador de Albaranes
    'para ello vemos si existe una factura con ese contador y si existe la incrementamos
    Do
        devuelve = DevuelveDesdeBDNew(cAgro, NombreTabla, "numalbar", "numalbar", Text1(0), "N")
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
    MenError = "Error al insertar en la tabla Cabecera de Albaranes (" & NombreTabla & ")."
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


'Private Sub CargaForaGrid()
'    If DataGrid2.Columns.Count <= 2 Then Exit Sub
'    ' *** posar als camps de fora del grid el valor de la columna corresponent ***
'    Text3(0) = DataGrid2.Columns(12).Text    'Fecha
'    Text3(1) = DataGrid2.Columns(13).Text    'Matricula
'    Text3(2) = DataGrid2.Columns(14).Text    'Destino
'    Text3(3) = DataGrid2.Columns(15).Text   'Variedad
'    Text3(4) = DataGrid2.Columns(16).Text   'Confeccion
'    ' *** Si fora del grid n'hi han camps de descripció, posar-los valor ***
'    ' **********************************************************************
'End Sub

Private Sub InsertarLinea()
'Inserta registre en les taules de Llínies
Dim nomframe As String
Dim b As Boolean
Dim Mens As String

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
'        Case 0: nomFrame = "FrameAux0" 'variedades
    nomframe = "FrameAux1" 'envases
    ' ***************************************************************
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        
        If InsertarLineaEnv(txtAux(3).Text) Then
            CalcularDatosAlbaran
            b = BloqueaRegistro("scaalb", "numalbar = " & Data1.Recordset!numalbar)
            CargaGrid DataGrid3, Adoaux(1), True
            If b Then BotonAnyadirLinea 1
        End If
    End If
End Sub

Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vtabla As String
Dim anc As Single
Dim i As Integer
    
    ModificaLineas = 1 'Posem Modo Afegir Llínia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModificaLineas = 0 Then Exit Sub
    End If
       
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    BloquearTxt Text1(6), True
    BloquearTxt Text1(1), True
    
    
    ' **************************************************

    ' *** posar el nom del les distintes taules de llínies ***
    vtabla = "rbodalbaran_variedad"
    ' ********************************************************
    
    vWhere = ObtenerWhereCab(False)
    
    NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)
    ' ***************************************************************

    AnyadirLinea DataGrid3, Adoaux(1)

    anc = DataGrid3.Top
    If DataGrid3.Row < 0 Then
        anc = anc + 215 '210
    Else
        anc = anc + DataGrid3.RowTop(DataGrid3.Row) + 5
    End If
  
    LLamaLineas ModificaLineas, anc, "DataGrid3"

    LimpiarCamposLin "FrameAux1"
'            txtAux(0).Text = Text1(6).Text 'codtipom
    txtAux(1).Text = Text1(0).Text 'numalbaran
'            txtAux(2).Text = Text1(1).Text 'fecfactu
    txtAux(3).Text = NumF
    PonerFoco txtAux(5)
    For i = 0 To 0
        Text2(i).Text = ""
    Next i
    txtAux(10).Enabled = False
    txtAux(10).visible = False
    BloquearTxt txtAux(6), True
    BloquearTxt txtAux(9), True
    BloquearTxt Text2(16), False
    BloquearBtn Me.btnBuscar(0), False
' ******************************************
End Sub

Private Sub LimpiarCamposLin(FrameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, FrameAux  'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Integer
Dim Cad As String
Dim Sql As String
Dim b As Boolean
Dim Mens As String
    
    On Error GoTo eModificarLinea

    ModificarLinea = False
    Sql = ""

    ' *** posa els noms del frames, tant si son de grid com si no ***
    nomframe = "FrameAux1" 'envases
    ' **************************************************************

    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        
        
        If DatosOkLineaEnv() Then
            '#### LAURA 15/11/2006
'            conn.BeginTrans
           
           'actualizar la linea de Albaran
            Sql = "UPDATE rbodalbaran_variedad Set unidades = " & DBSet(txtAux(4).Text, "N") & ", codvarie=" & DBSet(txtAux(5).Text, "N") & ", "
            Sql = Sql & "ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
            Sql = Sql & "cantidad= " & DBSet(txtAux(6).Text, "N") & ", "
            Sql = Sql & "precioar= " & DBSet(txtAux(7).Text, "N") & ", " 'precio
            Sql = Sql & "dtolinea= " & DBSet(txtAux(8).Text, "N") & ", "
            Sql = Sql & "importel= " & DBSet(txtAux(9).Text, "N") & ", " 'Importe
            Sql = Sql & "codigiva= " & DBSet(txtAux(10).Text, "N") & " " 'codigo de iva
            Sql = Sql & Replace(ObtenerWhereCP(True), NombreTabla, "rbodalbaran_variedad") & " AND numlinea=" & Adoaux(1).Recordset!numlinea
            conn.Execute Sql
        End If
            
        ModificaLineas = 0
        
        CalcularDatosAlbaran
        
        V = Adoaux(1).Recordset.Fields(1) 'el 2 es el nº de llinia
        CargaGrid DataGrid3, Adoaux(1), True

        ' *** si n'hi han tabs ***
'            SSTab1.Tab = 1

        DataGrid3.SetFocus
        Adoaux(1).Recordset.Find (Adoaux(1).Recordset.Fields(1).Name & " =" & V)

        LLamaLineas ModificaLineas, 0, "DataGrid3"

    End If
        
eModificarLinea:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Modificar Lineas Albarán Bodega" & vbCrLf & Err.Description & vbCrLf & Mens
'        conn.RollbackTrans
        ModificarLinea = False
    Else
'        conn.CommitTrans
        ModificarLinea = True
    End If
End Function
        

Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim b As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte
Dim Cliente As String

    DatosOkLlin = True
    
    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False

    b = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not b Then Exit Function
    
    'en variedades comprobamos que el albaran introducido corresponde al cliente
'    If nomFrame = "FrameAux0" Then
'        cliente = ""
'        cliente = DevuelveDesdeBDNew(cAgro, "albaran", "codclien", "numalbar", txtAux3(4).Text, "N")
'
'        If CLng(cliente) <> CLng(Data1.Recordset!CodClien) Then
'            MsgBox "El albarán introducido no es del cliente del la factura. Revise.", vbExclamation
'            b = False
'        End If
'    End If
    
    DatosOkLlin = b

EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " numalbar= " & Val(Text1(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function

'' ***** si n'hi han varios nivells de tabs *****
'Private Sub SituarTab(numTab As Integer)
'    On Error Resume Next
'
'    SSTab1.Tab = numTab
'
'    If Err.Number <> 0 Then Err.Clear
'End Sub

' **********************************************
Private Sub PonerDatosSocio(Codsocio As String, Optional nifSocio As String)
Dim vSocio As cSocio
Dim Observaciones As String
Dim Situacion As Integer
Dim NSitua As String
    
    On Error GoTo EPonerDatos
    
    If Codsocio = "" Then
        LimpiarDatosSocio
        Exit Sub
    End If

    Set vSocio = New cSocio
    
    'si se ha modificado el socio volver a cargar los datos
    If vSocio.Existe(Codsocio) Then
        If vSocio.LeerDatos(Codsocio) Then
            Text1(3).Text = vSocio.Codigo
            FormateaCampo Text1(3)
            If (Modo = 3) Or (Modo = 4) Then
                Text2(3).Text = vSocio.Nombre   'Nom socio
            End If
            Situacion = DevuelveValor("select bloqueo from rsituacion where codsitua = " & vSocio.Situacion)
            NSitua = DevuelveValor("select nomsitua from rsituacion where codsitua = " & vSocio.Situacion)
            If Situacion = 1 Then
                MsgBox "Socio bloqueado por : " & vbCrLf & NSitua, vbInformation, "Situación Especial del Socio"
                Text1(3).Text = ""
                Text2(3).Text = ""
            Else
                
                Observaciones = Trim(DBLet(vSocio.Observaciones, "T"))
                If Observaciones <> "" Then
                    MsgBox Observaciones, vbInformation, "Observaciones del Socio"
                End If
        
            End If
        End If
    Else
        LimpiarDatosSocio
        MsgBox "No existe el socio. Reintroduzca.", vbExclamation
    End If
    Set vSocio = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Socio", Err.Description
End Sub


Private Sub LimpiarDatosSocio()
Dim i As Byte

    Text1(3).Text = ""
    Text2(3).Text = ""
    
End Sub
    

Private Function InsertarLineaEnv(numlinea As String) As Boolean
'Inserta un registro en la tabla de lineas de Albaranes: slialb
Dim Sql As String
Dim vWhere As String
Dim b As Boolean
Dim DentroTRANS As Boolean

    InsertarLineaEnv = False
    Sql = ""
    DentroTRANS = False
    
    'Conseguir el siguiente numero de linea
    vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
'    numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
    
    
    If DatosOkLineaEnv() Then 'Lineas de factura
        'Inserta en tabla "rbodalbaran_variedad"
        Sql = "INSERT INTO rbodalbaran_variedad "
        Sql = Sql & "(numalbar,numlinea,codvarie,unidades,cantidad,precioar,dtolinea,importel,ampliaci,codigiva) "
        Sql = Sql & "VALUES (" & DBSet(txtAux(1).Text, "N") & ", " & numlinea & ","
        Sql = Sql & DBSet(txtAux(5).Text, "T") & "," & DBSet(txtAux(4).Text, "N") & ","
        Sql = Sql & DBSet(txtAux(6).Text, "N") & ","
        Sql = Sql & DBSet(txtAux(7).Text, "N") & "," & DBSet(txtAux(8).Text, "N") & ","
        Sql = Sql & DBSet(txtAux(9).Text, "N") & ","
        Sql = Sql & DBSet(Text2(16).Text, "T") & ","
        Sql = Sql & DBSet(txtAux(10).Text, "N") & ")"
     Else
        Exit Function
     End If
    
    If Sql <> "" Then
        On Error GoTo EInsertarLineaEnv
        conn.BeginTrans
        DentroTRANS = True
        
        'insertar la linea
        conn.Execute Sql
        b = True
    End If
    
    If b Then
        conn.CommitTrans
        InsertarLineaEnv = True
    Else
        conn.RollbackTrans
         InsertarLineaEnv = False
    End If
    Exit Function
    
EInsertarLineaEnv:
    If Err.Number <> 0 Then
        InsertarLineaEnv = False
        If DentroTRANS Then conn.RollbackTrans
        MuestraError Err.Number, "Insertar Lineas Albaranes Retirada de Bodega" & vbCrLf & Err.Description
'        b = False
    End If
'    If b Then
'        Conn.CommitTrans
'        InsertarLinea = True
'    Else
'        Conn.RollbackTrans
'         InsertarLinea = False
'    End If
End Function


Private Function DatosOkLineaEnv() As Boolean
Dim b As Boolean
Dim i As Byte
    
    On Error GoTo EDatosOkLineaEnv

    DatosOkLineaEnv = False
    b = True

    DatosOkLineaEnv = b
    
EDatosOkLineaEnv:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



Private Sub CalcularDatosAlbaran()
Dim i As Integer
Dim cadWHERE As String, Sql As String

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For i = 9 To 31
         Text1(i).Text = ""
    Next i
    
    'Comprobar que hay lineas de facturas_variedad para calcular totales
    cadWHERE = ObtenerWhereCP(False)
    Sql = "Select count(*) from " & NomTablaLineas & " Where " & Replace(cadWHERE, NombreTabla, NomTablaLineas)
    If RegistrosAListar(Sql) = 0 Then
        'Comprobar que hay lineas de rbodalbaran_variedad para calcular totales
        Sql = "Select count(*) from rbodalbaran_variedad Where " & Replace(cadWHERE, NombreTabla, "rbodalbaran_variedad")
        If RegistrosAListar(Sql) = 0 Then Exit Sub
    End If
    
    
    If CalcularDatosAlbaranVenta(cadWHERE, NombreTabla, NomTablaLineas) Then
'        PosicionarData
'        PonerCampos
    Else
        MuestraError Err.Number, "Calculando Totales", Err.Description
    End If
'    Set vFactu = Nothing
End Sub

'
'##Monica
'
Private Function CalcularDatosAlbaranVenta(cadWHERE As String, NomTabla As String, NomTablaLin As String) As Boolean
'cadWhere: cad para la where de la SQL que selecciona las lineas del albaran o la factura
'nomTabla: nombre de la tabla de albaranes(scaalp) o de AlbaranesXFactura(scafpa)
'           segun llamemos desde recepcion de facturas o desde Hco de Facturas
Dim Rs As ADODB.Recordset
Dim i As Integer

Dim Sql As String
Dim cadAux As String
Dim cadAux1 As String

'Aqui vamos acumulando los totales
Dim TotBruto As Currency
Dim TotNeto As Currency
Dim TotImpIVA As Currency

Dim ImpAux As Currency
Dim impiva As Currency
Dim ImpREC As Currency
Dim ImpBImIVA As Currency 'Importe Base imponible a la que hay q aplicar el IVA

Dim vBruto As Currency
Dim vNeto As Currency

Dim exentoIVA As Boolean
Dim conDesplaz As Boolean
    
Dim BaseImp As Currency
Dim BaseIVA1 As Currency
Dim BaseIVA2 As Currency
Dim BaseIVA3 As Currency
    
Dim BrutoFac As Currency
    
Dim ImpIVA1 As Currency
Dim ImpIVA2 As Currency
Dim ImpIVA3 As Currency
    
Dim PorceIVA1 As Currency
Dim PorceIVA2 As Currency
Dim PorceIVA3 As Currency
    
Dim ImpREC1 As Currency
Dim ImpREC2 As Currency
Dim ImpREC3 As Currency
    
Dim PorceREC1 As Currency
Dim PorceREC2 As Currency
Dim PorceREC3 As Currency
    
Dim TipoIVA1 As Currency
Dim TipoIVA2 As Currency
Dim TipoIVA3 As Currency
    
Dim TotalFac As Currency

Dim IvaAnt As Integer
Dim cadWhere1 As String
    
Dim Nulo2 As String
Dim Nulo3 As String

Dim vSeccion As CSeccion

    CalcularDatosAlbaranVenta = False
    On Error GoTo ECalcular

    BaseImp = 0
    BaseIVA1 = 0
    BaseIVA2 = 0
    BaseIVA3 = 0
    
    BrutoFac = 0
    
    ImpIVA1 = 0
    ImpIVA2 = 0
    ImpIVA3 = 0
    
    PorceIVA1 = 0
    PorceIVA2 = 0
    PorceIVA3 = 0
    
    ImpREC1 = 0
    ImpREC2 = 0
    ImpREC3 = 0
    
    PorceREC1 = 0
    PorceREC2 = 0
    PorceREC3 = 0
    
    TipoIVA1 = 0
    TipoIVA2 = 0
    TipoIVA3 = 0
    
    TotalFac = 0

    'Agrupar el importe bruto por tipos de iva
    cadWhere1 = Replace(cadWHERE, "rbodalbaran", "rbodalbaran_variedad")
    Sql = Sql & "SELECT rbodalbaran_variedad.codigiva, sum(importel) as bruto"
    Sql = Sql & " FROM rbodalbaran_variedad "
    Sql = Sql & " WHERE " & cadWhere1
    Sql = Sql & " GROUP BY 1 "
    Sql = Sql & " ORDER BY 1 "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    TotBruto = 0
    TotNeto = 0
    TotImpIVA = 0
    vBruto = 0
    vNeto = 0
    i = 1

    If Not Rs.EOF Then Rs.MoveFirst
    IvaAnt = Rs.Fields(0).Value
    While Not Rs.EOF
            vBruto = Rs.Fields(1).Value
            IvaAnt = Rs.Fields(0).Value
            
            TotBruto = TotBruto + vBruto
            ImpBImIVA = vBruto
        
             ' tenemos que abrir primero la conexion
            Set vSeccion = New CSeccion
            
            If vSeccion.LeerDatos(vParamAplic.SeccionBodega) Then
                If vSeccion.AbrirConta Then
                    'Obtener el % de IVA
                    cadAux = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(IvaAnt), "N")
                    vSeccion.CerrarConta
                End If
            End If
    
            Set vSeccion = Nothing

            'aplicar el IVA a la base imponible de ese tipo
            impiva = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
            
            'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
            'los vamos acumulando
            TotImpIVA = TotImpIVA + impiva
'????
'            If CInt(Me.Combo1(0).ListIndex) = 2 Then  ' tipoivac 0=normal 1=exento 2=recargo equivalencia
'                'Obtener el % de RECARGO
'                cadAux1 = DevuelveDesdeBDNew(cConta, "tiposiva", "porcerec", "codigiva", CStr(IvaAnt), "N")
'
'                'aplicar el RECARGO a la base imponible de ese tipo
'                ImpREC = CalcularPorcentaje(ImpBImIVA, CCur(cadAux1), 2)
'
'                'sumamos todos los RECARGOS para sumarselo a la base imponible total de la factura
'                'los vamos acumulando
'                TotImpIVA = TotImpIVA + ImpREC
'            Else
                cadAux1 = "0"
                ImpREC = 0
'            End If


            Select Case i
                Case 1  'IVA 1
                    TipoIVA1 = IvaAnt 'RS!codigiva

                    BaseIVA1 = ImpBImIVA 'BASE IMPONIBLE

                    PorceIVA1 = cadAux '% de IVA

                    'Importe total con IVA
                    ImpIVA1 = impiva
                    
                    PorceREC1 = cadAux1 '% de REC

                    'Importe total con RECARGO
                    ImpREC1 = ImpREC

                Case 2  'IVA 2
                    TipoIVA2 = IvaAnt 'RS!codigiva

                    BaseIVA2 = ImpBImIVA 'BASE IMPONIBLE

                    PorceIVA2 = cadAux '% de IVA

                    'Importe total con IVA
                    ImpIVA2 = impiva

                    PorceREC2 = cadAux1 '% de REC

                    'Importe total con RECARGO
                    ImpREC2 = ImpREC
                Case 3  'IVA 3
                    TipoIVA3 = IvaAnt 'RS!codigiva

                    BaseIVA3 = ImpBImIVA 'BASE IMPONIBLE

                    PorceIVA3 = cadAux '% de IVA

                    'Importe total con IVA
                    ImpIVA3 = impiva
                    
                    PorceREC3 = cadAux1 '% de REC

                    'Importe total con RECARGO
                    ImpREC3 = ImpREC
            End Select
            
            i = i + 1
        
        
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing

    'Base Imponible
    BaseImp = TotBruto

    'TOTAL de la factura
    TotalFac = BaseImp + TotImpIVA

    'ACTUALIZAMOS EL ALBARAN (tabla albaranes)
    
    For i = 9 To 31
        Text1(i).Text = ""
    Next i
    
    If BaseImp <> 0 Then Text1(12).Text = BaseImp
    
    If BaseIVA1 <> 0 Then Text1(15).Text = Format(BaseIVA1, "###,###,##0.00")
    If ImpIVA1 <> 0 Then Text1(16).Text = Format(ImpIVA1, "###,###,##0.00")
    If ImpREC1 <> 0 Then Text1(30).Text = Format(ImpREC1, "###,###,##0.00")
    If TipoIVA1 <> 0 Then Text1(13).Text = TipoIVA1
    If PorceREC1 <> 0 Then Text1(31).Text = Format(PorceREC1, "##0.00")
    If PorceIVA1 <> 0 Then Text1(14).Text = Format(PorceIVA1, "##0.00")
    
    If BaseIVA2 <> 0 Then Text1(19).Text = Format(BaseIVA2, "###,###,##0.00")
    If ImpIVA2 <> 0 Then Text1(20).Text = Format(ImpIVA2, "###,###,##0.00")
    If ImpREC2 <> 0 Then Text1(28).Text = Format(ImpREC2, "###,###,##0.00")
    If TipoIVA2 <> 0 Then Text1(17).Text = TipoIVA2
    If PorceIVA2 <> 0 Then Text1(18).Text = Format(PorceIVA2, "##0.00")
    If PorceREC2 <> 0 Then Text1(29).Text = Format(PorceREC2, "##0.00")
    
    If BaseIVA3 <> 0 Then Text1(23).Text = Format(BaseIVA3, "###,###,##0.00")
    If ImpIVA3 <> 0 Then Text1(24).Text = Format(ImpIVA3, "###,###,##0.00")
    If ImpREC3 <> 0 Then Text1(26).Text = Format(ImpREC3, "###,###,##0.00")
    If TipoIVA3 <> 0 Then Text1(21).Text = TipoIVA3
    If PorceIVA3 <> 0 Then Text1(22).Text = Format(PorceIVA3, "##0.00")
    If PorceREC3 <> 0 Then Text1(27).Text = Format(PorceREC3, "##0.00")
    
    If TotBruto <> 0 Then Text1(10).Text = Format(TotBruto, "###,###,##0.00")
    If TotalFac <> 0 Then Text1(25).Text = Format(TotalFac, "###,###,##0.00")

    CalcularDatosAlbaranVenta = True

ECalcular:
    If Err.Number <> 0 Then
        CalcularDatosAlbaranVenta = False
    Else
        CalcularDatosAlbaranVenta = True
    End If
End Function
