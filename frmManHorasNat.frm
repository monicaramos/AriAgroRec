VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManHorasNat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrada de Horas Trabajadores Natural "
   ClientHeight    =   9795
   ClientLeft      =   195
   ClientTop       =   180
   ClientWidth     =   17445
   Icon            =   "frmManHorasNat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   17445
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameFiltro 
      Height          =   705
      Left            =   5400
      TabIndex        =   33
      Top             =   60
      Width           =   2895
      Begin VB.ComboBox cboFiltro 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmManHorasNat.frx":000C
         Left            =   120
         List            =   "frmManHorasNat.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   210
         Width           =   2655
      End
   End
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3780
      TabIndex        =   30
      Top             =   60
      Width           =   1515
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   210
         TabIndex        =   31
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "C�lculo Horas Productivas"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Exportar Fichero Excel"
               Object.Tag             =   "2"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   150
      TabIndex        =   28
      Top             =   60
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
               Object.ToolTipText     =   "Listado Entradas Capataz"
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
      Index           =   9
      Left            =   6000
      MaxLength       =   6
      TabIndex        =   27
      Tag             =   "Forfait|T|S|||horas|codforfait||S|"
      Top             =   4950
      Width           =   585
   End
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
      Height          =   350
      Index           =   5
      Left            =   6630
      MaskColor       =   &H00000000&
      TabIndex        =   26
      ToolTipText     =   "Buscar forfaits"
      Top             =   4950
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
      Index           =   9
      Left            =   6840
      TabIndex        =   25
      Top             =   4950
      Visible         =   0   'False
      Width           =   915
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
      Index           =   8
      Left            =   14100
      TabIndex        =   24
      Top             =   4950
      Visible         =   0   'False
      Width           =   1125
   End
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
      Height          =   350
      Index           =   4
      Left            =   13890
      MaskColor       =   &H00000000&
      TabIndex        =   23
      ToolTipText     =   "Buscar variedad"
      Top             =   4920
      Visible         =   0   'False
      Width           =   195
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
      Index           =   8
      Left            =   12990
      MaxLength       =   6
      TabIndex        =   10
      Tag             =   "Variedad|N|S|||horas|codvarie|000000|S|"
      Top             =   4950
      Width           =   885
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
      Index           =   7
      Left            =   4560
      MaxLength       =   6
      TabIndex        =   2
      Tag             =   "C�digo|N|N|0|99|horas|codalmac|00|S|"
      Top             =   4950
      Width           =   465
   End
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
      Height          =   350
      Index           =   3
      Left            =   5070
      MaskColor       =   &H00000000&
      TabIndex        =   22
      ToolTipText     =   "Buscar almac�n"
      Top             =   4950
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
      Index           =   7
      Left            =   5265
      TabIndex        =   21
      Top             =   4950
      Visible         =   0   'False
      Width           =   675
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
      Index           =   6
      Left            =   8550
      MaxLength       =   6
      TabIndex        =   4
      Tag             =   "Horas Productivas|N|N|||horas|horasproduc|##0.00||"
      Top             =   4950
      Width           =   735
   End
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
      Height          =   350
      Index           =   2
      Left            =   12150
      MaskColor       =   &H00000000&
      TabIndex        =   20
      ToolTipText     =   "Buscar fecha"
      Top             =   4950
      Visible         =   0   'False
      Width           =   195
   End
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
      Height          =   350
      Index           =   1
      Left            =   4305
      MaskColor       =   &H00000000&
      TabIndex        =   19
      ToolTipText     =   "Buscar fecha"
      Top             =   4905
      Visible         =   0   'False
      Width           =   195
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
      Left            =   10980
      MaxLength       =   8
      TabIndex        =   7
      Tag             =   "Fecha Recibo|F|S|||horas|fecharec|dd/mm/yyyy||"
      Top             =   4980
      Width           =   1170
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
      Left            =   10110
      MaxLength       =   8
      TabIndex        =   6
      Tag             =   "Horas Extra|N|S|||horas|horasext|###,##0.00||"
      Top             =   4950
      Width           =   810
   End
   Begin VB.CheckBox chkAux 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   1
      Left            =   12780
      TabIndex        =   9
      Tag             =   "Int.Contable|N|N|||horas|intconta|||"
      Top             =   4935
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CheckBox chkAux 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   0
      Left            =   12450
      TabIndex        =   8
      Tag             =   "Int.Contable|N|N|||horas|pasaridoc|||"
      Top             =   4935
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
      Index           =   3
      Left            =   9360
      MaxLength       =   8
      TabIndex        =   5
      Tag             =   "Complementos|N|N|||horas|compleme|###,##0.00||"
      Top             =   4950
      Width           =   750
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
      Index           =   0
      Left            =   1125
      TabIndex        =   18
      Top             =   4905
      Visible         =   0   'False
      Width           =   1815
   End
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
      Height          =   350
      Index           =   0
      Left            =   930
      MaskColor       =   &H00000000&
      TabIndex        =   17
      ToolTipText     =   "Buscar trabajador"
      Top             =   4905
      Visible         =   0   'False
      Width           =   195
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
      Index           =   2
      Left            =   7800
      MaxLength       =   6
      TabIndex        =   3
      Tag             =   "Horas Dia|N|N|||horas|horasdia|##0.00||"
      Top             =   4935
      Width           =   705
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
      Left            =   14970
      TabIndex        =   11
      Top             =   9210
      Visible         =   0   'False
      Width           =   1065
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
      Left            =   16200
      TabIndex        =   12
      Top             =   9210
      Visible         =   0   'False
      Width           =   1065
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
      Index           =   1
      Left            =   2985
      MaxLength       =   10
      TabIndex        =   1
      Tag             =   "Fecha|F|N|||horas|fechahora|dd/mm/yyyy|S|"
      Top             =   4905
      Width           =   1320
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
      Left            =   60
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "C�digo|N|N|0|999999|horas|codtraba|000000|S|"
      Top             =   4905
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmManHorasNat.frx":0050
      Height          =   8145
      Left            =   150
      TabIndex        =   15
      Top             =   840
      Width           =   17130
      _ExtentX        =   30215
      _ExtentY        =   14367
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
      Left            =   16200
      TabIndex        =   16
      Top             =   9210
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   9060
      Width           =   2685
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
         Height          =   255
         Left            =   45
         TabIndex        =   14
         Top             =   180
         Width           =   2505
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   2790
      Top             =   0
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
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   16830
      TabIndex        =   32
      Top             =   210
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
         Caption         =   "&Ver todos"
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
      Begin VB.Menu mnCalculoHorasProd 
         Caption         =   "&C�lculo Horas Prod."
         Shortcut        =   ^C
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnExportacion 
         Caption         =   "Exportar Excel"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnfiltro 
      Caption         =   "Filtro"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnFiltro1 
         Caption         =   "A�o Actual"
      End
      Begin VB.Menu mnFiltro2 
         Caption         =   "A�o Actual y Anterior"
      End
      Begin VB.Menu mnBarra4 
         Caption         =   "-"
      End
      Begin VB.Menu mnFiltro3 
         Caption         =   "Sin Filtro"
      End
   End
End
Attribute VB_Name = "frmManHorasNat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MONICA  +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-

' **************** PER A QUE FUNCIONE EN UN ATRE MANTENIMENT ********************
' 0. Posar-li l'atribut Datasource a "adodc1" del Datagrid1. Canviar el Caption
'    del formulari
' 1. Canviar els TAGs i els Maxlength de TextAux(0) i TextAux(1)
' 2. En PonerModo(vModo) repasar els indexs del botons, per si es canvien
' 3. En la funci� BotonAnyadir() canviar la taula i el camp per a SugerirCodigoSiguienteStr
' 4. En la funci� BotonBuscar() canviar el nom de la clau primaria
' 5. En la funci� BotonEliminar() canviar la pregunta, les descripcions de la
'    variable SQL i el contingut del DELETE
' 6. En la funci� PonerLongCampos() posar els camps als que volem canviar el MaxLength quan busquem
' 7. En Form_Load() repasar la barra d'iconos (per si es vol canviar alg�n) i
'    canviar la consulta per a vore tots els registres
' 8. En Toolbar1_ButtonClick repasar els indexs de cada bot� per a que corresponguen
' 9. En la funci� CargaGrid canviar l'ORDER BY (normalment per la clau primaria);
'    canviar adem�s els noms dels camps, el format i si fa falta la cantitat;
'    repasar els index dels botons modificar i eliminar.
'    NOTA: si en Form_load ya li he posat clausula WHERE, canviar el `WHERE` de
'    `SQL = CadenaConsulta & " WHERE " & vSQL` per un `AND`
' 10. En txtAux_LostFocus canviar el mensage i el format del camp
' 11. En la funci� DatosOk() canviar els arguments de DevuelveDesdeBD i el mensage
'    en cas d'error
' 12. En la funci� SepuedeBorrar() canviar les comprovacions per a vore si es pot
'    borrar el registre
' *******************************SI N'HI HA COMBO*******************************
' 0. Comprovar que en el SQL de Form_Load() es fa�a refer�ncia a la taula del Combo
' 1. Pegar el Combo1 al  costat dels TextAux. Canviar-li el TAG
' 2. En BotonModificar() canviar el camp del Combo
' 3. En CargaCombo() canviar la consulta i els noms del camps, o posar els valor
'    a ma si no es llig de cap base de datos els valors del Combo

Option Explicit

Private Const IdPrograma = 8010



Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'atre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

Private CadenaConsulta As String
Private CadB As String

Private WithEvents frmTra As frmManTraba 'mantenimiento de trabajadores
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmMens As frmManTraba 'mantenimiento de trabajadores
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmAlm As frmBasico2 'mantenimiento de almacenes propios
Attribute frmAlm.VB_VarHelpID = -1
Private WithEvents frmVar As frmComVar 'variedades
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmFor As frmBasico2 'mantenimiento de forfaits de comercial
Attribute frmFor.VB_VarHelpID = -1

Dim Modo As Byte
'----------- MODOS --------------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la b�squeda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edici� del camp
'   3.-  Inserci� de nou registre
'   4.-  Modificar
'--------------------------------------------------
Dim PrimeraVez As Boolean
Dim Indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim indCodigo As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim i As Integer

' utilizado para buscar por checks
Private BuscaChekc As String

Private Filtro As Byte
Dim cadFiltro As String



Private Sub PonerModo(vModo)
Dim B As Boolean

    Modo = vModo
    
    B = (Modo = 2)
    If B Then
        PonerContRegIndicador lblIndicador, adodc1, CadB
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    For i = 0 To txtAux.Count - 1
        txtAux(i).visible = Not B
        txtAux(i).BackColor = vbWhite
    Next i
    
    txtAux2(0).visible = Not B
    txtAux2(7).visible = Not B
    txtAux2(8).visible = Not B
    txtAux2(9).visible = Not B
    
    For i = 0 To btnBuscar.Count - 1
        btnBuscar(i).visible = Not B
    Next i
    
    chkAux(0).visible = Not B
    chkAux(1).visible = Not B

    CmdAceptar.visible = Not B
    CmdCancelar.visible = Not B
    DataGrid1.Enabled = B
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = B
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu  'En funcion del usuario
    
    'Si estamos modo Modificar bloquear clave primaria
    BloquearTxt txtAux(0), (Modo = 4)
    BloquearTxt txtAux(1), (Modo = 4)
    BloquearTxt txtAux(7), (Modo = 4)
    BloquearTxt txtAux(8), (Modo = 4)
    BloquearTxt txtAux(9), (Modo = 4)
    BloquearTxt txtAux(5), (Modo = 4) Or (Modo = 3)
    txtAux(5).visible = (Modo = 1)
    BloquearBtn Me.btnBuscar(0), (Modo = 4)
    BloquearBtn Me.btnBuscar(1), (Modo = 4)
    BloquearBtn Me.btnBuscar(3), (Modo = 4)
    BloquearBtn Me.btnBuscar(4), (Modo = 4)
    BloquearBtn Me.btnBuscar(5), (Modo = 4)
    BloquearBtn Me.btnBuscar(2), (Modo = 4) Or (Modo = 3)
    
    BloquearChk Me.chkAux(0), (Modo = 4) Or (Modo = 3)
    BloquearChk Me.chkAux(1), (Modo = 4) Or (Modo = 3)
    Me.chkAux(0).visible = (Modo = 1)
    Me.chkAux(1).visible = (Modo = 1)
    
    'El nro de parte unicamente lo podemos buscar
'    txtAux(8).Enabled = (Modo = 1)
'    txtAux(8).visible = (Modo = 1)
    
End Sub


Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botones del la toobar y del menu, segun el modo en que estemos
Dim B As Boolean

    B = (Modo = 2)
    'Busqueda
    Toolbar1.Buttons(5).Enabled = B
    Me.mnBuscar.Enabled = B
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = B
    Me.mnVerTodos.Enabled = B
    'Imprimir
    Toolbar1.Buttons(8).Enabled = B
    Me.mnImprimir.Enabled = B
    
    'Insertar
    Toolbar1.Buttons(1).Enabled = B And Not DeConsulta
    Me.mnNuevo.Enabled = B And Not DeConsulta
    
    B = (B And adodc1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    Me.mnModificar.Enabled = B
    'Eliminar
    Toolbar1.Buttons(3).Enabled = B
    Me.mnEliminar.Enabled = B
    
    
    Toolbar2.Buttons(2).Enabled = B And UCase(Dir(App.Path & "\controlnomi.cfg")) = UCase("controlnomi.cfg")
    Me.mnExportacion.Enabled = B And UCase(Dir(App.Path & "\controlnomi.cfg")) = UCase("controlnomi.cfg")
    
'    'Imprimir
'    Toolbar1.Buttons(11).Enabled = b
'    Me.mnImprimir.Enabled = b
    
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
    CargaGrid CadB, True 'primer de tot carregue tot el grid
   
    'Situamos el grid al final
    AnyadirLinea DataGrid1, adodc1
         
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 240
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    For i = 0 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i
    txtAux(1).Text = Format(Now, "dd/mm/yyyy")
    txtAux(8).Text = "0"
    txtAux2(0).Text = ""
    txtAux2(7).Text = ""
    txtAux2(8).Text = ""
    txtAux2(9).Text = ""
    
    chkAux(0).Value = 0
    chkAux(1).Value = 0

    txtAux(2).Text = 0
    txtAux(3).Text = 0
    txtAux(4).Text = 0

    LLamaLineas anc, 3 'Pone el form en Modo=3, Insertar
       
    'Ponemos el foco
    PonerFoco txtAux(0)
End Sub

Private Sub BotonVerTodos()
    CadB = ""
    CargaGrid ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    ' ***************** canviar per la clau primaria ********
    CargaGrid "horas.codtraba = -1"
    '*******************************************************************************
    'Buscar
    For i = 0 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i
    chkAux(0).Value = 0
    chkAux(1).Value = 0
    Me.txtAux2(0).Text = ""
    Me.txtAux2(7).Text = ""
    Me.txtAux2(8).Text = ""
    Me.txtAux2(9).Text = ""
    
    
    LLamaLineas DataGrid1.Top + 240, 1 'Pone el form en Modo=1, Buscar
    PonerFoco txtAux(0)
End Sub

Private Sub BotonModificar()
    Dim anc As Single
    Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top '495 '545
    End If

    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text ' codtraba
    txtAux2(0).Text = DataGrid1.Columns(1).Text ' nomtraba
    txtAux(7).Text = DataGrid1.Columns(3).Text 'codalmac
    txtAux2(7).Text = DataGrid1.Columns(4).Text 'nomalmac
    txtAux(1).Text = DataGrid1.Columns(2).Text 'fechahora
    ' ***** canviar-ho pel nom del camp del combo *********
'    SelComboBool DataGrid1.Columns(2).Text, Combo1(0)
    ' *****************************************************
    txtAux(2).Text = DataGrid1.Columns(5).Text 'horasdia
    txtAux(6).Text = DataGrid1.Columns(6).Text 'horasproduc
    txtAux(3).Text = DataGrid1.Columns(7).Text 'compleme
    txtAux(4).Text = DataGrid1.Columns(8).Text 'horasextras
    txtAux(5).Text = DataGrid1.Columns(13).Text 'fecharecep
    txtAux(8).Text = DataGrid1.Columns(9).Text 'codigo variedad
    txtAux2(8).Text = DataGrid1.Columns(10).Text 'nombre variedad
    txtAux(9).Text = DataGrid1.Columns(11).Text 'codigo variedad
    txtAux2(9).Text = DataGrid1.Columns(12).Text 'nombre variedad
    
    Me.chkAux(0).Value = Me.adodc1.Recordset!pasaridoc
    Me.chkAux(1).Value = Me.adodc1.Recordset!intconta

    LLamaLineas anc, 4 'Pone el form en Modo=4, Modificar
   
    'Como es modificar
    PonerFoco txtAux(2)
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    'Fijamos el ancho
    For i = 0 To txtAux.Count - 1
        txtAux(i).Top = alto
    Next i
    
    ' ### [Monica] 12/09/2006
    txtAux2(0).Top = alto
    txtAux2(7).Top = alto
    txtAux2(8).Top = alto
    txtAux2(9).Top = alto
    For i = 0 To btnBuscar.Count - 1
        btnBuscar(i).Top = alto - 5
    Next i
    
    Me.chkAux(0).Top = alto
    Me.chkAux(1).Top = alto
End Sub

Private Sub BotonEliminar()
Dim Sql As String
Dim temp As Boolean

    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
'    If Not SepuedeBorrar Then Exit Sub
        
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCampo(txtAux(0))) Then Exit Sub
    ' ***************************************************************************
    
    '*************** canviar els noms i el DELETE **********************************
    Sql = "�Seguro que desea eliminar el Registro?"
    Sql = Sql & vbCrLf & "Trabajador: " & adodc1.Recordset.Fields(0) & " " & adodc1.Recordset.Fields(1)
    Sql = Sql & vbCrLf & "Fecha: " & adodc1.Recordset.Fields(2)
    Sql = Sql & vbCrLf & "Almac�n: " & adodc1.Recordset.Fields(3)
    Sql = Sql & vbCrLf & "Variedad: " & adodc1.Recordset.Fields(9) & " " & adodc1.Recordset.Fields(10)
    
    If DBLet(adodc1.Recordset.Fields(12), "T") <> "" Then
        Sql = Sql & vbCrLf & "Forfait: " & adodc1.Recordset.Fields(11) & " " & adodc1.Recordset.Fields(12)
    End If
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        Sql = "Delete from horas where codtraba=" & adodc1.Recordset!CodTraba
        Sql = Sql & " and fechahora = " & DBSet(adodc1.Recordset!FechaHora, "F")
        Sql = Sql & " and codalmac = " & DBLet(adodc1.Recordset!codAlmac)
        Sql = Sql & " and codvarie = " & DBLet(adodc1.Recordset!Codvarie, "N")
        
        If DBLet(adodc1.Recordset.Fields(12), "T") <> "" Then
            Sql = Sql & " and codforfait = " & DBLet(adodc1.Recordset!codforfait, "T")
        End If
        
        conn.Execute Sql
        CargaGrid CadB
        
        temp = SituarDataTrasEliminar(adodc1, NumRegElim, True)
        PonerModoOpcionesMenu
        adodc1.Recordset.Cancel
    End If
    Exit Sub
    
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los txtAux
    PonerLongCamposGnral Me, Modo, 3
End Sub

Private Sub btnBuscar_Click(Index As Integer)
 TerminaBloquear
    
    Select Case Index
        Case 0 'grupo de productos
            
            Indice = Index
            Set frmTra = New frmManTraba
            frmTra.DatosADevolverBusqueda = "0|2|"
            frmTra.Show vbModal
            Set frmTra = Nothing
            PonerFoco txtAux(Indice)
    
       Case 1, 2 ' Fecha
            Dim esq As Long
            Dim dalt As Long
            Dim menu As Long
            Dim obj As Object
        
            Set frmC = New frmCal
            
            Indice = Index
            
            esq = btnBuscar(Index).Left
            dalt = btnBuscar(Index).Top
                
            Set obj = btnBuscar(Index).Container
              
              While btnBuscar(Index).Parent.Name <> obj.Name
                    esq = esq + obj.Left
                    dalt = dalt + obj.Top
                    Set obj = obj.Container
              Wend
            
            menu = Me.Height - Me.ScaleHeight 'ac� tinc el heigth del men� i de la toolbar
        
            frmC.Left = esq + btnBuscar(Index).Parent.Left + 30
            frmC.Top = dalt + btnBuscar(Index).Parent.Top + btnBuscar(Index).Height + menu - 40
        
            btnBuscar(Index).Tag = Index '<===
            ' *** repasar si el camp es txtAux o Text1 ***
            If Index = 1 Then
                If txtAux(1).Text <> "" Then frmC.NovaData = txtAux(1).Text
            Else
                If txtAux(5).Text <> "" Then frmC.NovaData = txtAux(5).Text
            End If
            
            ' ********************************************
        
            frmC.Show vbModal
            Set frmC = Nothing
            ' *** repasar si el camp es txtAux o Text1 ***
            PonerFoco txtAux(1) '<===
            ' ********************************************
     
        Case 3 'codigo de almacen
            Set frmAlm = New frmBasico2
            
            AyudaAlmacenCom frmAlm, txtAux(7).Text
            
            Set frmAlm = Nothing
            
            PonerFoco txtAux(7)
            
        Case 4 'codigo de variedad
            AbrirFrmVariedades 9

        Case 5 ' codigo de forfait
            AbrirFrmManForfaits 9


    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.adodc1, 1
End Sub


Private Sub cboFiltro_Change()
    CargarSqlFiltro
    
End Sub

Private Sub cboFiltro_KeyPress(KeyAscii As Integer)
    If Modo = 2 Then CargaGrid CadB
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

    Select Case Modo
        Case 1 'BUSQUEDA
            CadB = ObtenerBusqueda3(Me, False, BuscaChekc)
            If CadB <> "" Then
                CargaGrid CadB
                PonerModo 2
'                lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
                PonerFocoGrid Me.DataGrid1
            End If
            
        Case 3 'INSERTAR
            If DatosOK Then
                If InsertarDesdeForm(Me) Then
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                        CargaGrid CadB
                        
                        cmdCancelar_Click
'                        If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveLast
                        If Not adodc1.Recordset.EOF Then
                            adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & NuevoCodigo)
                        End If
                        cmdRegresar_Click
                    Else
                        BotonAnyadir
                    End If
                    CadB = ""
                End If
            End If
            
        Case 4 'MODIFICAR
            If DatosOK Then
'                If ModificaDesdeFormulario(Me) Then
                If ModificaDesdeForm Then
                    TerminaBloquear
                    i = adodc1.Recordset.Fields(0)
                    PonerModo 2
                    CargaGrid CadB
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & i)
                    PonerFocoGrid Me.DataGrid1
                End If
            End If
    End Select
End Sub

Private Sub cmdCancelar_Click()
    On Error Resume Next
    
    Select Case Modo
        Case 1 'b�squeda
            CargaGrid CadB
        Case 3 'insertar
            DataGrid1.AllowAddNew = False
            CargaGrid
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        Case 4 'modificar
            TerminaBloquear
    End Select
    
    PonerModo 2
    
'    If CadB <> "" Then
'        lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
''    Else
''        lblIndicador.Caption = ""
'    End If
    
    PonerFocoGrid Me.DataGrid1
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim i As Integer
Dim J As Integer
Dim Aux As String

    If adodc1.Recordset.EOF Then
        MsgBox "Ning�n registro devuelto.", vbExclamation
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
            cad = cad & adodc1.Recordset.Fields(J) & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Modo = 2 Then PonerContRegIndicador lblIndicador, adodc1, CadB
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault

    If PrimeraVez Then
        PrimeraVez = False
        If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
            BotonAnyadir
        Else
            PonerModo 2
             If Me.CodigoActual <> "" Then
                SituarData Me.adodc1, "codprodu=" & CodigoActual, "", True
            End If
        End If
    End If
End Sub

Private Sub Form_Load()

    'Icono del formulario
    Me.Icon = frmPpal.Icon


    PrimeraVez = True

    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'el 1 es separadors
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2   'Todos
        'el 4 i el 5 son separadors
        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
        
        .Buttons(8).Image = 10  'imprimir
        
        'el 9 i el 10 son separadors
'        .Buttons(10).Image = 17  'calculo de las horas productivas
'        .Buttons(11).Image = 10  'imprimir
'        .Buttons(12).Image = 34  'Exportacion a Excel
'        .Buttons(13).Image = 11  'Salir
    End With
    
    With Me.Toolbar2
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 17  'calculo de las horas productivas
        .Buttons(2).Image = 34  'Exportacion a Excel
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 12
    End With
    
    
    '## A mano
'    chkVistaPrevia.Value = CheckValueLeer(Name)

'@@--
'    LeerFiltro True
'    PonerFiltro Filtro

    
    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT horas.codtraba, straba.nomtraba, horas.fechahora, horas.codalmac, salmpr.nomalmac, "
    CadenaConsulta = CadenaConsulta & "horas.horasdia, horas.horasproduc, horas.compleme, horas.horasext, "
    CadenaConsulta = CadenaConsulta & " horas.codvarie, variedades.nomvarie,"
    CadenaConsulta = CadenaConsulta & " horas.codforfait, forfaits.nomconfe,"
    CadenaConsulta = CadenaConsulta & " horas.fecharec, "
    CadenaConsulta = CadenaConsulta & " horas.pasaridoc,  IF(pasaridoc=1,'*','') as pasari, horas.intconta,  IF(intconta=1,'*','') as intcon "
    CadenaConsulta = CadenaConsulta & " FROM straba, salmpr, variedades, horas LEFT JOIN  forfaits ON horas.codforfait = forfaits.codforfait "
    CadenaConsulta = CadenaConsulta & " WHERE horas.codtraba = straba.codtraba and "
    CadenaConsulta = CadenaConsulta & " horas.codalmac = salmpr.codalmac and "
'    CadenaConsulta = CadenaConsulta & " horas.codforfait = forfaits.codforfait and "
    CadenaConsulta = CadenaConsulta & " horas.codvarie = variedades.codvarie "
'     CadenaConsulta = CadenaConsulta & " FROM (((straba inner join horas on straba.codtraba = horas.codtraba)"
'     CadenaConsulta = CadenaConsulta & " inner join salmpr on horas.codalmac = salmpr.codalmac)"
'     CadenaConsulta = CadenaConsulta & " inner join variedades on horas.codvarie = variedades.codvarie)"
'     CadenaConsulta = CadenaConsulta & " left join forfaits on horas.codforfait = forfaits.codforfait"
'     CadenaConsulta = CadenaConsulta & " where (1=1) "
    '************************************************************************
    
    CargaFiltros
    
    cboFiltro.ListIndex = 0
    
    
    CadB = ""
    CargaGrid "horas.codtraba is null "
    
'    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'        BotonAnyadir
'    Else
'        PonerModo 2
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Modo = 4 Then TerminaBloquear
    
    LeerFiltro False
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(7).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo almacen
    txtAux2(7).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre almacen
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    If Indice = 1 Then
        txtAux(1).Text = Format(vFecha, "dd/mm/yyyy") '<===
    Else
        txtAux(5).Text = Format(vFecha, "dd/mm/yyyy") '<===
    End If
    ' ********************************************
End Sub

Private Sub frmFor_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(9).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo forfait
    txtAux2(9).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre forfait
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo trabajador
    txtAux2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre trabajador
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo variedad
    txtAux2(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre variedad
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnCalculoHorasProd_Click()
    BotonCalculoHorasProd
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnExportacion_Click()
    If Not CargarCondicion Then Exit Sub
    
    Shell App.Path & "\nomina.exe /D|" & vUsu.CadenaConexion & "|", vbNormalFocus
        
End Sub


Private Sub mnFiltro1_Click()
    PonerFiltro 1
End Sub

Private Sub mnFiltro2_Click()
    PonerFiltro 2
End Sub

Private Sub mnFiltro3_Click()
    PonerFiltro 3
End Sub

Private Sub PonerFiltro(NumFilt As Byte)
    Filtro = NumFilt
    Me.mnFiltro1.Checked = (NumFilt = 1)
    Me.mnFiltro2.Checked = (NumFilt = 2)
    Me.mnFiltro3.Checked = (NumFilt = 3)
End Sub




Private Sub mnImprimir_Click()
    AbrirListadoNominas (15)
End Sub

Private Sub mnModificar_Click()
    'Comprobaciones
    '--------------
    If adodc1.Recordset.EOF Then Exit Sub
    
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCampo(txtAux(0))) Then Exit Sub
    
    
    'Preparamos para modificar
    '-------------------------
    If BLOQUEADesdeFormulario2(Me, adodc1, 1) Then BotonModificar
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
                mnNuevo_Click
        Case 2
                mnModificar_Click
        Case 3
                mnEliminar_Click
        Case 5
                mnBuscar_Click
        Case 6
                mnVerTodos_Click
        Case 8
            'MsgBox "Imprimir...under construction"
                mnImprimir_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String, Optional Ascendente As Boolean)
    Dim Sql As String
    Dim tots As String
    
'@@--
'    CadenaFiltro = AnyadeCadenaFiltro()

    CargarSqlFiltro
    
    If vSQL <> "" Then
        Sql = CadenaConsulta & " and " & cadFiltro & " AND " & vSQL
    Else
        Sql = CadenaConsulta & " and " & cadFiltro & "  "
    End If
    If Ascendente Then
        Sql = Sql & " ORDER BY  horas.fechahora, horas.codtraba "
    Else
        '********************* canviar el ORDER BY *********************++
        Sql = Sql & " ORDER BY  horas.fechahora desc, horas.codtraba, horas.codalmac, horas.codvarie "
        '**************************************************************++
    End If
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, Sql, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "S|txtAux(0)|T|C�digo|1000|;S|btnBuscar(0)|B||195|;S|txtAux2(0)|T|Trabajador|2100|;"
    tots = tots & "S|txtAux(1)|T|Fecha|1400|;S|btnBuscar(1)|B||195|;"
    tots = tots & "S|txtAux(7)|T|C�d.|600|;S|btnBuscar(3)|B||195|;S|txtAux2(7)|T|Almac�n|1000|;"
    tots = tots & "S|txtAux(2)|T|Horas|750|;"
    tots = tots & "S|txtAux(6)|T|H.Pr.|740|;"
    tots = tots & "S|txtAux(3)|T|Complem.|1100|;"
    tots = tots & "S|txtAux(4)|T|H.Extra|800|;"
    tots = tots & "S|txtAux(8)|T|Codigo|1000|;S|btnBuscar(4)|B||195|;"
    tots = tots & "S|txtAux2(8)|T|Variedad|1400|;"
    tots = tots & "S|txtAux(9)|T|Codigo|1000|;S|btnBuscar(5)|B||195|;"
    tots = tots & "S|txtAux2(9)|T|Forfait|1550|;"
    tots = tots & "S|txtAux(5)|T|F.Recibo|1400|;S|btnBuscar(2)|B||195|;N||||0|;S|chkAux(0)|CB|IA|360|;N||||0|;S|chkAux(1)|CB|IC|360|;"
    arregla tots, DataGrid1, Me, 350
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgLeft
    DataGrid1.Columns(3).Alignment = dbgLeft
    DataGrid1.Columns(9).Alignment = dbgLeft
'    DataGrid1.Columns(12).Alignment = dbgCenter
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 ' calculo de horas productivas
                mnCalculoHorasProd_Click
        Case 2
                mnExportacion_Click
    End Select
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    If Index = 6 And Modo = 3 Then
        txtAux(Index).Enabled = False
    End If

    ConseguirFoco txtAux(Index), Modo
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0 'codigo de trabajador
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "straba", "nomtraba")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe el Trabajador: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmTra = New frmManTraba
                        frmTra.DatosADevolverBusqueda = "0|1|"
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmTra.Show vbModal
                        Set frmTra = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                Else
                    'sacamos el almacen por defecto
                    If Modo = 1 Or Modo = 4 Then Exit Sub
                    
                    txtAux(7).Text = DevuelveDesdeBDNew(cAgro, "straba", "codalmac", "codtraba", txtAux(Index).Text, "N")
                    PonerFormatoEntero txtAux(7)
                    If txtAux(7).Text <> "" Then
                        txtAux2(7).Text = PonerNombreDeCod(txtAux(7), "salmpr", "nomalmac")
                    End If
                End If
            End If
        
        Case 1, 5 'fecha
            '[Monica]28/08/2013: comprobamos que la fecha est� en la campa�a
            PonerFormatoFecha txtAux(Index), True
            
        Case 2, 4, 6 'horas diarias, horas extra y horas productivas
            If txtAux(Index).Text <> "" Then
                cadMen = TransformaPuntosComas(txtAux(Index).Text)
                txtAux(Index).Text = Format(cadMen, "##0.00")
            End If
            
            ' si estamos insertando las horas productivas son las horas diarias
            If Index = 2 And Modo = 3 Then
                txtAux(6).Text = txtAux(2).Text
            End If
            
        Case 3 'complemento
            If txtAux(Index).Text <> "" Then
                cadMen = TransformaPuntosComas(txtAux(Index).Text)
                txtAux(Index).Text = Format(cadMen, "###,##0.00")
            End If
    
        Case 7 'codigo de almacen
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "salmpr", "nomalmac")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe el Almac�n: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "�Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmAlm = New frmComercial
                        frmAlm.DatosADevolverBusqueda = "0|1|"
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmAlm.Show vbModal
                        Set frmAlm = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
            End If
    
        Case 8 'codigo de variedad
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "variedades", "nomvarie")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe la Variedad " & txtAux(Index).Text & ". Reintroduzca." & vbCrLf
                    MsgBox cadMen, vbExclamation
                    PonerFoco txtAux(Index)
                End If
            Else
                txtAux2(Index).Text = ""
            End If
    
    End Select
    
End Sub

Private Function DatosOK() As Boolean
'Dim Datos As String
Dim B As Boolean
Dim Sql As String
Dim Mens As String


    B = CompForm(Me)
    If Not B Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
        Sql = "select count(*) from horas where codtraba = " & DBSet(txtAux(0).Text, "N")
        Sql = Sql & " and fechahora = " & DBSet(txtAux(1).Text, "F")
        Sql = Sql & " and codalmac = " & DBSet(txtAux(7).Text, "N")
        Sql = Sql & " and codvarie = " & DBSet(txtAux(8).Text, "N")
        If TotalRegistros(Sql) <> 0 Then
            MsgBox "El trabajador existe para esta fecha, almac�n, variedad. Reintroduzca.", vbExclamation
            PonerFoco txtAux(0)
            B = False
        End If
        
        If B Then
            If txtAux(9).Text = "" Then txtAux(9).Text = " "
        End If
    End If
    
    If B And (Modo = 3 Or Modo = 4) Then
        If Not EntreFechas(vParam.FecIniCam, txtAux(1).Text, vParam.FecFinCam) Then
            MsgBox "La fecha introducida no se encuentra dentro de campa�a. Revise.", vbExclamation
            B = False
        End If
    End If
    
    DatosOK = B
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub


'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del rat�n.
'Private Sub DataGrid1_GotFocus()
'  WheelHook DataGrid1
'End Sub
'Private Sub DataGrid1_Lostfocus()
'  WheelUnHook
'End Sub


Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 0: KEYBusqueda KeyAscii, 0 'codigo de trabajador
                Case 1: KEYBusqueda KeyAscii, 1 'fecha hora
                Case 7: KEYBusqueda KeyAscii, 3 'codigo almacen
                Case 8: KEYBusqueda KeyAscii, 4 'codigo variedad
                Case 9: KEYBusqueda KeyAscii, 5 'codigo forfait
                Case 5: KEYBusqueda KeyAscii, 2 'fecha recibo
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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
    btnBuscar_Click (Indice)
End Sub

Private Sub BotonCalculoHorasProd()
    AbrirListadoNominas (16)
    CargaGrid
End Sub

Private Sub AbrirFrmVariedades(Indice As Integer)
    indCodigo = 8
    Set frmVar = New frmComVar
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.CodigoActual = txtAux(indCodigo)
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub

Private Sub AbrirFrmManForfaits(Indice As Integer)
    Set frmFor = New frmBasico2
    AyudaForfaitsCom frmFor, txtAux(Indice)
    Set frmFor = Nothing
    
    PonerFoco txtAux(Indice)
End Sub

Private Function ModificaDesdeForm() As Boolean
Dim Sql As String

    On Error GoTo eModificaDesdeForm
    
    ModificaDesdeForm = False
    
    Sql = "update horas set "
    Sql = Sql & " horasdia = " & DBSet(ImporteSinFormato(txtAux(2).Text), "N")
    Sql = Sql & ", horasproduc = " & DBSet(ImporteSinFormato(txtAux(6).Text), "N")
    Sql = Sql & ", compleme = " & DBSet(ImporteSinFormato(txtAux(3).Text), "N")
    Sql = Sql & ", horasext = " & DBSet(ImporteSinFormato(txtAux(4).Text), "N")
    Sql = Sql & " where codtraba = " & DBSet(txtAux(0).Text, "N")
    Sql = Sql & " and fechahora = " & DBSet(txtAux(1).Text, "F")
    Sql = Sql & " and codalmac = " & DBSet(txtAux(7).Text, "N")
    Sql = Sql & " and codvarie = " & DBSet(txtAux(8).Text, "N")
    
    
    If txtAux(9).Text = "" Then
        Sql = Sql & " and codforfait = ' '"
    Else
        Sql = Sql & " and codforfait = " & DBSet(txtAux(9).Text, "T")
    End If
    
    conn.Execute Sql
    
    ModificaDesdeForm = True
    Exit Function
    
eModificaDesdeForm:
    MuestraError Err.Number, "Modificando registro", Err.Description
End Function


Private Function CargarCondicion() As Boolean
Dim Sql As String
Dim NFic As Integer

    On Error GoTo eCargarCondicion

    CargarCondicion = False

    If Dir(App.Path & "\condicionsql.txt", vbArchive) <> "" Then Kill (App.Path & "\condicionsql.txt")
        
    If Me.adodc1.Recordset.RecordCount = 0 Then
        CargarCondicion = True
        Exit Function
    End If
        
    NFic = FreeFile
    
    Open App.Path & "\condicionsql.txt" For Output As #NFic
    
    Print #NFic, adodc1.RecordSource
        
    Close #NFic
    
    CargarCondicion = True
    
    Exit Function
    
eCargarCondicion:
    MuestraError Err.Number, "Cargando condici�n", Err.Description
End Function

Private Sub LeerFiltro(Leer As Boolean)
Dim Sql As String

    Sql = App.Path & "\filtronom.dat"
    If Leer Then
        Filtro = 3
        If Dir(Sql) <> "" Then
            AbrirFicheroFiltro True, Sql
            If IsNumeric(Trim(Sql)) Then Filtro = CByte(Sql)
        End If
    Else
        AbrirFicheroFiltro False, Sql
    End If
End Sub


Private Sub AbrirFicheroFiltro(Leer As Boolean, Fichero As String)
Dim Sql As String
Dim i As Integer

On Error GoTo EAbrir
    i = FreeFile
    If Leer Then
        Open Fichero For Input As #i
        Fichero = "3"
        Line Input #i, Fichero
    Else
        Open Fichero For Output As #i
        Print #i, Filtro
    End If
    Close #i
    Exit Sub
EAbrir:
    Err.Clear
End Sub

Private Function AnyadeCadenaFiltro() As String
Dim Aux As String
'Filtro = 1: a�o actual
'Filtro = 2: a�o actual y anterior
'Filtro = 0: sin filtro
    Aux = ""
    If Filtro < 3 Then
        i = Year(Now)
        If Filtro = 1 Then
            'A�o actual
            Aux = " year(fechahora) >= " & i
        Else
            Aux = " year(fechahora) >=" & i - 1
        End If
    Else
        Aux = "(1=1)"
    End If  'filtro=0
    AnyadeCadenaFiltro = Aux
End Function

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub CargaFiltros()
Dim Aux As String
    
    cboFiltro.Clear
    
    cboFiltro.AddItem "Sin Filtro "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 0
    cboFiltro.AddItem "A�o Actual "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 1
    cboFiltro.AddItem "A�o Actual y Anterior"
    cboFiltro.ItemData(cboFiltro.NewIndex) = 2

End Sub
    
Private Sub CargarSqlFiltro()

    Screen.MousePointer = vbHourglass
    
    cadFiltro = ""
    
    i = Year(Now)
    
    Select Case Me.cboFiltro.ListIndex
        Case -1, 0 ' sin filtro
            cadFiltro = "(1=1)"
        
        Case 1 ' a�o actual
            cadFiltro = " year(fechahora) >= " & i
        
        Case 2 ' a�o actual y anterior
            cadFiltro = " year(fechahora) >=" & i - 1
    
    End Select
    
    Screen.MousePointer = vbDefault


End Sub

