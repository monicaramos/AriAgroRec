VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManHorasETT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrada de Horas ETT"
   ClientHeight    =   9975
   ClientLeft      =   195
   ClientTop       =   180
   ClientWidth     =   16845
   Icon            =   "frmManHorasETT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9975
   ScaleWidth      =   16845
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   90
      TabIndex        =   38
      Top             =   30
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   39
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
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3720
      TabIndex        =   36
      Top             =   30
      Width           =   2415
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   210
         TabIndex        =   37
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
               Object.ToolTipText     =   "Destajo Alicatado"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Penalización"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Bonificación"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Borrado Masivo"
               Object.Tag             =   "2"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "KILOS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   10740
      TabIndex        =   31
      Top             =   9045
      Width           =   2925
      Begin VB.TextBox txtAux2 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Index           =   4
         Left            =   1350
         TabIndex        =   33
         Top             =   180
         Width           =   1350
      End
      Begin VB.TextBox txtAux2 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Index           =   5
         Left            =   1350
         TabIndex        =   32
         Top             =   480
         Width           =   1350
      End
      Begin VB.Label Label4 
         Caption         =   "Alicatados: "
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
         Height          =   225
         Left            =   180
         TabIndex        =   35
         Top             =   210
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Tirón: "
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
         Height          =   225
         Left            =   180
         TabIndex        =   34
         Top             =   510
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "TOTALES"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   2490
      TabIndex        =   24
      Top             =   9045
      Width           =   8205
      Begin VB.TextBox txtAux2 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Index           =   3
         Left            =   6720
         TabIndex        =   27
         Top             =   300
         Width           =   1350
      End
      Begin VB.TextBox txtAux2 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Index           =   1
         Left            =   1110
         TabIndex        =   26
         Top             =   300
         Width           =   1350
      End
      Begin VB.TextBox txtAux2 
         Alignment       =   1  'Right Justify
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
         Height          =   360
         Index           =   2
         Left            =   3960
         TabIndex        =   25
         Top             =   300
         Width           =   1350
      End
      Begin VB.Label Label1 
         Caption         =   "Importe: "
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
         Height          =   225
         Left            =   240
         TabIndex        =   30
         Top             =   330
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "Complemento: "
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
         Height          =   225
         Left            =   2550
         TabIndex        =   29
         Top             =   330
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "Penalización: "
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
         Height          =   225
         Left            =   5400
         TabIndex        =   28
         Top             =   330
         Width           =   1335
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
      Left            =   10800
      MaxLength       =   7
      TabIndex        =   23
      Tag             =   "Kilos Tiron|N|S|||horasett|kilostiron|###,##0||"
      Top             =   4980
      Width           =   450
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
      Left            =   10260
      MaxLength       =   7
      TabIndex        =   7
      Tag             =   "Kilos Alicatados|N|S|||horasett|kilosalicatados|###,##0||"
      Top             =   4980
      Width           =   510
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
      Left            =   2760
      TabIndex        =   22
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
      Index           =   3
      Left            =   2580
      MaskColor       =   &H00000000&
      TabIndex        =   21
      ToolTipText     =   "Buscar variedad"
      Top             =   4950
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
      Index           =   6
      Left            =   1740
      MaxLength       =   6
      TabIndex        =   1
      Tag             =   "Variedad|N|N|||horasett|codvarie|000000|S|"
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
      Left            =   3960
      MaxLength       =   6
      TabIndex        =   2
      Tag             =   "Código|N|N|0|9999|horasett|codigoett|0000|S|"
      Top             =   4980
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
      Index           =   2
      Left            =   12450
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
      Left            =   5430
      MaskColor       =   &H00000000&
      TabIndex        =   19
      ToolTipText     =   "Buscar fecha"
      Top             =   4980
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
      Left            =   11280
      MaxLength       =   8
      TabIndex        =   8
      Tag             =   "Fecha Recibo|F|S|||horasett|fecharec|dd/mm/yyyy||"
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
      Left            =   9420
      MaxLength       =   11
      TabIndex        =   6
      Tag             =   "Penalizacion|N|S|||horasett|penaliza|###,##0.00||"
      Top             =   4980
      Width           =   810
   End
   Begin VB.CheckBox chkAux 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   1
      Left            =   13080
      TabIndex        =   10
      Tag             =   "Int.Contable|N|N|||horasett|intconta|||"
      Top             =   4935
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CheckBox chkAux 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   0
      Left            =   12750
      TabIndex        =   9
      Tag             =   "Int.Contable|N|N|||horasett|pasaridoc|||"
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
      Left            =   8610
      MaxLength       =   11
      TabIndex        =   5
      Tag             =   "Complementos|N|N|||horasett|complemento|###,##0.00||"
      Top             =   4980
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
      Left            =   5520
      TabIndex        =   18
      Top             =   4980
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
      Left            =   1410
      MaskColor       =   &H00000000&
      TabIndex        =   17
      ToolTipText     =   "Buscar fecha"
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
      Index           =   2
      Left            =   7500
      MaxLength       =   11
      TabIndex        =   4
      Tag             =   "Importe|N|N|||horasett|importe|###,##0.00||"
      Top             =   4980
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
      Left            =   14460
      TabIndex        =   11
      Top             =   9360
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
      Left            =   15615
      TabIndex        =   13
      Top             =   9360
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
      Left            =   30
      MaxLength       =   10
      TabIndex        =   0
      Tag             =   "Fecha|F|N|||horasett|fechahora|dd/mm/yyyy|S|"
      Top             =   4920
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
      Left            =   4560
      MaxLength       =   6
      TabIndex        =   3
      Tag             =   "Código|N|N|0|999999|horasett|codcapat|0000|S|"
      Top             =   4980
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmManHorasETT.frx":000C
      Height          =   8145
      Left            =   90
      TabIndex        =   15
      Top             =   780
      Width           =   16600
      _ExtentX        =   29289
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
      Left            =   15630
      TabIndex        =   16
      Top             =   9345
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   1
      Left            =   60
      TabIndex        =   12
      Top             =   9045
      Width           =   2385
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
         Top             =   330
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   5790
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
      Left            =   16290
      TabIndex        =   40
      Top             =   180
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
      Begin VB.Menu mnDestajo 
         Caption         =   "&Destajo Alicatado"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnPenalizacion 
         Caption         =   "&Penalizacion"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnBonificacion 
         Caption         =   "&Bonificacion"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnBarra4 
         Caption         =   "-"
      End
      Begin VB.Menu mnBorradoMasivo 
         Caption         =   "&Borrado Masivo"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Listado Entradas Capataz"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnBarra5 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmManHorasETT"
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
' 3. En la funció BotonAnyadir() canviar la taula i el camp per a SugerirCodigoSiguienteStr
' 4. En la funció BotonBuscar() canviar el nom de la clau primaria
' 5. En la funció BotonEliminar() canviar la pregunta, les descripcions de la
'    variable SQL i el contingut del DELETE
' 6. En la funció PonerLongCampos() posar els camps als que volem canviar el MaxLength quan busquem
' 7. En Form_Load() repasar la barra d'iconos (per si es vol canviar algún) i
'    canviar la consulta per a vore tots els registres
' 8. En Toolbar1_ButtonClick repasar els indexs de cada botó per a que corresponguen
' 9. En la funció CargaGrid canviar l'ORDER BY (normalment per la clau primaria);
'    canviar ademés els noms dels camps, el format i si fa falta la cantitat;
'    repasar els index dels botons modificar i eliminar.
'    NOTA: si en Form_load ya li he posat clausula WHERE, canviar el `WHERE` de
'    `SQL = CadenaConsulta & " WHERE " & vSQL` per un `AND`
' 10. En txtAux_LostFocus canviar el mensage i el format del camp
' 11. En la funció DatosOk() canviar els arguments de DevuelveDesdeBD i el mensage
'    en cas d'error
' 12. En la funció SepuedeBorrar() canviar les comprovacions per a vore si es pot
'    borrar el registre
' *******************************SI N'HI HA COMBO*******************************
' 0. Comprovar que en el SQL de Form_Load() es faça referència a la taula del Combo
' 1. Pegar el Combo1 al  costat dels TextAux. Canviar-li el TAG
' 2. En BotonModificar() canviar el camp del Combo
' 3. En CargaCombo() canviar la consulta i els noms del camps, o posar els valor
'    a ma si no es llig de cap base de datos els valors del Combo

Option Explicit

Private Const IdPrograma = 8015



Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'atre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

Private CadenaConsulta As String
Private CadB As String

Private WithEvents frmCap As frmManCapataz 'mantenimiento de capataces
Attribute frmCap.VB_VarHelpID = -1
Private WithEvents frmMens As frmManTraba 'mantenimiento de trabajadores
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmVar As frmComVar 'variedades
Attribute frmVar.VB_VarHelpID = -1

Dim Modo As Byte
'----------- MODOS --------------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'--------------------------------------------------
Dim PrimeraVez As Boolean
Dim Indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim indCodigo As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim i As Integer

' utilizado para buscar por checks
Private BuscaChekc As String


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
    txtAux2(6).visible = Not B
    
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
    BloquearTxt txtAux(6), (Modo = 4)
    txtAux(5).visible = (Modo = 1)
    BloquearBtn Me.btnBuscar(0), (Modo = 4)
    BloquearBtn Me.btnBuscar(1), (Modo = 4)
    BloquearBtn Me.btnBuscar(3), (Modo = 4)
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
    'borrado masivo
    Toolbar2.Buttons(4).Enabled = B
    Me.mnBorradoMasivo.Enabled = B
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


    'Destajo alicatado
    Toolbar2.Buttons(1).Enabled = (Modo = 2) And Not DeConsulta
    Me.mnDestajo.Enabled = (Modo = 2) And Not DeConsulta

    'penalizacion
    Toolbar2.Buttons(2).Enabled = (Modo = 2) And Not DeConsulta
    Me.mnPenalizacion.Enabled = (Modo = 2) And Not DeConsulta

    'bonificacion
    Toolbar2.Buttons(3).Enabled = (Modo = 2) And Not DeConsulta
    Me.mnBonificacion.Enabled = (Modo = 2) And Not DeConsulta



End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
    CargaGrid CadB, True 'primer de tot carregue tot el grid
'    CadB = ""
   
'    '******************** canviar taula i camp **************************
'    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'        NumF = NuevoCodigo
'    Else
'        NumF = SugerirCodigoSiguienteStr("productos", "codprodu")
'    End If
'    '********************************************************************
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
    txtAux2(0).Text = ""
    txtAux2(6).Text = ""
    
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
    CargaGrid "horasett.codcapat = -1"
    '*******************************************************************************
    'Buscar
    For i = 0 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i
    chkAux(0).Value = 0
    chkAux(1).Value = 0
    Me.txtAux2(0).Text = ""
    Me.txtAux2(6).Text = ""
    
    
    LLamaLineas DataGrid1.Top + 240, 1 'Pone el form en Modo=1, Buscar
    PonerFoco txtAux(1)
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
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top '540 '565 '495 '545
    End If
    
    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(4).Text ' codcapat
    txtAux2(0).Text = DataGrid1.Columns(5).Text ' nomcapat
    txtAux(6).Text = DataGrid1.Columns(1).Text 'codvarie
    txtAux2(6).Text = DataGrid1.Columns(2).Text 'nomvarie
    txtAux(1).Text = DataGrid1.Columns(0).Text 'fechahora
    txtAux(7).Text = DataGrid1.Columns(3).Text 'codigoett
    
    ' ***** canviar-ho pel nom del camp del combo *********
'    SelComboBool DataGrid1.Columns(2).Text, Combo1(0)
    ' *****************************************************
    txtAux(2).Text = DataGrid1.Columns(6).Text 'importe
    txtAux(3).Text = DataGrid1.Columns(7).Text 'complemento
    txtAux(4).Text = DataGrid1.Columns(8).Text 'penalizacion
    txtAux(8).Text = DataGrid1.Columns(9).Text ' kilos alicatados
    txtAux(9).Text = DataGrid1.Columns(10).Text ' kilos alicatados
    
    txtAux(5).Text = DataGrid1.Columns(11).Text 'fecharecep
    
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
    txtAux2(6).Top = alto
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
    Sql = "¿Seguro que desea eliminar el Registro?"
    Sql = Sql & vbCrLf & "Fecha: " & adodc1.Recordset.Fields(0)
    Sql = Sql & vbCrLf & "Variedad: " & adodc1.Recordset.Fields(1) & " " & adodc1.Recordset.Fields(2)
    Sql = Sql & vbCrLf & "Codigo ETT: " & adodc1.Recordset.Fields(3)
    Sql = Sql & vbCrLf & "Capataz: " & adodc1.Recordset.Fields(4) & " " & adodc1.Recordset.Fields(5)
    
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        Sql = "Delete from horasett where codcapat=" & adodc1.Recordset!codcapat
        Sql = Sql & " and fechahora = " & DBSet(adodc1.Recordset!FechaHora, "F")
        Sql = Sql & " and codigoett = " & DBLet(adodc1.Recordset!CodigoETT)
        Sql = Sql & " and codvarie = " & DBLet(adodc1.Recordset!Codvarie, "N")
        
        
        conn.Execute Sql
        CargaGrid CadB
'        If CadB <> "" Then
'            CargaGrid CadB
'            lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
'        Else
'            CargaGrid ""
'            lblIndicador.Caption = ""
'        End If
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
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los txtAux
    PonerLongCamposGnral Me, Modo, 3
End Sub

Private Sub btnBuscar_Click(Index As Integer)
 TerminaBloquear
    
    Select Case Index
        Case 1 'capataces
            Indice = 0
            Set frmCap = New frmManCapataz
            frmCap.DatosADevolverBusqueda = "0|1|"
            frmCap.Show vbModal
            Set frmCap = Nothing
            PonerFoco txtAux(Indice)
    
       Case 0, 2 ' Fecha
            Dim esq As Long
            Dim dalt As Long
            Dim menu As Long
            Dim obj As Object
        
            Set frmC = New frmCal
            
            
            esq = btnBuscar(Index).Left
            dalt = btnBuscar(Index).Top
                
            Set obj = btnBuscar(Index).Container
              
              While btnBuscar(Index).Parent.Name <> obj.Name
                    esq = esq + obj.Left
                    dalt = dalt + obj.Top
                    Set obj = obj.Container
              Wend
            
            menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
        
            frmC.Left = esq + btnBuscar(Index).Parent.Left + 30
            frmC.Top = dalt + btnBuscar(Index).Parent.Top + btnBuscar(Index).Height + menu - 40
        
            btnBuscar(Index).Tag = Index '<===
            ' *** repasar si el camp es txtAux o Text1 ***
            If Index = 0 Then
                Indice = 1
                If txtAux(1).Text <> "" Then frmC.NovaData = txtAux(1).Text
            Else
                Indice = 5
                If txtAux(5).Text <> "" Then frmC.NovaData = txtAux(5).Text
            End If
            
            ' ********************************************
        
            frmC.Show vbModal
            Set frmC = Nothing
            ' *** repasar si el camp es txtAux o Text1 ***
            If Index = 0 Then
                PonerFoco txtAux(1) '<===
            Else
                PonerFoco txtAux(5) '<===
            End If
            ' ********************************************
     
        Case 3 'codigo de variedad
            AbrirFrmVariedades 6

    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.adodc1, 1
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
                    CargaGrid CadB
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
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
                    i = adodc1.Recordset.Fields(1)
                    PonerModo 2
                    CargaGrid CadB
'                    If CadB <> "" Then
'                        CargaGrid CadB
'                        lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
'                    Else
'                        CargaGrid
'                        lblIndicador.Caption = ""
'                    End If
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(1).Name & " =" & i)
                    PonerFocoGrid Me.DataGrid1
                End If
            End If
    End Select
End Sub

Private Sub cmdCancelar_Click()
    On Error Resume Next
    
    Select Case Modo
        Case 1 'búsqueda
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

'    With Me.Toolbar1
'        .HotImageList = frmPpal.imgListComun_OM
'        .DisabledImageList = frmPpal.imgListComun_BN
'        .ImageList = frmPpal.imgListComun
'        'el 1 es separadors
'        .Buttons(2).Image = 1   'Buscar
'        .Buttons(3).Image = 2   'Todos
'        'el 4 i el 5 son separadors
'        .Buttons(6).Image = 3   'Insertar
'        .Buttons(7).Image = 4   'Modificar
'        .Buttons(8).Image = 5   'Borrar
'        'el 9 i el 10 son separadors
'        .Buttons(10).Image = 31  'destajo alicatado
'        .Buttons(11).Image = 32  'penalizacion
'        .Buttons(12).Image = 26  'bonificacion
'
'        .Buttons(14).Image = 28  'borrado masivo
'        .Buttons(15).Image = 10  'Salir
'        .Buttons(16).Image = 11  'Salir
'    End With

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
    
    With Me.Toolbar2
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 31  'destajo alicatado
        .Buttons(2).Image = 32  'penalizacion
        .Buttons(3).Image = 26  'bonificacion
        .Buttons(4).Image = 28  'borrado masivo
    End With

    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 12
    End With





    '## A mano
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    
    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT horasett.fechahora, horasett.codvarie, variedades.nomvarie, horasett.codigoett,"
    CadenaConsulta = CadenaConsulta & " horasett.codcapat, rcapataz.nomcapat, horasett.importe, horasett.complemento,  "
    CadenaConsulta = CadenaConsulta & " horasett.penaliza, horasett.kilosalicatados, horasett.kilostiron, "
    CadenaConsulta = CadenaConsulta & " horasett.fecharec, "
    CadenaConsulta = CadenaConsulta & " horasett.pasaridoc,  IF(pasaridoc=1,'*','') as pasari, horasett.intconta,  IF(intconta=1,'*','') as intcon "
    CadenaConsulta = CadenaConsulta & " FROM rcapataz, variedades, horasett  "
    CadenaConsulta = CadenaConsulta & " WHERE horasett.codcapat = rcapataz.codcapat and "
'    CadenaConsulta = CadenaConsulta & " horas.codforfait = forfaits.codforfait and "
    CadenaConsulta = CadenaConsulta & " horasett.codvarie = variedades.codvarie "
    '************************************************************************
    
    CadB = ""
    CargaGrid "horasett.codcapat = -1"
    
'    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'        BotonAnyadir
'    Else
'        PonerModo 2
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Modo = 4 Then TerminaBloquear
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


Private Sub frmCap_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo capataz
    txtAux2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre capataz
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo variedad
    txtAux2(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre variedad
End Sub

Private Sub mnBonificacion_Click()
    BotonBonificacion
End Sub

Private Sub mnBorradoMasivo_Click()
    BotonBorradoMasivo
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnDestajo_Click()
    BotonDestajo
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    AbrirListadoNominas (29)
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

Private Sub mnPenalizacion_Click()
    BotonPenalizacion
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 5
                mnBuscar_Click
        Case 6
                mnVerTodos_Click
        Case 1
                mnNuevo_Click
        Case 2
                mnModificar_Click
        Case 3
                mnEliminar_Click
        Case 8 ' imprimir
                mnImprimir_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String, Optional Ascendente As Boolean)
    Dim Sql As String
    Dim tots As String
    
'    adodc1.ConnectionString = Conn
    If vSQL <> "" Then
        Sql = CadenaConsulta & " AND " & vSQL
    Else
        Sql = CadenaConsulta
    End If
    If Ascendente Then
        Sql = Sql & " ORDER BY  horasett.fechahora, horasett.codvarie "
    Else
        '********************* canviar el ORDER BY *********************++
        Sql = Sql & " ORDER BY  horasett.fechahora desc, horasett.codvarie, horasett.codigoett, horasett.codcapat "
        '**************************************************************++
    End If
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, Sql, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "S|txtAux(1)|T|Fecha|1400|;S|btnBuscar(0)|B||195|;"
    tots = tots & "S|txtAux(6)|T|Codigo|1000|;S|btnBuscar(3)|B||195|;S|txtAux2(6)|T|Variedad|1600|;"
    tots = tots & "S|txtAux(7)|T|Codigo ETT|1200|;"
    tots = tots & "S|txtAux(0)|T|Codigo|1000|;S|btnBuscar(1)|B||195|;S|txtAux2(0)|T|Capataz|1400|;"
    tots = tots & "S|txtAux(2)|T|Importe|1200|;"
    tots = tots & "S|txtAux(3)|T|Complemento|1500|;"
    tots = tots & "S|txtAux(4)|T|Penalización|1400|;S|txtAux(8)|T|Kilos Alic.|1100|;S|txtAux(9)|T|Kilos Tiron|1100|;"
    tots = tots & "S|txtAux(5)|T|F.Recibo|1400|;S|btnBuscar(2)|B||195|;N||||0|;S|chkAux(0)|CB|IA|360|;N||||0|;S|chkAux(1)|CB|IC|360|;"
    
    arregla tots, DataGrid1, Me, 350
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgLeft
    DataGrid1.Columns(3).Alignment = dbgLeft
'    DataGrid1.Columns(10).Alignment = dbgCenter
'    DataGrid1.Columns(12).Alignment = dbgCenter

    CalcularTotales Sql


End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 ' destajo
                mnDestajo_Click
        Case 2 ' penalizacion
                mnPenalizacion_Click
        Case 3 ' bonificacion
                mnBonificacion_Click
        Case 4
                mnBorradoMasivo_Click
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
                txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "rcapataz", "nomcapat")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe el Capataz: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCap = New frmManCapataz
                        frmCap.DatosADevolverBusqueda = "0|1|"
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmCap.Show vbModal
                        Set frmCap = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
            End If
        
        Case 1, 5 'fecha
            '[Monica]28/08/2013: comprobamos que la fecha esté en la campaña
            PonerFormatoFecha txtAux(Index), True
            
        Case 2, 3, 4 'importe, complemento, penalizacion
            If txtAux(Index).Text <> "" Then
                cadMen = TransformaPuntosComas(txtAux(Index).Text)
                txtAux(Index).Text = Format(cadMen, "###,##0.00")
            End If
    
        Case 6 'codigo de variedad
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
    
        Case 8 ' kilos alicatados
            PonerFormatoEntero txtAux(Index)
    
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
        Sql = "select count(*) from horasett where codcapat = " & DBSet(txtAux(0).Text, "N")
        Sql = Sql & " and fechahora = " & DBSet(txtAux(1).Text, "F")
        Sql = Sql & " and codigoett = " & DBSet(txtAux(7).Text, "N")
        Sql = Sql & " and codvarie = " & DBSet(txtAux(6).Text, "N")
        If TotalRegistros(Sql) <> 0 Then
            MsgBox "El capataz existe para esta fecha, codigoett, variedad. Reintroduzca.", vbExclamation
            PonerFoco txtAux(0)
            B = False
        End If
    End If
    
    If B And (Modo = 3 Or Modo = 4) Then
        If Not EntreFechas(vParam.FecIniCam, txtAux(1).Text, vParam.FecFinCam) Then
            MsgBox "La fecha introducida no se encuentra dentro de campaña. Revise.", vbExclamation
            B = False
        End If
    End If
    
    DatosOK = B
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub



'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGrid1_GotFocus()
'  WheelHook DataGrid1
'End Sub
'Private Sub DataGrid1_Lostfocus()
'  WheelUnHook
'End Sub

'Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
'    KEYpress KeyAscii
'End Sub
Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 1: KEYBusqueda KeyAscii, 0 'fecha
                Case 6: KEYBusqueda KeyAscii, 3 'variedad
                Case 0: KEYBusqueda KeyAscii, 1 'capataz
                Case 5: KEYBusqueda KeyAscii, 2 'fecha
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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
    btnBuscar_Click (Indice)
End Sub

Private Sub BotonDestajo()
    AbrirListadoNominas (20)
    CargaGrid
End Sub

Private Sub BotonPenalizacion()
    AbrirListadoNominas (21)
    CargaGrid
End Sub

Private Sub BotonBonificacion()
    AbrirListadoNominas (22)
    CargaGrid
End Sub

Private Sub BotonBorradoMasivo()
    AbrirListadoNominas (23)
    CargaGrid
End Sub

Private Sub AbrirFrmVariedades(Indice As Integer)
    indCodigo = 6
    Set frmVar = New frmComVar
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.CodigoActual = txtAux(indCodigo)
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub


Private Function ModificaDesdeForm() As Boolean
Dim Sql As String

    On Error GoTo eModificaDesdeForm
    
    ModificaDesdeForm = False
    
    Sql = "update horasett set "
    Sql = Sql & " importe = " & DBSet(ImporteSinFormato(txtAux(2).Text), "N")
    Sql = Sql & ", complemento = " & DBSet(ImporteSinFormato(txtAux(3).Text), "N")
    Sql = Sql & ", penaliza = " & DBSet(ImporteSinFormato(txtAux(4).Text), "N")
    Sql = Sql & " where codcapat = " & DBSet(txtAux(0).Text, "N")
    Sql = Sql & " and fechahora = " & DBSet(txtAux(1).Text, "F")
    Sql = Sql & " and codigoett = " & DBSet(txtAux(7).Text, "N")
    Sql = Sql & " and codvarie = " & DBSet(txtAux(6).Text, "N")
    
    conn.Execute Sql
    
    ModificaDesdeForm = True
    Exit Function
    
eModificaDesdeForm:
    MuestraError Err.Number, "Modificando registro", Err.Description
End Function


Private Sub CalcularTotales(cadena As String)
Dim Importe  As Currency
Dim Compleme As Currency
Dim Penaliza As Currency
Dim KilosAlicatados As Long
Dim KilosTiron As Long

Dim Rs As ADODB.Recordset
Dim Sql As String

    On Error Resume Next
    
    Sql = "select sum(importe) importe , sum(complemento) compleme, sum(penaliza) penaliza, sum(kilosalicatados) kilosalicatados, sum(kilostiron) kilostiron from (" & cadena & ") aaaaa"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Importe = 0
    Compleme = 0
    Penaliza = 0
    txtAux2(1).Text = ""
    txtAux2(2).Text = ""
    txtAux2(3).Text = ""
    txtAux2(4).Text = ""
    txtAux2(5).Text = ""
    
    If TotalRegistrosConsulta(cadena) = 0 Then Exit Sub
    
    If Not Rs.EOF Then
        If Rs.Fields(0).Value <> 0 Then Importe = DBLet(Rs.Fields(0).Value, "N") 'Solo es para saber que hay registros que mostrar
        If Rs.Fields(1).Value <> 0 Then Compleme = DBLet(Rs.Fields(1).Value, "N") 'Solo es para saber que hay registros que mostrar
        If Rs.Fields(2).Value <> 0 Then Penaliza = DBLet(Rs.Fields(2).Value, "N") 'Solo es para saber que hay registros que mostrar
        If Rs.Fields(3).Value <> 0 Then KilosAlicatados = DBLet(Rs.Fields(3).Value, "N") 'Solo es para saber que hay registros que mostrar
        If Rs.Fields(4).Value <> 0 Then KilosTiron = DBLet(Rs.Fields(4).Value, "N") 'Solo es para saber que hay registros que mostrar
    
        txtAux2(1).Text = Format(Importe, "###,###,##0.00")
        txtAux2(2).Text = Format(Compleme, "###,###,##0.00")
        txtAux2(3).Text = Format(Penaliza, "###,###,##0.00")
        txtAux2(4).Text = Format(KilosAlicatados, "###,###,##0")
        txtAux2(5).Text = Format(KilosTiron, "###,###,##0")
    End If
    Rs.Close
    Set Rs = Nothing

    
    DoEvents
    
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

