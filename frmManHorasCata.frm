VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManHorasCata 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrada de Horas Catadau"
   ClientHeight    =   10695
   ClientLeft      =   195
   ClientTop       =   180
   ClientWidth     =   19005
   Icon            =   "frmManHorasCata.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10695
   ScaleWidth      =   19005
   StartUpPosition =   2  'CenterScreen
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
      Left            =   4005
      TabIndex        =   52
      Top             =   10280
      Visible         =   0   'False
      Width           =   4785
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
      Index           =   6
      Left            =   15750
      MaskColor       =   &H00000000&
      TabIndex        =   51
      ToolTipText     =   "Buscar campo"
      Top             =   4635
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
      Index           =   13
      Left            =   10755
      MaxLength       =   11
      TabIndex        =   9
      Tag             =   "Importe Kilómetros|N|S|||horas|importekms|###,##0.00||"
      Top             =   4590
      Width           =   810
   End
   Begin VB.CheckBox chkAux 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   1
      Left            =   13545
      TabIndex        =   48
      Tag             =   "Es Horas Capataz|N|N|||horas|escapataz|||"
      Top             =   4620
      Visible         =   0   'False
      Width           =   255
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
      Index           =   4
      Left            =   1935
      TabIndex        =   46
      Top             =   9135
      Width           =   510
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
      Index           =   12
      Left            =   8685
      MaxLength       =   7
      TabIndex        =   6
      Tag             =   "Kilos|N|S|||horas|kilos|###,##0||"
      Top             =   4590
      Width           =   600
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
      Index           =   11
      Left            =   6660
      MaxLength       =   4
      TabIndex        =   4
      Tag             =   "Código|N|N|0|9999|horas|codcateg|0000|N|"
      Top             =   4590
      Width           =   800
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
      Left            =   7530
      MaskColor       =   &H00000000&
      TabIndex        =   45
      ToolTipText     =   "Buscar categoria"
      Top             =   4590
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
      Index           =   11
      Left            =   4005
      TabIndex        =   43
      Top             =   9900
      Width           =   4785
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
      Index           =   10
      Left            =   14535
      MaxLength       =   10
      TabIndex        =   40
      Tag             =   "Nro Parte|N|S|||horas|nroparte|0000000||"
      Top             =   4635
      Width           =   1170
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   90
      TabIndex        =   37
      Top             =   0
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   38
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
      Left            =   3720
      TabIndex        =   35
      Top             =   0
      Width           =   1545
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   210
         TabIndex        =   36
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
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Alta Rápida"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Eventuales"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Trabajadores de Capataz"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Anticipos"
               Object.Tag             =   "2"
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
      Left            =   11160
      MaxLength       =   11
      TabIndex        =   10
      Tag             =   "Penalizacion|N|S|||horas|penaliza|###,##0.00||"
      Top             =   4590
      Width           =   810
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
      Left            =   7830
      MaxLength       =   6
      TabIndex        =   5
      Tag             =   "Horas Dia|N|S|||horas|horasdia|##0.00||"
      Top             =   4590
      Width           =   780
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
      Height          =   1275
      Left            =   9045
      TabIndex        =   28
      Top             =   9000
      Width           =   6765
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
         Index           =   5
         Left            =   3375
         TabIndex        =   49
         Top             =   720
         Width           =   1560
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
         Index           =   3
         Left            =   4995
         TabIndex        =   33
         Top             =   720
         Width           =   1560
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
         Left            =   1755
         TabIndex        =   30
         Top             =   720
         Width           =   1560
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
         Left            =   135
         TabIndex        =   29
         Top             =   720
         Width           =   1560
      End
      Begin VB.Label Label8 
         Caption         =   "Kilómetros: "
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
         Left            =   3375
         TabIndex        =   50
         Top             =   360
         Width           =   1335
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
         Left            =   4995
         TabIndex        =   34
         Top             =   360
         Width           =   1335
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
         Left            =   1755
         TabIndex        =   32
         Top             =   360
         Width           =   1395
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
         Left            =   180
         TabIndex        =   31
         Top             =   375
         Width           =   945
      End
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
      Left            =   3585
      TabIndex        =   26
      Top             =   4590
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
      Index           =   4
      Left            =   3345
      MaskColor       =   &H00000000&
      TabIndex        =   25
      ToolTipText     =   "Buscar trabajador"
      Top             =   4590
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
      Left            =   4005
      TabIndex        =   24
      Top             =   9135
      Width           =   4785
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
      TabIndex        =   23
      ToolTipText     =   "Buscar variedad"
      Top             =   4560
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
      Tag             =   "Variedad|N|N|||horas|codvarie|000000|S|"
      Top             =   4560
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
      Left            =   2835
      MaxLength       =   6
      TabIndex        =   2
      Tag             =   "Código|N|N|0|999999|horas|codtraba|000000|S|"
      Top             =   4590
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
      Left            =   13320
      MaskColor       =   &H00000000&
      TabIndex        =   22
      ToolTipText     =   "Buscar fecha"
      Top             =   4590
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
      Left            =   6375
      MaskColor       =   &H00000000&
      TabIndex        =   21
      ToolTipText     =   "Buscar capataz"
      Top             =   4560
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
      Left            =   12060
      MaxLength       =   10
      TabIndex        =   11
      Tag             =   "Fecha Recibo|F|S|||horas|fecharec|dd/mm/yyyy||"
      Top             =   4590
      Width           =   1170
   End
   Begin VB.CheckBox chkAux 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   0
      Left            =   14205
      TabIndex        =   12
      Tag             =   "Int.Contable|N|N|||horas|intconta|||"
      Top             =   4620
      Visible         =   0   'False
      Width           =   225
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
      Left            =   10035
      MaxLength       =   11
      TabIndex        =   8
      Tag             =   "Complementos|N|S|||horas|compleme|###,##0.00||"
      Top             =   4590
      Width           =   1020
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
      Left            =   4005
      TabIndex        =   20
      Top             =   9510
      Width           =   4785
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
      TabIndex        =   19
      ToolTipText     =   "Buscar fecha"
      Top             =   4530
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
      Left            =   9360
      MaxLength       =   11
      TabIndex        =   7
      Tag             =   "Importe|N|N|||horas|importe|###,##0.00||"
      Top             =   4590
      Width           =   675
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
      Left            =   16650
      TabIndex        =   13
      Top             =   9855
      Visible         =   0   'False
      Width           =   1035
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
      Left            =   17820
      TabIndex        =   15
      Top             =   9855
      Visible         =   0   'False
      Width           =   1095
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
      Tag             =   "Fecha|F|N|||horas|fechahora|dd/mm/yyyy|S|"
      Top             =   4530
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
      Left            =   5505
      MaxLength       =   4
      TabIndex        =   3
      Tag             =   "Código|N|N|0|9999|horas|codcapat|0000|N|"
      Top             =   4560
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmManHorasCata.frx":000C
      Height          =   8145
      Left            =   90
      TabIndex        =   17
      Top             =   780
      Width           =   18805
      _ExtentX        =   33179
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
      Left            =   17820
      TabIndex        =   18
      Top             =   9855
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Index           =   1
      Left            =   60
      TabIndex        =   14
      Top             =   9630
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
         Left            =   40
         TabIndex        =   16
         Top             =   240
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
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   8
      Left            =   13980
      MaxLength       =   6
      TabIndex        =   27
      Tag             =   "Código|N|N|0|99|horas|codalmac|00|S|"
      Top             =   4080
      Width           =   465
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   18405
      TabIndex        =   39
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
   Begin VB.Label Label9 
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
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2880
      TabIndex        =   53
      Top             =   10280
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label7 
      Caption         =   "Grupo Trabajador"
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
      Height          =   270
      Left            =   135
      TabIndex        =   47
      Top             =   9180
      Width           =   2025
   End
   Begin VB.Label Label6 
      Caption         =   "Categoria"
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
      Height          =   270
      Left            =   2880
      TabIndex        =   44
      Top             =   9900
      Width           =   945
   End
   Begin VB.Label Label5 
      Caption         =   "Capataz"
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
      Left            =   2880
      TabIndex        =   42
      Top             =   9540
      Width           =   945
   End
   Begin VB.Label Label4 
      Caption         =   "Variedad"
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
      Left            =   2880
      TabIndex        =   41
      Top             =   9180
      Width           =   945
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
      Begin VB.Menu mnAltaRapida 
         Caption         =   "&Alta Rápida"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnEventuales 
         Caption         =   "E&ventuales"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnTrabajadores 
         Caption         =   "&Trabajadores de Capataz"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnBarra4 
         Caption         =   "-"
      End
      Begin VB.Menu mnBorradoMasivo 
         Caption         =   "&Borrado Masivo"
         Shortcut        =   ^O
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
Attribute VB_Name = "frmManHorasCata"
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

Private Const IdPrograma = 8010



Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'atre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

Private CadenaConsulta As String
Private cadB As String

Private WithEvents frmCap As frmManCapataz 'mantenimiento de capataces
Attribute frmCap.VB_VarHelpID = -1
Private WithEvents frmTra As frmManTraba 'mantenimiento de trabajadores
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmCat As frmManCategorias 'mantenimiento de categorias
Attribute frmCat.VB_VarHelpID = -1
Private WithEvents frmCam As frmManCampos ' mantenimiento de campos
Attribute frmCam.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmVar As frmBasico2 'ComVar 'variedades
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
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim indCodigo As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim i As Integer

' utilizado para buscar por checks
Private BuscaChekc As String


Private Sub PonerModo(vModo)
Dim b As Boolean

    Modo = vModo
    
    b = (Modo = 2)
    If b Then
        PonerContRegIndicador lblIndicador, adodc1, cadB
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    For i = 0 To txtAux.Count - 1
        txtAux(i).visible = Not b
        txtAux(i).BackColor = vbWhite
    Next i
    
'    txtAux2(0).visible = Not B
'    txtAux2(6).visible = Not B
    txtAux2(7).visible = Not b
    
    For i = 0 To btnBuscar.Count - 1
        btnBuscar(i).visible = Not b
    Next i
    If vParamAplic.Cooperativa <> 18 Then btnBuscar(6).visible = False
    
    chkAux(0).visible = Not b
    chkAux(1).visible = Not b

    CmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    DataGrid1.Enabled = b
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = b
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu  'En funcion del usuario
    
    'Si estamos modo Modificar bloquear clave primaria
    BloquearTxt txtAux(1), (Modo = 4)
    BloquearTxt txtAux(7), (Modo = 4)
    BloquearTxt txtAux(6), (Modo = 4)
    ' la fecha la dejamos poner y modificar pq ellos no imprimen recibos
    'txtAux(5).visible = (Modo = 1)
    BloquearBtn Me.btnBuscar(0), (Modo = 4)
    BloquearBtn Me.btnBuscar(4), (Modo = 4)
    BloquearBtn Me.btnBuscar(3), (Modo = 4)
    
    If vParamAplic.Cooperativa = 18 Then BloquearBtn Me.btnBuscar(6), (Modo = 4)
    
' la fecha la dejamos poner y modificar pq ellos no imprimen recibos
'    BloquearBtn Me.btnBuscar(2), (Modo = 4) Or (Modo = 3)
    
    BloquearChk Me.chkAux(0), (Modo = 4) Or (Modo = 3)
    Me.chkAux(0).visible = (Modo = 1)
    
    '[Monica]23/05/2018: si es o no capataz para el calculo de plus de capataz en el proceso de nomina
    Me.chkAux(1).visible = ((Modo = 1) Or (Modo = 3) Or (Modo = 4))
    
    ' fecha de recibo
    BloquearTxt txtAux(5), (Modo <> 1)
    BloquearBtn btnBuscar(2), (Modo <> 1)
    
    If vParamAplic.Cooperativa = 18 Then
        ' para el caso de frutas inma  es el txtaux(10) es el campo
    Else
        'El nro de parte unicamente lo podemos buscar
        txtAux(10).Enabled = (Modo = 1)
        txtAux(10).visible = (Modo = 1)
    End If
End Sub


Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botones del la toobar y del menu, segun el modo en que estemos
Dim b As Boolean

    b = (Modo = 2)
    'Busqueda
    Toolbar1.Buttons(5).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = b
    Me.mnVerTodos.Enabled = b
    'Imprimir
    Toolbar1.Buttons(8).Enabled = b
    Me.mnBorradoMasivo.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(1).Enabled = b And Not DeConsulta
    Me.mnNuevo.Enabled = b And Not DeConsulta
    
    b = (b And adodc1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    'Eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b


    'alta rapida
    Toolbar2.Buttons(1).Enabled = (Modo = 2) And Not DeConsulta
    Me.mnAltaRapida.Enabled = (Modo = 2) And Not DeConsulta

    'eventuales
    Toolbar2.Buttons(2).Enabled = (Modo = 2) And Not DeConsulta
    Me.mnEventuales.Enabled = (Modo = 2) And Not DeConsulta

    'trabajadores
    Toolbar2.Buttons(3).Enabled = (Modo = 2) And Not DeConsulta
    Me.mnTrabajadores.Enabled = (Modo = 2) And Not DeConsulta

End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
    CargaGrid cadB, True 'primer de tot carregue tot el grid
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
    txtAux2(7).Text = ""
    txtAux2(8).Text = ""
    txtAux2(11).Text = ""
    
    txtAux(8).Text = vParamAplic.AlmacenNOMI ' pq es clave primaria
    
    chkAux(0).Value = 0
    chkAux(1).Value = 0
    
    txtAux(2).Text = 0
    txtAux(3).Text = 0
    txtAux(4).Text = 0
    txtAux(9).Text = 0
    '[Monica]07/06/2018: importe kms
    txtAux(13).Text = 0
    
    LLamaLineas anc, 3 'Pone el form en Modo=3, Insertar
       
    'Ponemos el foco
    PonerFoco txtAux(1)
End Sub

Private Sub BotonVerTodos()
    cadB = ""
    CargaGrid ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    ' ***************** canviar per la clau primaria ********
    CargaGrid "horas.codcapat = -1"
    '*******************************************************************************
    'Buscar
    For i = 0 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i
    chkAux(0).Value = 0
    chkAux(1).Value = 0
    Me.txtAux2(0).Text = ""
    Me.txtAux2(1).Text = ""
    Me.txtAux2(2).Text = ""
    Me.txtAux2(5).Text = ""
    Me.txtAux2(6).Text = ""
    Me.txtAux2(7).Text = ""
    
    
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
    txtAux(0).Text = DataGrid1.Columns(5).Text ' codcapat
    txtAux2(0).Text = DataGrid1.Columns(6).Text ' nomcapat
    txtAux(6).Text = DataGrid1.Columns(1).Text 'codvarie
    txtAux2(6).Text = DataGrid1.Columns(2).Text 'nomvarie
    txtAux(1).Text = DataGrid1.Columns(0).Text 'fechahora
    txtAux(7).Text = DataGrid1.Columns(3).Text 'codtraba
    txtAux2(7).Text = DataGrid1.Columns(4).Text 'nomtraba
    
    txtAux(8).Text = vParamAplic.AlmacenNOMI ' pq es clave primaria

    ' ***** canviar-ho pel nom del camp del combo *********
'    SelComboBool DataGrid1.Columns(2).Text, Combo1(0)
    ' *****************************************************
    txtAux(11).Text = DataGrid1.Columns(7).Text 'categoria
    txtAux(12).Text = DataGrid1.Columns(10).Text 'importe
    
    txtAux(2).Text = DataGrid1.Columns(11).Text 'importe
    txtAux(3).Text = DataGrid1.Columns(12).Text 'complemento
    txtAux(13).Text = DataGrid1.Columns(13).Text 'importe kms
    txtAux(4).Text = DataGrid1.Columns(9).Text 'horas
    txtAux(9).Text = DataGrid1.Columns(14).Text 'penalizacion
    txtAux(5).Text = DataGrid1.Columns(15).Text 'fecharecep
    
    Me.chkAux(0).Value = Me.adodc1.Recordset!intconta
    Me.chkAux(1).Value = Me.adodc1.Recordset!escapataz
    

    LLamaLineas anc, 4 'Pone el form en Modo=4, Modificar
   
    'Como es modificar
    PonerFoco txtAux(0)
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
'    txtAux2(0).Top = alto
'    txtAux2(6).Top = alto
    txtAux2(7).Top = alto
    For i = 0 To btnBuscar.Count - 1
        btnBuscar(i).Top = alto - 15
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
    
    If Not SepuedeBorrar Then Exit Sub
        
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCampo(txtAux(0))) Then Exit Sub
    ' ***************************************************************************
    
    '*************** canviar els noms i el DELETE **********************************
    Sql = "¿Seguro que desea eliminar el Registro?"
    Sql = Sql & vbCrLf & "Fecha: " & adodc1.Recordset.Fields(0)
    Sql = Sql & vbCrLf & "Variedad: " & adodc1.Recordset.Fields(1) & " " & adodc1.Recordset.Fields(2)
    Sql = Sql & vbCrLf & "Trabajador: " & adodc1.Recordset.Fields(3)
    Sql = Sql & vbCrLf & "Capataz: " & adodc1.Recordset.Fields(5) & " " & adodc1.Recordset.Fields(6)
    
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        Sql = "Delete from horas where codcapat=" & adodc1.Recordset!codcapat
        Sql = Sql & " and fechahora = " & DBSet(adodc1.Recordset!FechaHora, "F")
        Sql = Sql & " and codtraba = " & DBLet(adodc1.Recordset!CodTraba)
        Sql = Sql & " and codvarie = " & DBLet(adodc1.Recordset!codvarie, "N")
        
        
        conn.Execute Sql
        CargaGrid cadB
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

Private Function SepuedeBorrar() As Boolean


    SepuedeBorrar = (DBLet(Me.adodc1.Recordset!Fecharec, "F") = "" And Me.adodc1.Recordset!intconta = 0)

End Function

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
            AbrirFrmCapataz 0
    
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
                indice = 1
                If txtAux(1).Text <> "" Then frmC.NovaData = txtAux(1).Text
            Else
                indice = 5
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
            
        Case 4 ' codigo de trabajador
            AbrirFrmTrabajador 7
        
        Case 5
            AbrirFrmCategoria 11
            
        Case 6 ' codigo de campo
            AbrirFrmCampos 10
            
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
            cadB = ObtenerBusqueda3(Me, False, BuscaChekc)
            If cadB <> "" Then
                CargaGrid cadB
                PonerModo 2
'                lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
                PonerFocoGrid Me.DataGrid1
            End If
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    CargaGrid cadB
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
                    cadB = ""
                End If
            End If
            
        Case 4 'MODIFICAR
            If DatosOk Then
                If ModificaDesdeForm Then
                    TerminaBloquear
                    PosicionarData
                End If
            End If
    End Select
End Sub

Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    Cad = "codtraba = " & adodc1.Recordset.Fields(3) & " and  codvarie = " & adodc1.Recordset.Fields(1) & " and fechahora = " & DBSet(adodc1.Recordset.Fields(0), "F")
    ' ***************************************
    CargaGrid cadB
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarDataMULTI(adodc1, Cad, Indicador, True) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       PonerModo 0
    End If
   
    ' ***********************************************************************************
End Sub


Private Sub cmdCancelar_Click()
    On Error Resume Next
    
    Select Case Modo
        Case 1 'búsqueda
            CargaGrid cadB
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
Dim Cad As String
Dim i As Integer
Dim J As Integer
Dim Aux As String

    If adodc1.Recordset.EOF Then
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
            Cad = Cad & adodc1.Recordset.Fields(J) & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Modo = 2 Then
        PonerContRegIndicador lblIndicador, adodc1, cadB
        CargaForaGrid
    End If
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
        .Buttons(4).Image = 13 ' 28  'ahora anticipos antes borrado masivo borrado masivo
    End With

    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 12
    End With

    '[Monica]14/12/2017: para el caso de Catadau
    If vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19 Then
        Me.Caption = "Entrada de horas"
    End If


    '## A mano
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT horas.fechahora, horas.codvarie, variedades.nomvarie, horas.codtraba, straba.nomtraba, "
    CadenaConsulta = CadenaConsulta & " horas.codcapat, rcapataz.nomcapat, "
    
    If vParamAplic.Cooperativa = 18 Then
        CadenaConsulta = CadenaConsulta & " horas.codcampo, "
        CadenaConsulta = CadenaConsulta & " horas.codcateg , rcategorias.nomcateg, "
        CadenaConsulta = CadenaConsulta & " horas.horasdia, horas.kilos, horas.importe, horas.compleme, horas.importekms, horas.penaliza, "
        CadenaConsulta = CadenaConsulta & " horas.fecharec, "
        CadenaConsulta = CadenaConsulta & " horas.escapataz,  IF(escapataz=1,'*','') as escapa, "
        CadenaConsulta = CadenaConsulta & " horas.intconta,  IF(intconta=1,'*','') as intcon, "
        CadenaConsulta = CadenaConsulta & " horas.codalmac "
    
    Else
    
        CadenaConsulta = CadenaConsulta & " horas.codcateg , rcategorias.nomcateg, "
        CadenaConsulta = CadenaConsulta & " horas.horasdia, horas.kilos, horas.importe, horas.compleme, horas.importekms, horas.penaliza, "
        CadenaConsulta = CadenaConsulta & " horas.fecharec, "
        CadenaConsulta = CadenaConsulta & " horas.escapataz,  IF(escapataz=1,'*','') as escapa, "
        CadenaConsulta = CadenaConsulta & " horas.intconta,  IF(intconta=1,'*','') as intcon, "
        CadenaConsulta = CadenaConsulta & " horas.codalmac, horas.nroparte "
    
    End If
    
    CadenaConsulta = CadenaConsulta & " FROM  variedades, straba, rcapataz, horas left join rcategorias on horas.codcateg = rcategorias.codcateg "
    CadenaConsulta = CadenaConsulta & " WHERE horas.codcapat = rcapataz.codcapat and  "
    CadenaConsulta = CadenaConsulta & " horas.codtraba = straba.codtraba and "
    CadenaConsulta = CadenaConsulta & " horas.codvarie = variedades.codvarie "
    
    '************************************************************************
    If vParamAplic.Cooperativa = 18 Then
        txtAux(10).Tag = "Campo|N|S|||horas|codcampo|00000000|S|"
        txtAux(10).TabIndex = 4
        
        Label9.visible = True
        txtAux2(8).visible = True
    Else
        Me.Height = 10840
    End If
    
    cadB = ""
    CargaGrid "horas.codcapat = -1"
    
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


Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    If indice = 1 Then
        txtAux(1).Text = Format(vFecha, "dd/mm/yyyy") '<===
    Else
        txtAux(5).Text = Format(vFecha, "dd/mm/yyyy") '<===
    End If
    ' ********************************************
End Sub


Private Sub frmCap_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo capataz
    txtAux2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre capataz
End Sub

Private Sub frmCat_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo categoria
    txtAux2(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre categoria
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo trabajador
    txtAux2(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre trabajador
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo variedad
    txtAux2(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre variedad
End Sub

Private Sub frmCam_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo campo
End Sub

Private Sub mnAltaRapida_Click()
    BotonAltaRapida
End Sub

Private Sub mnBorradoMasivo_Click()
    BotonBorradoMasivo
End Sub

Private Sub mnAnticipos_Click()
    BotonAnticipos
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    AbrirListadoNominas (28)
End Sub

Private Sub mnEventuales_Click()
    BotonEventuales
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

Private Sub mnTrabajadores_Click()
    BotonTrabajadores
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
        Case 8
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
        Sql = Sql & " ORDER BY  horas.fechahora, horas.codvarie "
    Else
        '********************* canviar el ORDER BY *********************++
        Sql = Sql & " ORDER BY  horas.fechahora desc, horas.codvarie, horas.codtraba, horas.codcapat "
        '**************************************************************++
    End If
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, Sql, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "S|txtAux(1)|T|Fecha|1400|;S|btnBuscar(0)|B||195|;"
    tots = tots & "S|txtAux(6)|T|Variedad|1100|;S|btnBuscar(3)|B||195|;N|txtAux2(6)|T|Variedad|1600|;"
    tots = tots & "S|txtAux(7)|T|Codigo|1000|;S|btnBuscar(4)|B||195|;S|txtAux2(7)|T|Trabajador|2700|;"
    tots = tots & "S|txtAux(0)|T|Capataz|850|;S|btnBuscar(1)|B||195|;N|txtAux2(0)|T|Capataz|1400|;"
    If vParamAplic.Cooperativa = 18 Then
        tots = tots & "S|txtAux(10)|T|Campo|1200|;S|btnBuscar(6)|B||195|;"
        tots = tots & "S|txtAux(11)|T|Categoria|1100|;S|btnBuscar(5)|B||195|;N|txtAux2(11)|T|Categoria|1400|;"
        tots = tots & "S|txtAux(4)|T|Horas|800|;S|txtAux(12)|T|Kilos|1200|;"
        tots = tots & "S|txtAux(2)|T|Importe|1200|;"
        tots = tots & "S|txtAux(3)|T|Complem.|1100|;"
        tots = tots & "S|txtAux(13)|T|Kms(€)|1100|;"
        tots = tots & "S|txtAux(9)|T|Penalización|1400|;"
        tots = tots & "S|txtAux(5)|T|F.Recibo|1400|;S|btnBuscar(2)|B||195|;N||||0|;S|chkAux(1)|CB|Cap|460|;N||||0|;S|chkAux(0)|CB|IC|360|;N|txtAux(8)|T|Almacen|800|;"
    Else
        tots = tots & "S|txtAux(11)|T|Categoria|1100|;S|btnBuscar(5)|B||195|;N|txtAux2(11)|T|Categoria|1400|;"
        tots = tots & "S|txtAux(4)|T|Horas|800|;S|txtAux(12)|T|Kilos|1200|;"
        tots = tots & "S|txtAux(2)|T|Importe|1200|;"
        tots = tots & "S|txtAux(3)|T|Complem.|1100|;"
        tots = tots & "S|txtAux(13)|T|Kms(€)|1100|;"
        tots = tots & "S|txtAux(9)|T|Penalización|1400|;"
        tots = tots & "S|txtAux(5)|T|F.Recibo|1400|;S|btnBuscar(2)|B||195|;N||||0|;S|chkAux(1)|CB|Cap|460|;N||||0|;S|chkAux(0)|CB|IC|360|;N|txtAux(8)|T|Almacen|800|;"
        tots = tots & "S|txtAux(10)|T|Nro.Parte|1100|;"
    End If
    arregla tots, DataGrid1, Me, 350
    
    CargaForaGrid
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgLeft
    DataGrid1.Columns(1).Alignment = dbgLeft
    DataGrid1.Columns(2).Alignment = dbgLeft
    DataGrid1.Columns(3).Alignment = dbgLeft
    DataGrid1.Columns(5).Alignment = dbgLeft
    DataGrid1.Columns(7).Alignment = dbgLeft
    
    CalcularTotales Sql

    
'    DataGrid1.Columns(10).Alignment = dbgCenter
'    DataGrid1.Columns(12).Alignment = dbgCenter
End Sub

Private Sub CargaForaGrid()
    If DataGrid1.Columns.Count <= 2 Then Exit Sub
    ' *** posar als camps de fora del grid el valor de la columna corresponent ***
    ' *** Si fora del grid n'hi han camps de descripció, posar-los valor ***
    ' **********************************************************************
    If adodc1.Recordset.EOF Then Exit Sub
    txtAux2(6).Text = DataGrid1.Columns(2).Text
    If vParamAplic.Cooperativa = 18 Then
        txtAux2(0).Text = DataGrid1.Columns(6).Text
        txtAux2(11).Text = DataGrid1.Columns(9).Text
        txtAux2(4).Text = GrupoTrabajo(DataGrid1.Columns(3).Text)
        
        txtAux2(8).Text = PartidaCampo(DataGrid1.Columns(7).Text)
    Else
        txtAux2(0).Text = DataGrid1.Columns(6).Text
        txtAux2(11).Text = DataGrid1.Columns(8).Text
        txtAux2(4).Text = GrupoTrabajo(DataGrid1.Columns(3).Text)
    End If
End Sub

' solo la ejecuto si es Frutas Inma
Private Function PartidaCampo(vCampo As String) As String
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error GoTo ePartidaCampo

    PartidaCampo = ""

    Sql = "select rcampos.codparti, rpartida.nomparti from rcampos inner join rpartida on rcampos.codparti = rpartida.codparti "
    Sql = Sql & " where rcampos.codcampo = " & DBSet(vCampo, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        PartidaCampo = DBLet(Rs.Fields(1).Value)
    End If
    
    Exit Function
    
ePartidaCampo:
    MuestraError Err.Number, "Poner partida campo", Err.Description
End Function


Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 ' destajo
            mnAltaRapida_Click
        Case 2 ' penalizacion
            mnEventuales_Click
        Case 3 ' bonificacion
            mnTrabajadores_Click
        Case 4
'            mnBorradoMasivo_Click
            mnAnticipos_Click
    End Select
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim PrecioHora As Currency
Dim Tipo As String
Dim Sql As String

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0 'codigo de capataz
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
            
        Case 2, 3, 9, 13 'importe, complemento, importekms
            PonerFormatoDecimal txtAux(Index), 3
    
        Case 4 ' horas
            PonerFormatoDecimal txtAux(Index), 4
            If Modo = 3 Or Modo = 4 Then
'                If ComprobarCero(txtAux(Index)) <> 0 Then
'                    If txtAux(11).Text <> "" Then
'                        PrecioHora = DevuelveValor("select precio from rcategorias where codcateg = " & DBSet(txtAux(11).Text, "N"))
'
'                        txtAux(2).Text = Round2(ComprobarCero(txtAux(Index).Text) * PrecioHora, 2)
'                        PonerFormatoDecimal txtAux(2), 1
'                    Else
'                        If txtAux(7).Text <> "" Then
'                            PrecioHora = DevuelveValor("select impsalar from salarios inner join straba on salarios.codcateg = straba.codcateg where straba.codtraba = " & DBSet(txtAux(7).Text, "N"))
'
'                            txtAux(2).Text = Round2(ComprobarCero(txtAux(Index).Text) * PrecioHora, 2)
'                            PonerFormatoDecimal txtAux(2), 1
'                        End If
'                    End If
'                End If
            End If
    
        Case 12 ' kilos
            PonerFormatoEntero txtAux(Index)
            If Modo = 3 Or Modo = 4 Then
'                If ComprobarCero(txtAux(Index)) <> 0 Then
'                    If txtAux(11).Text <> "" Then
'                        PrecioHora = DevuelveValor("select precio from rcategorias where codcateg = " & DBSet(txtAux(11).Text, "N"))
'
'                        txtAux(2).Text = Round2(ComprobarCero(txtAux(Index).Text) * PrecioHora, 2)
'                        PonerFormatoDecimal txtAux(2), 1
'                    End If
'                End If
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
    
        Case 7 ' codigo de trabajador
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "straba", "nomtraba")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe el Trabajador: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
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
                    txtAux2(4).Text = GrupoTrabajo(txtAux(7).Text)
                End If
            End If
        
        Case 11 'codigo de categoria
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "rcategorias", "nomcateg")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe la Categoria " & txtAux(Index).Text & ". Reintroduzca." & vbCrLf
                    MsgBox cadMen, vbExclamation
                    PonerFoco txtAux(Index)
                Else
                    If Modo = 3 Or Modo = 4 Then
                        Tipo = DevuelveDesdeBDNew(cAgro, "rcategorias", "tipo", "codcateg", txtAux(11), "N")
                        txtAux(4).Enabled = (ComprobarCero(Tipo) = 0)
                        txtAux(12).Enabled = (ComprobarCero(Tipo) = 1)
                    End If
'
'                    If txtAux(4).Enabled Then txtAux(12).Text = ""
'                    If txtAux(12).Enabled Then txtAux(4).Text = ""
'
'                    PrecioHora = DevuelveValor("select precio from rcategorias where codcateg = " & DBSet(txtAux(11).Text, "N"))
'
'                    txtAux(2).Text = Round2(ComprobarCero(txtAux(Index).Text) * PrecioHora, 2)
'                    PonerFormatoDecimal txtAux(2), 1
'
'
                End If
            Else
                txtAux2(Index).Text = ""
            End If
            
        Case 10 ' codigo de campo
            If PonerFormatoEntero(txtAux(Index)) Then
                If vParamAplic.Cooperativa = 18 Then
                    If txtAux(6).Text <> "" Then
                        Sql = "select codcampo from rcampos where codcampo= " & DBSet(txtAux(10).Text, "N") & " and codvarie = " & DBSet(txtAux(6).Text, "N")
                        If TotalRegistrosConsulta(Sql) = 0 Then
                            If MsgBox("El campo introducido no es de la variedad introducida. " & vbCrLf & vbCrLf & "¿ Desea continuar ? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                                PonerFoco txtAux(6)
                            End If
                        End If
                    End If
                    txtAux2(8).Text = PartidaCampo(txtAux(Index))
                End If
            End If
            
    End Select
    
    
    If Index = 4 Or Index = 7 Or Index = 11 Or Index = 12 Then
        If Modo = 3 Or Modo = 4 Then
            If txtAux(11).Text <> "" Then
                Tipo = DevuelveDesdeBDNew(cAgro, "rcategorias", "tipo", "codcateg", txtAux(11), "N")
                
                If txtAux(4).Enabled Then txtAux(12).Text = ""
                If txtAux(12).Enabled Then txtAux(4).Text = ""
                
                PrecioHora = DevuelveValor("select precio from rcategorias where codcateg = " & DBSet(txtAux(11).Text, "N"))
        
                If Tipo = 0 Then
                    txtAux(2).Text = Round2(ComprobarCero(txtAux(4).Text) * PrecioHora, 2)
                Else
                    txtAux(2).Text = Round2(ComprobarCero(txtAux(12).Text) * PrecioHora, 2)
                End If
                PonerFormatoDecimal txtAux(2), 1
            Else
                If txtAux(7).Text <> "" Then
                    PrecioHora = DevuelveValor("select impsalar from salarios inner join straba on salarios.codcateg = straba.codcateg where straba.codtraba = " & DBSet(txtAux(7).Text, "N"))
            
                    txtAux(2).Text = Round2(ComprobarCero(txtAux(4).Text) * PrecioHora, 2)
                    PonerFormatoDecimal txtAux(2), 1
                End If
            End If
        End If
    End If
    
End Sub

Private Function DatosOk() As Boolean
'Dim Datos As String
Dim b As Boolean
Dim Sql As String
Dim Mens As String

    '[Monica]02/02/2018: en el caso de que no hayan horas ponemos un cero
    If (Modo = 3 Or Modo = 4) And txtAux(4).Text = "" Then txtAux(4).Text = "0"

    b = CompForm(Me)
    If Not b Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
        If vParamAplic.Cooperativa = 18 Then
            Sql = "select count(*) from horas where codcampo = " & DBSet(txtAux(10).Text, "N")
            Sql = Sql & " and fechahora = " & DBSet(txtAux(1).Text, "F")
            Sql = Sql & " and codtraba = " & DBSet(txtAux(7).Text, "N")
            Sql = Sql & " and codvarie = " & DBSet(txtAux(6).Text, "N")
            If TotalRegistros(Sql) <> 0 Then
                MsgBox "El campo existe para esta fecha, codtraba, variedad. Reintroduzca.", vbExclamation
                PonerFoco txtAux(0)
                b = False
            End If
        Else
            Sql = "select count(*) from horas where codcapat = " & DBSet(txtAux(0).Text, "N")
            Sql = Sql & " and fechahora = " & DBSet(txtAux(1).Text, "F")
            Sql = Sql & " and codtraba = " & DBSet(txtAux(7).Text, "N")
            Sql = Sql & " and codvarie = " & DBSet(txtAux(6).Text, "N")
            If TotalRegistros(Sql) <> 0 Then
                MsgBox "El capataz existe para esta fecha, codtraba, variedad. Reintroduzca.", vbExclamation
                PonerFoco txtAux(0)
                b = False
            End If
        End If
    End If
    '[Monica]11/03/2019: para el caso de frutas inma se mete el campo
    If b And vParamAplic.Cooperativa = 18 And (Modo = 3 Or Modo = 4) Then
        If ComprobarCero(txtAux(10).Text) <> 0 And ComprobarCero(txtAux(6).Text) <> 0 Then
            Sql = "select codcampo from rcampos where codcampo= " & DBSet(txtAux(10).Text, "N") & " and codvarie = " & DBSet(txtAux(6).Text, "N")
            If TotalRegistrosConsulta(Sql) = 0 Then
                If MsgBox("El campo introducido no es de la variedad introducida. " & vbCrLf & vbCrLf & "¿ Desea continuar ? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                    PonerFoco txtAux(6)
                    b = False
                End If
            End If
        End If
    End If
    
    If b And (Modo = 3 Or Modo = 4) Then
        If Not EntreFechas(vParam.FecIniCam, txtAux(1).Text, vParam.FecFinCam) Then
            MsgBox "La fecha introducida no se encuentra dentro de campaña. Revise.", vbExclamation
            b = False
        End If
    End If
    
    DatosOk = b
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub



'Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
'    KEYpress KeyAscii
'End Sub
Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 1: KEYBusqueda KeyAscii, 0 'fecha
                Case 6: KEYBusqueda KeyAscii, 3 'variedad
                Case 7: KEYBusqueda KeyAscii, 4 'trabajador
                Case 0: KEYBusqueda KeyAscii, 1 'capataz
                Case 11: KEYBusqueda KeyAscii, 1 'categoria
                Case 5: KEYBusqueda KeyAscii, 2 'fecha de recibo
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

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    btnBuscar_Click (indice)
End Sub



Private Sub BotonAltaRapida()
    AbrirListadoNominas (24)
    CargaGrid
End Sub

Private Sub BotonEventuales()
    AbrirListadoNominas (25)
    CargaGrid
End Sub

Private Sub BotonTrabajadores()
    vCadBusqueda = ""
    AbrirListadoNominas (41)
    cadB = vCadBusqueda
    CargaGrid cadB
    PonerModoOpcionesMenu
End Sub

Private Sub BotonBorradoMasivo()
    AbrirListadoNominas (27)
    CargaGrid
End Sub


Private Sub BotonAnticipos()
    frmManHorasAntNat.Show vbModal
End Sub




Private Sub AbrirFrmVariedades(indice As Integer)
    indCodigo = 6
    
    Set frmVar = New frmBasico2
    
    AyudaVariedad frmVar
'    frmVar.DatosADevolverBusqueda = "0|1|"
'    frmVar.CodigoActual = txtAux(indCodigo)
'    frmVar.Show vbModal
    Set frmVar = Nothing
    
    PonerFoco txtAux(indCodigo)
    
End Sub

Private Sub AbrirFrmCapataz(indice As Integer)
    indCodigo = 0
    Set frmCap = New frmManCapataz
    frmCap.DatosADevolverBusqueda = "0|1|"
    frmCap.CodigoActual = txtAux(indCodigo)
    frmCap.Show vbModal
    Set frmCap = Nothing
    
    PonerFoco txtAux(indCodigo)
End Sub

Private Sub AbrirFrmTrabajador(indice As Integer)
    indCodigo = 7
    Set frmTra = New frmManTraba
    frmTra.DatosADevolverBusqueda = "0|2|"
'    frmTra.CodigoActual = txtAux(indCodigo)
    frmTra.Show vbModal
    Set frmTra = Nothing
    
    PonerFoco txtAux(indCodigo)

End Sub


Private Sub AbrirFrmCategoria(indice As Integer)
    indCodigo = 11
    Set frmCat = New frmManCategorias
    frmCat.DatosADevolverBusqueda = "0|1|"
'    frmTra.CodigoActual = txtAux(indCodigo)
    frmCat.Show vbModal
    Set frmCat = Nothing
    
    PonerFoco txtAux(indCodigo)

End Sub


Private Sub AbrirFrmCampos(indice As Integer)
    indCodigo = 10
    
    Set frmCam = New frmManCampos
    frmCam.DatosADevolverBusqueda = "0|1|"
    frmCam.Show vbModal
    Set frmCam = Nothing
    
    PonerFoco txtAux(indCodigo)

End Sub






Private Function ModificaDesdeForm() As Boolean
Dim Sql As String

    On Error GoTo eModificaDesdeForm
    
    ModificaDesdeForm = False
    
    Sql = "update horas set "
    Sql = Sql & " importe = " & DBSet(ImporteSinFormato(txtAux(2).Text), "N")
    Sql = Sql & ", compleme = " & DBSet(ImporteSinFormato(txtAux(3).Text), "N", "S")
    '[Monica]07/06/2018: importe de kms
    Sql = Sql & ", importekms = " & DBSet(ImporteSinFormato(txtAux(13).Text), "N", "S")
    
    Sql = Sql & ", horasdia = " & DBSet(ImporteSinFormato(txtAux(4).Text), "N")
    Sql = Sql & ", penaliza = " & DBSet(ImporteSinFormato(txtAux(9).Text), "N", "S")
    Sql = Sql & ", fecharec = " & DBSet(txtAux(5).Text, "F", "S")
    Sql = Sql & ", codcapat = " & DBSet(txtAux(0).Text, "N")
    Sql = Sql & ", escapataz = " & DBSet(Me.chkAux(1).Value, "N")
    Sql = Sql & " where (1=1) "
    Sql = Sql & " and fechahora = " & DBSet(txtAux(1).Text, "F")
    Sql = Sql & " and codtraba = " & DBSet(txtAux(7).Text, "N")
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
Dim ImporteKms As Currency

Dim Rs As ADODB.Recordset
Dim Sql As String

    On Error Resume Next
    
    Sql = "select sum(importe) importe , sum(coalesce(compleme,0)) compleme ,sum(coalesce(penaliza,0)) penaliza, sum(coalesce(importekms,0)) from (" & cadena & ") aaaaa"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Importe = 0
    Compleme = 0
    txtAux2(1).Text = ""
    txtAux2(2).Text = ""
    txtAux2(3).Text = ""
    txtAux2(5).Text = ""
    
    If TotalRegistrosConsulta(cadena) = 0 Then Exit Sub
    
    If Not Rs.EOF Then
        If Rs.Fields(0).Value <> 0 Then Importe = DBLet(Rs.Fields(0).Value, "N") 'Solo es para saber que hay registros que mostrar
        If Rs.Fields(1).Value <> 0 Then Compleme = DBLet(Rs.Fields(1).Value, "N") 'Solo es para saber que hay registros que mostrar
        If Rs.Fields(2).Value <> 0 Then Penaliza = DBLet(Rs.Fields(2).Value, "N") 'Solo es para saber que hay registros que mostrar
        '[Monica]07/06/2018: importekms
        If Rs.Fields(3).Value <> 0 Then ImporteKms = DBLet(Rs.Fields(3).Value, "N") 'Solo es para saber que hay registros que mostrar
    
        txtAux2(1).Text = Format(Importe, "###,###,##0.00")
        txtAux2(2).Text = Format(Compleme, "###,###,##0.00")
        txtAux2(3).Text = Format(Penaliza, "###,###,##0.00")
        txtAux2(5).Text = Format(ImporteKms, "###,###,##0.00")
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
