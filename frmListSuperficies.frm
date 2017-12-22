VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListSuperficies 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6390
   Icon            =   "frmListSuperficies.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameSuperficies 
      Height          =   6960
      Left            =   30
      TabIndex        =   17
      Top             =   0
      Width           =   6285
      Begin VB.Frame Frame2 
         Caption         =   "Años de Arboles"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   2025
         Left            =   330
         TabIndex        =   29
         Top             =   1980
         Width           =   5445
         Begin VB.TextBox txtcodigo 
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
            Left            =   1650
            MaxLength       =   6
            TabIndex        =   6
            Tag             =   "Arboles|N|S|0|9999|rcampos|nroarbol|#,###||"
            Top             =   1170
            Width           =   1350
         End
         Begin VB.TextBox txtcodigo 
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
            Left            =   1650
            MaxLength       =   6
            TabIndex        =   8
            Tag             =   "Arboles|N|S|0|9999|rcampos|nroarbol|#,###||"
            Top             =   1560
            Width           =   1350
         End
         Begin VB.TextBox txtcodigo 
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
            Left            =   3600
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "Arboles|N|S|0|9999|rcampos|nroarbol|#,###||"
            Top             =   1170
            Width           =   1455
         End
         Begin VB.TextBox txtcodigo 
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
            Left            =   3600
            MaxLength       =   10
            TabIndex        =   9
            Tag             =   "Arboles|N|S|0|9999|rcampos|nroarbol|#,###||"
            Top             =   1560
            Width           =   1455
         End
         Begin VB.TextBox txtcodigo 
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
            Left            =   1650
            MaxLength       =   6
            TabIndex        =   2
            Tag             =   "Arboles|N|S|0|9999|rcampos|nroarbol|#,###||"
            Top             =   360
            Width           =   1350
         End
         Begin VB.TextBox txtcodigo 
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
            Left            =   1650
            MaxLength       =   6
            TabIndex        =   4
            Tag             =   "Arboles|N|S|0|9999|rcampos|nroarbol|#,###||"
            Top             =   750
            Width           =   1350
         End
         Begin VB.TextBox txtcodigo 
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
            Left            =   3600
            MaxLength       =   10
            TabIndex        =   3
            Tag             =   "Arboles|N|S|0|9999|rcampos|nroarbol|#,###||"
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtcodigo 
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
            Left            =   3600
            MaxLength       =   10
            TabIndex        =   5
            Tag             =   "Arboles|N|S|0|9999|rcampos|nroarbol|#,###||"
            Top             =   750
            Width           =   1455
         End
         Begin VB.Label Label14 
            Caption         =   "Rango 4:"
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
            Left            =   270
            TabIndex        =   41
            Top             =   1560
            Width           =   1005
         End
         Begin VB.Label Label13 
            Caption         =   "Rango 3:"
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
            Left            =   270
            TabIndex        =   40
            Top             =   1170
            Width           =   1005
         End
         Begin VB.Label Label12 
            Caption         =   "Rango 2:"
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
            Left            =   270
            TabIndex        =   39
            Top             =   750
            Width           =   1005
         End
         Begin VB.Label Label11 
            Caption         =   "Rango 1:"
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
            Left            =   270
            TabIndex        =   38
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label Label10 
            Caption         =   "<="
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
            Left            =   3270
            TabIndex        =   37
            Top             =   1560
            Width           =   255
         End
         Begin VB.Label Label9 
            Caption         =   "<="
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
            Left            =   3270
            TabIndex        =   36
            Top             =   1170
            Width           =   255
         End
         Begin VB.Label Label8 
            Caption         =   "<="
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
            Left            =   3270
            TabIndex        =   35
            Top             =   750
            Width           =   255
         End
         Begin VB.Label Label7 
            Caption         =   "<="
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
            Left            =   3270
            TabIndex        =   34
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label6 
            Caption         =   ">="
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
            Left            =   1290
            TabIndex        =   33
            Top             =   1560
            Width           =   255
         End
         Begin VB.Label Label5 
            Caption         =   ">="
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
            Left            =   1290
            TabIndex        =   32
            Top             =   1170
            Width           =   255
         End
         Begin VB.Label Label4 
            Caption         =   ">="
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
            Left            =   1290
            TabIndex        =   31
            Top             =   750
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   ">="
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
            Left            =   1290
            TabIndex        =   30
            Top             =   360
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo Hanegadas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   1635
         Left            =   3150
         TabIndex        =   28
         Top             =   4080
         Width           =   2625
         Begin VB.OptionButton Option3 
            Caption         =   "Cultivable"
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
            Left            =   555
            TabIndex        =   43
            Top             =   1275
            Width           =   1620
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Cooperativa"
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
            Left            =   555
            TabIndex        =   12
            Top             =   330
            Width           =   1890
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Sigpac"
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
            Left            =   555
            TabIndex        =   13
            Top             =   645
            Width           =   1890
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Catastro"
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
            Left            =   555
            TabIndex        =   14
            Top             =   960
            Width           =   1620
         End
      End
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   6030
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tipo Superficie"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   1635
         Left            =   330
         TabIndex        =   26
         Top             =   4080
         Width           =   2625
         Begin VB.OptionButton Option2 
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
            Height          =   240
            Index           =   0
            Left            =   570
            TabIndex        =   10
            Top             =   525
            Width           =   1665
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Hectáreas"
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
            Left            =   570
            TabIndex        =   11
            Top             =   945
            Width           =   1755
         End
      End
      Begin VB.TextBox txtNombre 
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
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text5"
         Top             =   1470
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   0
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text5"
         Top             =   1110
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   1545
         MaxLength       =   3
         TabIndex        =   1
         Top             =   1470
         Width           =   870
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   1545
         MaxLength       =   3
         TabIndex        =   0
         Top             =   1110
         Width           =   870
      End
      Begin VB.CommandButton Command4 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListSuperficies.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command3 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListSuperficies.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
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
         Left            =   3525
         TabIndex        =   15
         Top             =   6390
         Width           =   1065
      End
      Begin VB.CommandButton cmdCancel 
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
         Left            =   4695
         TabIndex        =   16
         Top             =   6390
         Width           =   1065
      End
      Begin VB.Label Label2 
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   42
         Top             =   6360
         Width           =   3225
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1230
         MouseIcon       =   "frmListSuperficies.frx":0620
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar producto"
         Top             =   1500
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1230
         MouseIcon       =   "frmListSuperficies.frx":0772
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar producto"
         Top             =   1110
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   28
         Left            =   525
         TabIndex        =   25
         Top             =   1545
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   525
         TabIndex        =   24
         Top             =   1155
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00972E0B&
         Height          =   240
         Index           =   11
         Left            =   300
         TabIndex        =   23
         Top             =   900
         Width           =   525
      End
      Begin VB.Label Label3 
         Caption         =   "Superficies de Cultivo Edad plantaciones"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   210
         TabIndex        =   20
         Top             =   240
         Width           =   5865
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6030
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmListSuperficies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: LAURA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(Cliente As String, Observaciones As String)

Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean


Private WithEvents frmSec As frmManSeccion 'Secciones
Attribute frmSec.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'Socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmVar As frmComVar 'Variedades de comercial
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmCal As frmManCalidades 'Calidades
Attribute frmCal.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmSit As frmManSituCamp 'Situacion campos
Attribute frmSit.VB_VarHelpID = -1
Private WithEvents frmCla As frmBasico2 'Ayuda Clases de comercial
Attribute frmCla.VB_VarHelpID = -1
Private WithEvents frmPro As frmBasico2 'Ayuda Productos de comercial
Attribute frmPro.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'Mensajes
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmMens1 As frmMensajes 'Mensajes
Attribute frmMens1.VB_VarHelpID = -1


'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe
Private ConSubInforme As Boolean

Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String

Dim Indice As Integer

Dim PrimeraVez As Boolean
Dim Contabilizada As Byte

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me

End Sub



Private Sub cmdAceptar_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim cadDesde As Date
Dim cadhasta As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim nTabla As String

Dim Nregs As Long
Dim FecFac As Date

Dim B As Boolean
Dim TipoPrec As Byte

Dim vSQL As String

    InicializarVbles
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If DatosOK Then
        '======== FORMULA  ====================================
        'D/H CLASE
        cDesde = Trim(txtCodigo(0).Text)
        cHasta = Trim(txtCodigo(1).Text)
        nDesde = txtNombre(0).Text
        nHasta = txtNombre(1).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{variedades.codclase}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
        End If
        
        vSQL = ""
        If txtCodigo(0).Text <> "" Then vSQL = vSQL & " and variedades.codclase >= " & DBSet(txtCodigo(0).Text, "N")
        If txtCodigo(1).Text <> "" Then vSQL = vSQL & " and variedades.codclase <= " & DBSet(txtCodigo(1).Text, "N")
        
        'SECCION
        If Not AnyadirAFormula(cadSelect, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
        
        'CAMPOS DADOS DE ALTA
        If Not AnyadirAFormula(cadSelect, "{rcampos.fecbajas} is null ") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "isnull({rcampos.fecbajas}) ") Then Exit Sub
        
        nTabla = "(rcampos INNER JOIN variedades ON rcampos.codvarie = variedades.codvarie) "
        nTabla = nTabla & " INNER JOIN rsocios_seccion ON rcampos.codsocio = rsocios_seccion.codsocio "

        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 16
        frmMens.cadWHERE = vSQL
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        
                    
        'Comprobar si hay registros a Mostrar antes de abrir el Informe
        If HayRegParaInforme(nTabla, cadSelect) Then
            If CargarTemporalSuperficie(nTabla, cadSelect) Then
'                'tipo de hanegada
'                If Option3(0).Value Then cadParam = cadParam & "pTipoHa=0|"
'                If Option3(1).Value Then cadParam = cadParam & "pTipoHa=1|"
'                If Option3(2).Value Then cadParam = cadParam & "pTipoHa=2|"
'                numParam = numParam + 1
                
                'tipo de superficie (hectareas o hanegadas)
                If Option2(0).Value Then CadParam = CadParam & "pTipo=0|"
                If Option2(1).Value Then CadParam = CadParam & "pTipo=1|"
                numParam = numParam + 1
                
                CadParam = CadParam & "pRango1="""
                If txtCodigo(2).Text <> "" Then
                    If txtCodigo(3).Text = "" Then
                        CadParam = CadParam & ">=" & txtCodigo(2).Text
                    Else
                        CadParam = CadParam & txtCodigo(2).Text
                        If txtCodigo(3).Text <> "" Then
                            CadParam = CadParam & " y " & txtCodigo(3).Text
                        End If
                    End If
                End If
                CadParam = CadParam & """|"
                
                CadParam = CadParam & "pRango2="""
                If txtCodigo(4).Text <> "" Then
                    If txtCodigo(5).Text = "" Then
                        CadParam = CadParam & ">=" & txtCodigo(4).Text
                    Else
                        CadParam = CadParam & txtCodigo(4).Text
                        If txtCodigo(5).Text <> "" Then
                            CadParam = CadParam & " y " & txtCodigo(5).Text
                        End If
                    End If
                End If
                CadParam = CadParam & """|"
                
                CadParam = CadParam & "pRango3="""
                If txtCodigo(6).Text <> "" Then
                    If txtCodigo(7).Text = "" Then
                        CadParam = CadParam & ">=" & txtCodigo(6).Text
                    Else
                        CadParam = CadParam & txtCodigo(6).Text
                        If txtCodigo(7).Text <> "" Then
                            CadParam = CadParam & " y " & txtCodigo(7).Text
                        End If
                    End If
                End If
                CadParam = CadParam & """|"
                
                CadParam = CadParam & "pRango4="""
                If txtCodigo(8).Text <> "" Then
                    If txtCodigo(9).Text = "" Then
                        CadParam = CadParam & ">=" & txtCodigo(8).Text
                    Else
                        CadParam = CadParam & txtCodigo(8).Text
                        If txtCodigo(9).Text <> "" Then
                            CadParam = CadParam & " y " & txtCodigo(9).Text
                        End If
                    End If
                End If
                CadParam = CadParam & """|"
                
                numParam = numParam + 4
                                    
                cadTitulo = "Superficies de Cultivo y Edad de las Plantaciones"
                cadNombreRPT = "rInfSuperficies.rpt"
                
                cadFormula = ""
                If Not AnyadirAFormula(cadFormula, "{tmpsuperficies.codusu} = " & vUsu.Codigo) Then Exit Sub
                ConSubInforme = False
                
                LlamarImprimir
            End If
        End If
    End If

End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        Me.Option2(0).Value = True
        Me.Option3(0).Value = True
        
        txtCodigo(2).Text = "1"
        txtCodigo(3).Text = "2"
        txtCodigo(4).Text = "3"
        txtCodigo(5).Text = "5"
        txtCodigo(6).Text = "6"
        txtCodigo(7).Text = "12"
        txtCodigo(8).Text = "13"
        
        PonerFoco txtCodigo(0)
                    
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    PrimeraVez = True
    limpiar Me

    
    For H = 0 To 1
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H

    FrameSuperficiesVisible True, H, W
    tabla = "rcampos"
    Me.pb1.visible = False
        
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
'    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 70
End Sub

Private Sub frmCal_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de calidades
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmCla_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") ' codigo de clase
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) ' descripcion
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " {rcampos.codvarie} in (" & CadenaSeleccion & ")"
        Sql2 = " {rcampos.codvarie} in [" & CadenaSeleccion & "]"
    Else
        Sql = " {rcampos.codvarie} = -1 "
    End If
    If Not AnyadirAFormula(cadSelect, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub



Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1  'Clases
            AbrirFrmClase (Index)
        
        
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub


Private Sub Option3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Option3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'15/02/2007
'    KEYpress KeyAscii
'ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'clase desde
            Case 1: KEYBusqueda KeyAscii, 1 'clase hasta
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente
Dim B As Boolean

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1  'CLASES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "clases", "nomclase", "codclase", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
        
        Case 2, 3, 4, 5, 6, 7, 8, 9
            If txtCodigo(Index).Text <> "" Then
                PonerFormatoEntero txtCodigo(Index)
            Else
                If Index < 9 Then
                    PonerFoco txtCodigo(Index + 1)
                Else
                    cmdAceptar.SetFocus
                End If
            End If
            
        
    End Select
End Sub

Private Sub FrameSuperficiesVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameSuperficies.visible = visible
    If visible = True Then
        Me.FrameSuperficies.Top = -90
        Me.FrameSuperficies.Left = 0
        Me.FrameSuperficies.Height = 7500
        Me.FrameSuperficies.Width = 6480
        W = Me.FrameSuperficies.Width
        H = Me.FrameSuperficies.Height
    End If
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    CadParam = ""
    numParam = 0
End Sub

Private Function PonerDesdeHasta(codD As String, codH As String, nomD As String, nomH As String, param As String) As Boolean
'IN: codD,codH --> codigo Desde/Hasta
'    nomD,nomH --> Descripcion Desde/Hasta
'Añade a cadFormula y cadSelect la cadena de seleccion:
'       "(codigo>=codD AND codigo<=codH)"
' y añade a cadParam la cadena para mostrar en la cabecera informe:
'       "codigo: Desde codD-nomd Hasta: codH-nomH"
Dim devuelve As String
Dim devuelve2 As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(codD, codH, Codigo, TipCod)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    If TipCod <> "F" Then 'Fecha
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        devuelve2 = CadenaDesdeHastaBD(codD, codH, Codigo, TipCod)
        If devuelve2 = "Error" Then Exit Function
        If Not AnyadirAFormula(cadSelect, devuelve2) Then Exit Function
    End If
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            CadParam = CadParam & AnyadirParametroDH(param, codD, codH, nomD, nomH)
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function

Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = CadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .Opcion = 0
        .ConSubInforme = ConSubInforme
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmCalidad(Indice As Integer)
    indCodigo = Indice
    Set frmCal = New frmManCalidades
    frmCal.DatosADevolverBusqueda = "2|3|"
    frmCal.Show vbModal
    Set frmCal = Nothing
End Sub

Private Sub AbrirFrmSeccion(Indice As Integer)
    indCodigo = Indice
    Set frmSec = New frmManSeccion
    frmSec.DatosADevolverBusqueda = "0|1|"
    frmSec.Show vbModal
    Set frmSec = Nothing
End Sub

Private Sub AbrirFrmSocios(Indice As Integer)
    indCodigo = Indice
    Set frmSoc = New frmManSocios
    frmSoc.DatosADevolverBusqueda = "0|1|"
    frmSoc.Show vbModal
    Set frmSoc = Nothing
End Sub

Private Sub AbrirFrmSituacion(Indice As Integer)
    indCodigo = Indice
    Set frmSit = New frmManSituCamp
    frmSit.DatosADevolverBusqueda = "0|1|"
    frmSit.Show vbModal
    Set frmSit = Nothing
End Sub

Private Sub AbrirFrmSocio(Indice As Integer)
    indCodigo = Indice
    Set frmSoc = New frmManSocios
    frmSoc.DatosADevolverBusqueda = "0|1|"
    frmSoc.Show vbModal
    Set frmSoc = Nothing
End Sub

Private Sub AbrirFrmClase(Indice As Integer)
    indCodigo = Indice
    Set frmCla = New frmBasico2
    
    AyudaClasesCom frmCla, txtCodigo(Indice)
    
    Set frmCla = Nothing
End Sub

Private Sub AbrirFrmProducto(Indice As Integer)
    indCodigo = Indice
    Set frmPro = New frmBasico2
    
    AyudaProductosCom frmPro, txtCodigo(Indice).Text
    
    Set frmPro = Nothing
End Sub

Private Sub AbrirFrmVariedad(Indice As Integer)
    indCodigo = Indice
    Set frmVar = New frmComVar
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub


Private Function DatosOK() As Boolean
Dim B As Boolean
Dim Sql As String
Dim Sql2 As String
Dim vClien As cSocio
' añadido
Dim Mens As String
Dim numfactu As String
Dim numser As String
Dim Fecha As Date
Dim vCont As CTiposMov
Dim tipoMov As String
Dim I As Integer

    B = False
    
    '[Monica] 24/06/2010
    If vParamAplic.Seccionhorto = "" Then
        MsgBox "No tiene asignada la sección de Horto en parámetros. Revise.", vbExclamation
        DatosOK = False
        Exit Function
    End If
   
    ' comprobamos que haya algún rango de nro de arboles
    For I = 2 To 8
        If txtCodigo(I).Text <> "" Then
            B = True
        End If
    Next I
    
    If Not B Then
        MsgBox "Debe introducir algún rango de Árboles. Revise.", vbExclamation
        PonerFoco txtCodigo(2)
    End If
    
    DatosOK = B

End Function



Private Function CargarTemporalSuperficie(cTabla As String, cWhere As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String
Dim Sql2 As String

Dim cad As String
Dim HayReg As Boolean
Dim Nregs As Long

    On Error GoTo eCargarTemporal
    
    CargarTemporalSuperficie = False

    pb1.visible = True
    Label2(0).visible = True

    Sql2 = "delete from tmpsuperficies where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
    End If
        
    ' insertamos en la temporal con la suma de superficie a cero
    '                                       variedad
    Sql = "insert into tmpsuperficies (codusu, codvarie, superficie1, superficie2, superficie3, superficie4)    "
    Sql = Sql & "select " & DBSet(vUsu.Codigo, "N") & ",rcampos.codvarie,0,0,0,0 from " & cTabla
    Sql = Sql & " where " & cWhere
    Sql = Sql & " group by 1,2 "
    Sql = Sql & " order by 1,2 "
    
    conn.Execute Sql
    
    If Option3(0).Value Then Sql = "select rcampos.codvarie, sum(supcoope) from " & cTabla
    If Option3(1).Value Then Sql = "select rcampos.codvarie, sum(supcatas) from " & cTabla
    If Option3(2).Value Then Sql = "select rcampos.codvarie, sum(supsigpa) from " & cTabla
    If Option3(3).Value Then Sql = "select rcampos.codvarie, sum(supculti) from " & cTabla
    Sql = Sql & " where " & cWhere
    
    If txtCodigo(2).Text <> "" Or txtCodigo(3).Text <> "" Then
        ' rango 1
        Sql2 = Sql
        If txtCodigo(2).Text <> "" Then
            Sql2 = Sql2 & " and year(" & DBSet(Now, "F") & ") - anoplant >= " & DBSet(txtCodigo(2).Text, "N")
        End If
        If txtCodigo(3).Text <> "" Then
            Sql2 = Sql2 & " and year(" & DBSet(Now, "F") & ") - anoplant <= " & DBSet(txtCodigo(3).Text, "N")
        End If
        Sql2 = Sql2 & " group by 1 "
        Sql2 = Sql2 & " order by 1 "
        
        pb1.visible = True
        Label2(0).Caption = "Actualizando Rango 1"
        Nregs = TotalRegistrosConsulta(Sql2)
        If Nregs <> 0 Then
            pb1.Max = Nregs
            pb1.Value = 0
        End If
        
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql2, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        While Not Rs.EOF
            IncrementarProgresNew pb1, 1
            Me.Refresh
            DoEvents
            
            Sql1 = "update tmpsuperficies set "
            Sql1 = Sql1 & " superficie1 = superficie1 + " & DBSet(Rs.Fields(1).Value, "N")
            Sql1 = Sql1 & " where codusu = " & DBSet(vUsu.Codigo, "N") & " and codvarie = "
            Sql1 = Sql1 & DBSet(Rs.Fields(0).Value, "N")
            conn.Execute Sql1
            
            Rs.MoveNext
        
        Wend
        Set Rs = Nothing
    End If
    
    If txtCodigo(4).Text <> "" Or txtCodigo(5).Text <> "" Then
        ' rango 2
        Sql2 = Sql
        If txtCodigo(4).Text <> "" Then
            Sql2 = Sql2 & " and year(" & DBSet(Now, "F") & ") - anoplant >= " & DBSet(txtCodigo(4).Text, "N")
        End If
        If txtCodigo(5).Text <> "" Then
            Sql2 = Sql2 & " and year(" & DBSet(Now, "F") & ") - anoplant <= " & DBSet(txtCodigo(5).Text, "N")
        End If
        Sql2 = Sql2 & " group by 1 "
        Sql2 = Sql2 & " order by 1 "
        
        Label2(0).Caption = "Actualizando Rango 2"
        Nregs = TotalRegistrosConsulta(Sql2)
        If Nregs <> 0 Then
            pb1.Max = Nregs
            pb1.Value = 0
        End If
        
        
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql2, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        While Not Rs.EOF
            IncrementarProgresNew pb1, 1
            Me.Refresh
            DoEvents
            
            Sql1 = "update tmpsuperficies set "
            Sql1 = Sql1 & " superficie2 = superficie2 + " & DBSet(Rs.Fields(1).Value, "N")
            Sql1 = Sql1 & " where codusu = " & DBSet(vUsu.Codigo, "N") & " and codvarie = "
            Sql1 = Sql1 & DBSet(Rs.Fields(0).Value, "N")
            conn.Execute Sql1
            
            Rs.MoveNext
        
        Wend
        Set Rs = Nothing
    End If
    
    If txtCodigo(6).Text <> "" Or txtCodigo(7).Text <> "" Then
        ' rango 3
        Sql2 = Sql
        If txtCodigo(6).Text <> "" Then
            Sql2 = Sql2 & " and year(" & DBSet(Now, "F") & ") - anoplant >= " & DBSet(txtCodigo(6).Text, "N")
        End If
        If txtCodigo(7).Text <> "" Then
            Sql2 = Sql2 & " and year(" & DBSet(Now, "F") & ") - anoplant <= " & DBSet(txtCodigo(7).Text, "N")
        End If
        Sql2 = Sql2 & " group by 1 "
        Sql2 = Sql2 & " order by 1 "
        
        Label2(0).Caption = "Actualizando Rango 3"
        Nregs = TotalRegistrosConsulta(Sql2)
        If Nregs <> 0 Then
            pb1.Max = Nregs
            pb1.Value = 0
        End If
    
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql2, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        While Not Rs.EOF
            IncrementarProgresNew pb1, 1
            Me.Refresh
            DoEvents
            
            Sql1 = "update tmpsuperficies set "
            Sql1 = Sql1 & " superficie3 = superficie3 + " & DBSet(Rs.Fields(1).Value, "N")
            Sql1 = Sql1 & " where codusu = " & DBSet(vUsu.Codigo, "N") & " and codvarie = "
            Sql1 = Sql1 & DBSet(Rs.Fields(0).Value, "N")
            conn.Execute Sql1
            
            Rs.MoveNext
        
        Wend
        Set Rs = Nothing
    End If
    
    If txtCodigo(8).Text <> "" Or txtCodigo(9).Text <> "" Then
        ' rango 4
        Sql2 = Sql
        If txtCodigo(8).Text <> "" Then
            Sql2 = Sql2 & " and year(" & DBSet(Now, "F") & ") - anoplant >= " & DBSet(txtCodigo(8).Text, "N")
        End If
        If txtCodigo(9).Text <> "" Then
            Sql2 = Sql2 & " and year(" & DBSet(Now, "F") & ") - anoplant <= " & DBSet(txtCodigo(9).Text, "N")
        End If
        Sql2 = Sql2 & " group by 1 "
        Sql2 = Sql2 & " order by 1 "
        
        Label2(0).Caption = "Actualizando Rango 4"
        Nregs = TotalRegistrosConsulta(Sql2)
        If Nregs <> 0 Then
            pb1.Max = Nregs
            pb1.Value = 0
        End If
    
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql2, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        While Not Rs.EOF
            IncrementarProgresNew pb1, 1
            Me.Refresh
            DoEvents
            
            Sql1 = "update tmpsuperficies set "
            Sql1 = Sql1 & " superficie4 = superficie4 + " & DBSet(Rs.Fields(1).Value, "N")
            Sql1 = Sql1 & " where codusu = " & DBSet(vUsu.Codigo, "N") & " and codvarie = "
            Sql1 = Sql1 & DBSet(Rs.Fields(0).Value, "N")
            conn.Execute Sql1
            
            Rs.MoveNext
        
        Wend
        Set Rs = Nothing
    End If
    Me.Label2(0).visible = False
    Me.pb1.visible = False
    Me.Refresh
    
    
    Sql1 = "delete from tmpsuperficies where superficie1=0 and superficie2=0 and superficie3=0 and superficie4=0"
    Sql1 = Sql1 & " and codusu=" & DBSet(vUsu.Codigo, "N")
    conn.Execute Sql1
    
    
    CargarTemporalSuperficie = True
    Exit Function
    
eCargarTemporal:
    CargarTemporalSuperficie = False
    MuestraError "Cargando temporal de superficies", Err.Description
End Function



