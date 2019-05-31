VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmContaFacSoc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Integración Contable de Facturas de Socios"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7515
   Icon            =   "frmContaFacSoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCobros 
      Height          =   7740
      Left            =   150
      TabIndex        =   15
      Top             =   180
      Width           =   7275
      Begin VB.Frame Frame1 
         Caption         =   "Datos para Selección"
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
         Height          =   2595
         Left            =   90
         TabIndex        =   16
         Top             =   225
         Width           =   7050
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
            Index           =   2
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   360
            Width           =   3465
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
            Left            =   1500
            MaxLength       =   10
            TabIndex        =   0
            Tag             =   "Seccion|N|S|||sparam|codsecci|000||"
            Top             =   360
            Width           =   825
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
            Left            =   2175
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2040
            Width           =   3690
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
            Left            =   4515
            MaxLength       =   7
            TabIndex        =   2
            Tag             =   "Nº de Factura|N|N|0|9999999|schfac|numfactu|0000000|S|"
            Top             =   945
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
            Left            =   1500
            MaxLength       =   7
            TabIndex        =   1
            Tag             =   "Nº de Factura|N|N|0|9999999|schfac|numfactu|0000000|S|"
            Top             =   945
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
            Index           =   6
            Left            =   4485
            MaxLength       =   10
            TabIndex        =   4
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   1635
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
            Index           =   5
            Left            =   1500
            MaxLength       =   10
            TabIndex        =   3
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   1635
            Width           =   1350
         End
         Begin VB.Label Label1 
            Caption         =   "Sección"
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
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   35
            Top             =   360
            Width           =   915
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   2
            Left            =   1200
            ToolTipText     =   "Buscar sección"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo de Factura"
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
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   32
            Top             =   2070
            Width           =   1695
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Factura"
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
            Left            =   255
            TabIndex        =   31
            Top             =   705
            Width           =   765
         End
         Begin VB.Label Label4 
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
            Index           =   12
            Left            =   3585
            TabIndex        =   30
            Top             =   990
            Width           =   645
         End
         Begin VB.Label Label4 
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
            Index           =   13
            Left            =   480
            TabIndex        =   29
            Top             =   945
            Width           =   735
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   6
            Left            =   4200
            Picture         =   "frmContaFacSoc.frx":000C
            ToolTipText     =   "Buscar fecha"
            Top             =   1605
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   5
            Left            =   1215
            Picture         =   "frmContaFacSoc.frx":0097
            ToolTipText     =   "Buscar fecha"
            Top             =   1620
            Width           =   240
         End
         Begin VB.Label Label4 
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
            Index           =   14
            Left            =   3570
            TabIndex        =   28
            Top             =   1650
            Width           =   645
         End
         Begin VB.Label Label4 
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
            Index           =   15
            Left            =   510
            TabIndex        =   27
            Top             =   1620
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Factura"
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
            Height          =   255
            Index           =   2
            Left            =   255
            TabIndex        =   26
            Top             =   1290
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos para la contabilización"
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
         Height          =   3285
         Left            =   90
         TabIndex        =   17
         Top             =   2850
         Width           =   7065
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
            Index           =   11
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   12
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   2730
            Width           =   1350
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
            Index           =   11
            Left            =   3780
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   2730
            Width           =   3135
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
            Index           =   10
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   11
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   2355
            Width           =   1350
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
            Index           =   10
            Left            =   3780
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   2355
            Width           =   3135
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
            Index           =   9
            Left            =   3780
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   1575
            Width           =   3135
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
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   9
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1575
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
            Index           =   0
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   6
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   390
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
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   8
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1170
            Width           =   1350
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
            Index           =   3
            Left            =   3780
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   1170
            Width           =   3135
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
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   765
            Width           =   1350
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
            Index           =   4
            Left            =   3780
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   1980
            Width           =   3135
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
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   10
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1980
            Width           =   1350
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   11
            Left            =   2160
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   2730
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Aportación"
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
            Height          =   285
            Index           =   4
            Left            =   180
            TabIndex        =   41
            Top             =   2775
            Width           =   1575
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   10
            Left            =   2160
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   2355
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Retención"
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
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   39
            Top             =   2400
            Width           =   1395
         End
         Begin VB.Label Label1 
            Caption         =   "F.Pago Negativas"
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
            Index           =   2
            Left            =   180
            TabIndex        =   37
            Top             =   1620
            Width           =   1980
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   9
            Left            =   2160
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   1575
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Recepción"
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
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   33
            Top             =   390
            Width           =   1980
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   0
            Left            =   2160
            Picture         =   "frmContaFacSoc.frx":0122
            ToolTipText     =   "Buscar fecha"
            Top             =   390
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   2160
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   1170
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "F.Pago Positivas"
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
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   25
            Top             =   1215
            Width           =   2070
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Vencimiento"
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
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   23
            Top             =   810
            Width           =   1920
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   1
            Left            =   2160
            Picture         =   "frmContaFacSoc.frx":01AD
            ToolTipText     =   "Buscar fecha"
            Top             =   765
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Banco Prevista"
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
            Height          =   195
            Index           =   24
            Left            =   180
            TabIndex        =   19
            Top             =   2025
            Width           =   1935
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   2160
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   1980
            Width           =   240
         End
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
         Left            =   6045
         TabIndex        =   14
         Top             =   7155
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
         Left            =   4860
         TabIndex        =   13
         Top             =   7155
         Width           =   1065
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   315
         Left            =   90
         TabIndex        =   20
         Top             =   6180
         Width           =   7080
         _ExtentX        =   12488
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   6555
         Width           =   7065
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   90
         TabIndex        =   21
         Top             =   6825
         Width           =   7050
      End
   End
End
Attribute VB_Name = "frmContaFacSoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MONICA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto


Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCtas As frmCtasConta 'cuentas de contabilidad
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmFPa As frmComFpa 'ForpaConta 'formas de pago de la contabilidad
Attribute frmFPa.VB_VarHelpID = -1
Private WithEvents frmSec As frmManSeccion 'secciones
Attribute frmSec.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNomRPT As String 'Nombre del informe
Private conSubRPT As Boolean 'Si el informe tiene subreports

Dim indCodigo As Integer 'indice para txtCodigo

'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String

Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

Dim vSeccion As CSeccion
Dim Tipo As Byte

Dim cContaFra As cContabilizarFacturas


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me

End Sub

Private Sub cmdAceptar_Click()
Dim Sql As String
Dim i As Byte
Dim cadWhere As String
Dim cDesde As String
Dim cHasta As String

    If Not DatosOk Then Exit Sub

    TerminaBloquear


    cadselect = "rfactsoc.contabilizado=0 "
    cadselect = cadselect & " and rfactsoc.codtipom = " & DBSet(Mid(Combo1(0).Text, 1, 3), "T")

    'D/H Fecha factura
    cDesde = Trim(txtCodigo(5).Text)
    cHasta = Trim(txtCodigo(6).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rfactsoc.fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If

    'D/H numero de factura
    cDesde = Trim(txtCodigo(7).Text)
    cHasta = Trim(txtCodigo(8).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rfactsoc.numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHColec= """) Then Exit Sub
    End If

    If Not HayRegParaInforme("rfactsoc", cadselect) Then Exit Sub

    '[Monica]13/05/2013:
    If vParamAplic.Cooperativa = 12 Then
        If ComprobarNrosRegistro(cadselect) Then
            If ComprobarFormasPago(cadselect) Then
                ContabilizarFacturas "rfactsoc", cadselect
            End If
        End If
    Else
        ContabilizarFacturas "rfactsoc", cadselect
    End If
    
     'Eliminar la tabla TMP
    BorrarTMPFacturas
    'Desbloqueamos ya no estamos contabilizando facturas
    DesBloqueoManual ("CONSOC") 'CONtabilizar facturas SOCios

eError:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "No se ha realizado el proceso de contabilización de facturas de socio. Llame a soporte."
    End If

    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    cmdCancel_Click
End Sub

Private Function ComprobarNrosRegistro(cadselect) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Existe As Boolean
Dim numFac As Long
Dim Inicio As Long

    On Error GoTo eComprobarNrosRegistro


    ComprobarNrosRegistro = False

    ' comprobamos que no existan los registros que vamos a crear
    Sql = "select numfactu, fecfactu from rfactsoc where " & cadselect
   
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Existe = False
    
    While Not Rs.EOF And Not Existe
        If Mid(Combo1(0).Text, 1, 3) = "FRS" Then
            'cuando es rectificativa el inicio será yy1, p.e.2013 --> 13100000
            Inicio = (CInt(Mid(Year(DBLet(Rs!fecfactu, "F")), 3, 2) & "1") * 100000)
        Else
            'cuando es una liquidacion normal el inicio es yy, p.e.2013 --> 13000000
            Inicio = (CInt(Mid(Year(DBLet(Rs!fecfactu, "F")), 3, 2)) * 1000000)
        End If
        
        numFac = Inicio + DBLet(Rs!numfactu, "N")
        
        
        If vParamAplic.ContabilidadNueva Then
            Sql2 = "select numregis from factpro where numregis = " & DBSet(numFac, "N") & " and anofactu = year(" & DBSet(Rs!fecfactu, "F") & ")"
        Else
            Sql2 = "select numregis from cabfactprov where numregis = " & DBSet(numFac, "N") & " and anofacpr = year(" & DBSet(Rs!fecfactu, "F") & ")"
        End If
        Set Rs2 = New ADODB.Recordset
        
        Rs2.Open Sql2, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs2.EOF Then
            Existe = True
            MsgBox "Existe el nro de Registro " & numFac & ". Revise.", vbExclamation
        End If
        
        Set Rs2 = Nothing
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
   
    ComprobarNrosRegistro = Not Existe
    Exit Function

eComprobarNrosRegistro:
    MuestraError Err.Number, "Comprobar Nros de Registro", Err.Description
End Function

Private Function ComprobarFormasPago(cadselect) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Existe As Boolean
Dim numFac As Long
Dim Inicio As Long

    On Error GoTo eComprobarFormasPago

    ComprobarFormasPago = False

    ' comprobamos que no existan los registros que vamos a crear
    Sql = "select codforpa from rfactsoc where " & cadselect
   
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Existe = True
    
    While Not Rs.EOF And Existe
        
        Sql2 = "select count(*) from forpago where codforpa = " & DBSet(Rs!Codforpa, "N")
        If TotalRegistros(Sql2) = 0 Then
            Existe = False
            MsgBox "No existe la forma de pago en la Tesoreria. Revise.", vbExclamation
        End If
        
        Set Rs2 = Nothing
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
   
    ComprobarFormasPago = Existe
    Exit Function

eComprobarFormasPago:
    MuestraError Err.Number, "Comprobar Formas de Pago", Err.Description
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(2)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection
Dim i As Integer
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
    For i = 2 To 4
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 9 To 11
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    txtCodigo(2).Text = Format(vParamAplic.Seccionhorto, "000")
    txtNombre(2).Text = PonerNombreDeCod(txtCodigo(2), "rseccion", "nomsecci", "codsecci", "N")
    
    ConexionConta
    
    ' formas de pago
    txtCodigo(3).Text = Format(vParamAplic.ForpaPosi, "000")
    txtNombre(3).Text = DevuelveDesdeBDNew(cAgro, "forpago", "nomforpa", "codforpa", txtCodigo(3).Text, "N")
    txtCodigo(9).Text = Format(vParamAplic.ForpaNega, "000")
    txtNombre(9).Text = DevuelveDesdeBDNew(cAgro, "forpago", "nomforpa", "codforpa", txtCodigo(9).Text, "N")
    ' cuentas contables
    txtCodigo(4).Text = vParamAplic.CtaBancoSoc   ' cuenta contable de banco prevista
    txtNombre(4).Text = PonerNombreCuenta(txtCodigo(4), 0)
    txtCodigo(10).Text = vParamAplic.CtaRetenSoc ' cuenta contable de retencion
    txtNombre(10).Text = PonerNombreCuenta(txtCodigo(10), 0)
    txtCodigo(11).Text = vParamAplic.CtaAportaSoc ' cuenta contable de aportacion
    txtNombre(11).Text = PonerNombreCuenta(txtCodigo(11), 0)
    
    txtCodigo(5).Text = Format(Now, "dd/mm/yyyy") ' fecha de factura desde
    txtCodigo(6).Text = Format(Now, "dd/mm/yyyy") ' fecha de factura hasta
    txtCodigo(1).Text = Format(Now, "dd/mm/yyyy") ' fecha de vencimiento
    txtCodigo(0).Text = Format(Now, "dd/mm/yyyy") ' fecha de recepcion
            
    '###Descomentar
'    CommitConexion

    FrameCobrosVisible True, H, W
    Pb1.visible = False

    CargaCombo

    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.CmdCancel.Cancel = True
'   Me.Width = w + 70
'   Me.Height = h + 350

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not vSeccion Is Nothing Then
        vSeccion.CerrarConta
        Set vSeccion = Nothing
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(1).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmTDia_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Cuentas contables
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFPa_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de formas de pago de contabilidad
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtCodigo(indCodigo).Text = Format(txtCodigo(indCodigo).Text, "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtCodigo(indCodigo).Text = Format(txtCodigo(indCodigo).Text, "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
    ConexionConta
End Sub

Private Sub imgFec_Click(Index As Integer)
'FEchas
    Dim esq, dalt As Long
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

    ' es desplega dalt i cap a la esquerra
    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + 420 + 30

    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
    imgFec(1).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index).Text <> "" Then frmC.NovaData = txtCodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(1).Tag))
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 2 ' seccion
            AbrirFrmSeccion (Index)
        Case 3, 9 ' forma de pago de la tesoreria
'            AbrirFrmForpaConta (Index)
            AbrirFrmForpa (Index)
        Case 4 'cuenta contable banco
            AbrirFrmCuentas (Index)
        Case 10, 11 ' cuentas contables de retnecion y de aportacion
            AbrirFrmCuentas (Index)
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub Optcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.CmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub OptNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.CmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'14/02/2007
'    KEYpress KeyAscii
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 5: KEYFecha KeyAscii, 2 'fecha desde factura
            Case 6: KEYFecha KeyAscii, 3 'fecha hasta factura
            Case 1: KEYFecha KeyAscii, 1 'fecha vencimiento
            Case 4: KEYBusqueda KeyAscii, 4 'cta contable banco
            Case 10: KEYBusqueda KeyAscii, 10 'cta contable retencion
            Case 11: KEYBusqueda KeyAscii, 11 'cta contable aportacion
            Case 3: KEYBusqueda KeyAscii, 3 'forma de pago positivas
            Case 9: KEYBusqueda KeyAscii, 9 'forma de pago negativas
            Case 0: KEYFecha KeyAscii, 0 'fecha de recepcion
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


    Select Case Index
        Case 2 ' SECCION
            If txtCodigo(Index).Text <> "" Then
                PonerFormatoEntero txtCodigo(Index)
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rseccion", "nomsecci", "codsecci", "N")
                ConexionConta
                
                PonerCamposDefecto txtCodigo(2).Text
            Else
                Cad = "Debe introducir obligatoriamente una sección. " & vbCrLf & vbCrLf & "     ¿ Desea continuar ?"
                If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then cmdCancel_Click
            End If

        Case 3, 9 ' FORMAS DE PAGO DE LA CONTABILIDAD(POSITIVAS Y NEGATIVAS)
            If vSeccion Is Nothing Then Exit Sub
            
            If vParamAplic.ContabilidadNueva Then
                If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtCodigo(Index).Text, "N")
            Else
                If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtCodigo(Index).Text, "N")
            End If
            If txtNombre(Index).Text = "" Then
                MsgBox "Forma de Pago  no existe en la contabilidad. Reintroduzca.", vbExclamation
            End If

        Case 4, 10, 11 ' CUENTAS CONTABLES ( banco, retencion y aportacion )
            If vSeccion Is Nothing Then Exit Sub
        
            If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = PonerNombreCuenta(txtCodigo(Index), 2)
            If txtNombre(Index).Text = "" Then
                MsgBox "Número de Cuenta contable no existe en la contabilidad. Reintroduzca.", vbExclamation
            End If

        Case 5, 6 'FECHAS
            If txtCodigo(Index).Text <> "" Then
                If PonerFormatoFecha(txtCodigo(Index)) Then
                    If Index = 5 Then
                        txtCodigo(6).Text = txtCodigo(5).Text
                    End If
                End If
            End If

        Case 0, 1 'FECHAS de vencimiento
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)


    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadselect = ""
    cadParam = ""
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
        If Not AnyadirAFormula(cadselect, devuelve) Then Exit Function
    Else
        devuelve2 = CadenaDesdeHastaBD(codD, codH, Codigo, TipCod)
        If devuelve2 = "Error" Then Exit Function
        If Not AnyadirAFormula(cadselect, devuelve2) Then Exit Function
    End If
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            cadParam = cadParam & AnyadirParametroDH(param, codD, codH, nomD, nomH)
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function

Private Sub AbrirFrmCuentas(indice As Integer)
    indCodigo = indice
    Set frmCtas = New frmCtasConta
    frmCtas.DatosADevolverBusqueda = "0|1|"
    frmCtas.CodigoActual = txtCodigo(indCodigo)
'    frmCtas.Conexion = cContaFacSoc
'    frmCtas.Facturas = False
    frmCtas.Show vbModal
    Set frmCtas = Nothing
End Sub

Private Sub AbrirFrmForpaConta(indice As Integer)
    indCodigo = indice
    Set frmFPa = New frmForpaConta
    frmFPa.DatosADevolverBusqueda = "0|1|"
    frmFPa.CodigoActual = txtCodigo(indCodigo)
'    frmFpa.Conexion = cContaFacSoc
    frmFPa.Show vbModal
    Set frmFPa = Nothing
End Sub

Private Sub AbrirFrmSeccion(indice As Integer)
    indCodigo = indice
    Set frmSec = New frmManSeccion
    frmSec.DatosADevolverBusqueda = "0|1|"
    frmSec.CodigoActual = txtCodigo(indCodigo)
'    frmSec.Facturas = False
    frmSec.Show vbModal
    Set frmSec = Nothing
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Orden1 As String
Dim Orden2 As String
Dim FFin As Date
Dim cta As String

   b = True

   If txtCodigo(6).Text = "" Then
        MsgBox "Introduzca la Fecha de Factura a contabilizar.", vbExclamation
        b = False
        PonerFoco txtCodigo(6)
   Else
        ' comprobamos que la contabilizacion se encuentre en los ejercicios contables
         Orden1 = ""
         Orden1 = DevuelveDesdeBDNew(cConta, "parametros", "fechaini", "", "", "", "", "", "", "", "", "", "")

         Orden2 = ""
         Orden2 = DevuelveDesdeBDNew(cConta, "parametros", "fechafin", "", "", "", "", "", "", "", "", "", "")
         FIni = CDate(Orden1)
         FFin = CDate(Orden2)
         
         '[Monica]25/06/2018: si no es Montifrut que es la que utiliza como fecha de recepcion la fecha de factura
         '                    antes solo estaba la comprobacion del else
         If vParamAplic.Cooperativa <> 12 Then
            If Not (CDate(Orden1) <= CDate(txtCodigo(0).Text) And CDate(txtCodigo(0).Text) < CDate(Day(FIni) & "/" & Month(FIni) & "/" & Year(FIni) + 2)) Then
               MsgBox "La Fecha de la contabilización no es del ejercicio actual ni del siguiente. Reintroduzca.", vbExclamation
               b = False
               PonerFoco txtCodigo(0)
            End If
         Else
            If Not (CDate(Orden1) <= CDate(txtCodigo(6).Text) And CDate(txtCodigo(6).Text) < CDate(Day(FIni) & "/" & Month(FIni) & "/" & Year(FIni) + 2)) Then
               MsgBox "La Fecha de la contabilización no es del ejercicio actual ni del siguiente. Reintroduzca.", vbExclamation
               b = False
               PonerFoco txtCodigo(6)
            End If
         End If
   End If

   If txtCodigo(0).Text = "" And b And vParamAplic.Cooperativa <> 12 Then
        MsgBox "Introduzca la Fecha de Recepción de Factura.", vbExclamation
        b = False
        PonerFoco txtCodigo(0)
   End If

   If txtCodigo(1).Text = "" And b And vParamAplic.Cooperativa <> 12 Then
        MsgBox "Introduzca la Fecha de Vencimiento a contabilizar.", vbExclamation
        b = False
        PonerFoco txtCodigo(1)
   End If

   If txtCodigo(3).Text = "" And b And vParamAplic.Cooperativa <> 12 Then
        MsgBox "Introduzca la Forma de Pago para contabilizar.", vbExclamation
        b = False
        PonerFoco txtCodigo(3)
   End If

   'cta contable de banco
   If b Then
        If txtCodigo(4).Text = "" Then
             MsgBox "Introduzca la Cta.Contable de Banco para contabilizar.", vbExclamation
             b = False
             PonerFoco txtCodigo(4)
        Else
             cta = ""
             cta = DevuelveDesdeBDNew(cConta, "cuentas", "codmacta", "codmacta", txtCodigo(4).Text, "T")
             If cta = "" Then
                 MsgBox "La cuenta contable de Banco no existe. Reintroduzca.", vbExclamation
                 b = False
                 PonerFoco txtCodigo(4)
             End If
        End If
    End If
   
   'cta contable de retencion
   If b Then
        If txtCodigo(10).Text = "" Then
             MsgBox "Introduzca la Cta.Contable de Retención para contabilizar.", vbExclamation
             b = False
             PonerFoco txtCodigo(10)
        Else
             cta = ""
             cta = DevuelveDesdeBDNew(cConta, "cuentas", "codmacta", "codmacta", txtCodigo(10).Text, "T")
             If cta = "" Then
                 MsgBox "La cuenta contable de Retención no existe. Reintroduzca.", vbExclamation
                 b = False
                 PonerFoco txtCodigo(10)
             End If
        End If
    End If
   
   'cta contable de aportacion
   If b Then
        If txtCodigo(11).Text = "" Then
             MsgBox "Introduzca la Cta.Contable de Aportación para contabilizar.", vbExclamation
             b = False
             PonerFoco txtCodigo(11)
        Else
             cta = ""
             cta = DevuelveDesdeBDNew(cConta, "cuentas", "codmacta", "codmacta", txtCodigo(11).Text, "T")
             If cta = "" Then
                 MsgBox "La cuenta contable de Aportación no existe. Reintroduzca.", vbExclamation
                 b = False
                 PonerFoco txtCodigo(11)
             End If
        End If
    End If
   
   'forma de pago positivas
   If b Then
        If vParamAplic.Cooperativa = 12 Then
        
        Else
            If txtCodigo(3).Text = "" Then
                 MsgBox "Introduzca la Forma de Pago para facturas positivas para contabilizar.", vbExclamation
                 b = False
                 PonerFoco txtCodigo(3)
            Else
                ' comprobamos que está en ariagro
                 cta = ""
                 cta = DevuelveDesdeBDNew(cAgro, "forpago", "nomforpa", "codforpa", txtCodigo(3).Text, "T")
                 If cta = "" Then
                     MsgBox "La Forma de Pago para facturas positivas no existe. Reintroduzca.", vbExclamation
                     b = False
                     PonerFoco txtCodigo(3)
                 End If
                 If b Then
                    ' comprobamos que esta en la conta
                    cta = ""
                    If vParamAplic.ContabilidadNueva Then
                        cta = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtCodigo(3).Text, "T")
                    Else
                        cta = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtCodigo(3).Text, "T")
                    End If
                    If cta = "" Then
                        MsgBox "La Forma de Pago para facturas positivas no existe en Tesoreria. Revise.", vbExclamation
                        b = False
                        PonerFoco txtCodigo(3)
                    End If
                 End If
            End If
        End If
    End If
   
   'forma de pago negativas
   If b Then
        If vParamAplic.Cooperativa = 12 Then
        
        Else
            If txtCodigo(9).Text = "" Then
                 MsgBox "Introduzca la Forma de Pago para facturas negativas para contabilizar.", vbExclamation
                 b = False
                 PonerFoco txtCodigo(9)
            Else
                 cta = ""
                 cta = DevuelveDesdeBDNew(cAgro, "forpago", "nomforpa", "codforpa", txtCodigo(9).Text, "T")
                 If cta = "" Then
                     MsgBox "La Forma de Pago para facturas negativas no existe. Reintroduzca.", vbExclamation
                     b = False
                     PonerFoco txtCodigo(9)
                 End If
                 If b Then
                    cta = ""
                    If vParamAplic.ContabilidadNueva Then
                        cta = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtCodigo(9).Text, "T")
                    Else
                        cta = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtCodigo(9).Text, "T")
                    End If
                    If cta = "" Then
                        MsgBox "La Forma de Pago para facturas negativas no existe en Tesoreria. Revise.", vbExclamation
                        b = False
                        PonerFoco txtCodigo(9)
                    End If
                 End If
            End If
        End If
   End If
   
   '[Monica]17/04/2019: para asegurarnos de que se integran en el ariconta que deben
   If b Then
        If vParamAplic.Cooperativa = 18 Then
            If txtCodigo(2).Text = vParamAplic.Seccionhorto Then
                If Mid(Combo1(0).Text, 1, 1) <> "F" And Combo1(0).Text <> "FRS" Then
                    If MsgBox("Seguro que quiere integrar las facturas en la seccion " & txtNombre(2).Text, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                        b = False
                    End If
                End If
            End If
        End If
   End If
   
   DatosOk = b

End Function



Private Sub ContabilizarFacturas(cadTabla As String, cadWhere As String)
'Contabiliza Facturas de Clientes o de Proveedores
Dim Sql As String
Dim b As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String



    Sql = "CONSOC" 'contabilizar facturas de socios
    
    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se pueden Contabilizar Facturas de Socios. Hay otro usuario contabilizando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


     'comprobar que se han rellenado los dos campos de fecha
     'sino rellenar con fechaini o fechafin del ejercicio
     'que guardamos en vbles Orden1,Orden2
     If txtCodigo(5).Text = "" Then
        txtCodigo(5).Text = vEmpresa.FechaIni  'fechaini del ejercicio de la conta
     End If

     If txtCodigo(6).Text = "" Then
        txtCodigo(6).Text = vEmpresa.FechaFin  'fecha fin del ejercicio de la conta
     End If


     'Comprobar que el intervalo de fechas D/H esta dentro del ejercicio de la
     'contabilidad par ello mirar en la BD de la Conta los parámetros
     
    '[Monica]25/06/2018: si no es Montifrut que es la que utiliza como fecha de recepcion la fecha de factura
    '                    antes solo estaba la comprobacion del else
    If vParamAplic.Cooperativa <> 12 Then
        If Not ComprobarFechasConta(0) Then Exit Sub
    Else
        If Not ComprobarFechasConta(6) Then Exit Sub
    End If

    'comprobar si existen  facturas anteriores al periodo solicitado
    'sin contabilizar
    If Me.txtCodigo(5).Text <> "" Then 'anteriores a fechadesde
        Sql = "SELECT COUNT(*) FROM " & cadTabla
        Sql = Sql & " WHERE fecfactu <"
        Sql = Sql & DBSet(txtCodigo(5), "F") & " AND contabilizado=0 and codtipom = " & DBSet(Mid(Combo1(0).Text, 1, 3), "T")
        If RegistrosAListar(Sql) > 0 Then
            MsgBox "Hay Facturas anteriores sin contabilizar.", vbExclamation
            Exit Sub
        End If
    End If


'    'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
'    If Not BloqueaRegistro(cadTabla, cadWhere) Then
'        MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If


    '==========================================================
    'REALIZAR COMPROBACIONES ANTES DE CONTABILIZAR FACTURAS
    '==========================================================

'    Me.lblProgess(0).Caption = "Comprobaciones: "
'    CargarProgres Me.ProgressBar1, 100

    'Cargar tabla TEMP con las Facturas que vamos a Trabajar
    b = CrearTMPFacturas(cadTabla, cadWhere)
    If Not b Then Exit Sub
    

    'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
'    TerminaBloquear
    Sql = cadTabla & " INNER JOIN tmpFactu ON " & cadTabla
    Sql = Sql & ".codtipom=tmpFactu.codtipom AND "
    
    Sql = Sql & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "
    If Not BloqueaRegistro(Sql, cadWhere) Then
        MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


    'Visualizar la barra de Progreso
    Me.Pb1.visible = True
    Me.lblProgres(0).Caption = "Comprobaciones: "
    CargarProgres Me.Pb1, 100


    'comprobar que todas las LETRAS SERIE existen en la contabilidad y en Ariagrorec
    '--------------------------------------------------------------------------
'[Monica] 29/10/2010 : no comprobamos la letra de serie en las facturas de socio
'    Me.lblProgres(1).Caption = "Comprobando letras de serie ..."
'    b = ComprobarLetraSerie(cadTABLA)
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    DoEvents
    If Not b Then Exit Sub


    'comprobar que todas las CUENTAS de los distintos socios que vamos a
    'contabilizar existen en la Conta: rsocios_seccion.codmacpro IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Contables socios en contabilidad ..."
    
    b = ComprobarCtaContable_new(cadTabla, 1, , CInt(txtCodigo(2).Text))
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    DoEvents
    If Not b Then Exit Sub


    '[Monica]08/04/2015: para el caso de catadau comprobamos las cuentas de asociados en el caso de que lo sean
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    DoEvents
    
    If vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19 Then
        Me.lblProgres(1).Caption = "Comprobando Cuentas Contables asociados en contabilidad ..."
        
        b = ComprobarCtaContable_new(cadTabla, 14, , CInt(txtCodigo(2).Text))
        If Not b Then Exit Sub
    End If


    'comprobar que todas las CUENTAS de anticipos/liquidaciones de las variedades que vamos a
    'contabilizar existen en la Conta: variedades.ctaanticipo o variedades.ctaliquidacion IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Contables variedades en contabilidad ..."
    
    '[Monica] 07/01/2010 solo se comprueba si estamos en liquidacion de industria
    If Mid(Combo1(0).Text, 1, 3) = "FLI" Then
         b = ComprobarCtaContable_new(cadTabla, 8, 12, CInt(txtCodigo(2).Text))
    End If
            
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    DoEvents
    If Not b Then Exit Sub
    
    Tipo = DevuelveValor("select tipodocu from usuarios.stipom where codtipom = " & DBSet(Mid(Combo1(0).Text, 1, 3), "T"))
    
    '[Monica] 30/03/2010 en el caso de rectificativas hablar con Manolo
    If Tipo <> 11 Then ' solo si no son rectificativas
        '[Monica]16/07/2014: añadido el caso de facturas de trasnporte de terceros de picassent
        If (Mid(Combo1(0).Text, 1, 3) = "FTS" Or Mid(Combo1(0).Text, 1, 3) = "FTT") Then
            b = ComprobarCtaContable_new(cadTabla, 8, 13, CInt(txtCodigo(2).Text))
        Else
            b = ComprobarCtaContable_new(cadTabla, 8, Tipo, CInt(txtCodigo(2).Text))
        End If
    End If
    
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    DoEvents
    If Not b Then Exit Sub

    'comprobar que todas las CUENTAS de gastos a pie de factura
    b = ComprobarCtaContable_new(cadTabla, 12, Tipo, CInt(txtCodigo(2).Text))
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    DoEvents
    If Not b Then Exit Sub

    'comprobar que todos las TIPO IVA de las distintas fecturas que vamos a
    'contabilizar existen en la Conta: rfactsoc.tipoiva IN (conta.tiposiva.codigiva)
    '--------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Tipos de IVA en contabilidad ..."
    b = ComprobarIVA(cadTabla)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    DoEvents
    If Not b Then Exit Sub


    'comprobar si hay contabilidad ANALITICA: conta.parametros.autocoste=1
    'y verificar que las cuentas de sfamia.ctaventa empiezan por el digito
    'de conta.parametros.grupogto o conta.parametros.grupovta
    'obtener el centro de coste del usuario para insertarlo en linfact
    If vEmpresa.TieneAnalitica Then  'hay contab. analitica
       Me.lblProgres(1).Caption = "Comprobando Contabilidad Analítica ..."
           
       b = ComprobarCtaContable_new(cadTabla, 7, Tipo, CInt(txtCodigo(2).Text))

       '(si tiene analítica requiere un centro de coste para insertar en conta.linfact)
       If b Then
            CCoste = ""
            b = ComprobarCCoste_new(CCoste, cadTabla)
            If Not b Then Exit Sub
       End If
       CCoste = ""
    End If
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    DoEvents


    '===========================================================================
    'CONTABILIZAR FACTURAS
    '===========================================================================
    Me.lblProgres(0).Caption = "Contabilizar Facturas Socios: "
    CargarProgres Me.Pb1, 10
    Me.lblProgres(1).Caption = "Insertando Facturas en Contabilidad..."


    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    LOG.Insertar 3, vUsu, "Contabilizar facturas Socios: " & vbCrLf & cadTabla & vbCrLf & cadWhere
    Set LOG = Nothing
    '-----------------------------------------------------------------------------


    '---- Crear tabla TEMP para los posible errores de facturas
    tmpErrores = CrearTMPErrFact(cadTabla)

    '---- Pasar las Facturas a la Contabilidad
    b = PasarFacturasAContab(cadTabla, CCoste)

    '---- Mostrar ListView de posibles errores (si hay)
    If Not b Then
        If tmpErrores Then
            'Cargar un listview con la tabla TEMP de Errores y mostrar
            'las facturas que fallaron
            frmMensajes.OpcionMensaje = 10
            frmMensajes.Show vbModal
        Else
            MsgBox "No pueden mostrarse los errores.", vbInformation
        End If
    Else
        MsgBox "El proceso ha finalizado correctamente.", vbInformation
    End If

    'Este bien o mal, si son proveedores abriremos el listado
    'Imprimimiremos un listado de contabilizacion de facturas
    '------------------------------------------------------
    If cadTabla = "rfactsoc" Or cadTabla = "rcafter" Then
        If DevuelveValor("Select count(*) from tmpinformes where codusu = " & vUsu.Codigo) > 0 Then
            InicializarVbles
            cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
            numParam = numParam + 1
            
            cadParam = "|pDHFecha=""" & vUsu.Nombre & "   Hora: " & Format(Now, "hh:mm") & """|"
            numParam = numParam + 1
            cadFormula = "({tmpinformes.codusu} =" & vUsu.Codigo & ")"
            conSubRPT = False
            If cadTabla = "rfactsoc" Then
                cadTitulo = "Listado contabilizacion FRASOC"
                cadNomRPT = "rContabSOC.rpt"
            Else
                cadTitulo = "Listado contabilizacion FRATER"
                cadNomRPT = "rContabTER.rpt"
            End If
            
            LlamarImprimir
        End If
    End If


    '---- Eliminar tabla TEMP de Errores
    BorrarTMPErrFact

End Sub

Private Function PasarFacturasAContab(cadTabla As String, CCoste As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim b As Boolean
Dim i As Integer
Dim numfactu As Integer
Dim Codigo1 As String

    On Error GoTo EPasarFac

    PasarFacturasAContab = False

    '---- Obtener el total de Facturas a Insertar en la contabilidad
    Sql = "SELECT count(*) "
    Sql = Sql & " FROM " & cadTabla & " INNER JOIN tmpFactu "
    Codigo1 = "codtipom"
    Sql = Sql & " ON " & cadTabla & "." & Codigo1 & "=tmpFactu." & Codigo1
    Sql = Sql & " AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        numfactu = Rs.Fields(0)
    Else
        numfactu = 0
    End If
    Rs.Close
    Set Rs = Nothing


    'Modificacion como David
    '-----------------------------------------------------------
    ' Mosrtaremos para cada factura de PROVEEDOR
    ' que numregis le ha asignado
    Sql = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    conn.Execute Sql
    
    Set cContaFra = New cContabilizarFacturas
    
    If Not cContaFra.EstablecerValoresInciales(ConnConta) Then
        'NO ha establcedio los valores de la conta.  Le dejaremos seguir, avisando que
        ' obviamente, no va a contabilizar las FRAS
        Sql = "Si continua, las facturas se insertaran en el registro, pero no serán contabilizadas" & vbCrLf
        Sql = Sql & "en este momento. Deberán ser contabilizadas desde el ARICONTA" & vbCrLf & vbCrLf
        Sql = Sql & Space(50) & "¿Continuar?"
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    End If
    
    '---- Pasar cada una de las facturas seleccionadas a la Conta
    If numfactu > 0 Then
        CargarProgres Me.Pb1, numfactu

        'seleccinar todas las facturas que hemos insertado en la temporal (las que vamos a contabilizar)
        Sql = "SELECT * "
        Sql = Sql & " FROM tmpFactu "

        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenStatic, adLockPessimistic, adCmdText
        i = 1

        b = True
        'pasar a contabilidad cada una de las facturas seleccionadas
        While Not Rs.EOF
            Sql = cadTabla & "." & Codigo1 & "=" & DBSet(Rs.Fields(0), "T") & " and numfactu=" & DBSet(Rs!numfactu, "T")
            Sql = Sql & " and fecfactu=" & DBSet(Rs!fecfactu, "F")
            
            If PasarFacturaSoc(Sql, CCoste, Orden2, txtCodigo(2).Text, Tipo, CDate(txtCodigo(0).Text), CDate(txtCodigo(1).Text), txtCodigo(3).Text, txtCodigo(9).Text, txtCodigo(4).Text, txtCodigo(10).Text, txtCodigo(11).Text, Mid(Combo1(0).Text, 1, 3), cContaFra, vSeccion.TipIvaREA) = False And b Then b = False
 
            '---- Laura 26VRS_/10/2006
            'Al pasar cada factura al hacer el commit desbloqueamos los registros
            'que teniamos bloqueados y los volvemos a bloquear
            'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
            Sql = cadTabla & " INNER JOIN tmpFactu ON " & cadTabla & "." & Codigo1 & "=tmpFactu." & Codigo1 & " AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "
            If Not BloqueaRegistro(Sql, cadTabla & "." & Codigo1 & "=tmpFactu." & Codigo1 & " AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu") Then
'                MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
'                Screen.MousePointer = vbDefault
'                Exit Sub
            End If
            '----

            IncrementarProgres Me.Pb1, 1
            Me.lblProgres(1).Caption = "Insertando Facturas en Contabilidad...   (" & i & " de " & numfactu & ")"
            Me.Refresh
            DoEvents
            i = i + 1
            Rs.MoveNext
        Wend

        Rs.Close
        Set Rs = Nothing
    End If
    
    Set cContaFra = Nothing



EPasarFac:
    If Err.Number <> 0 Then b = False

    If b Then
        PasarFacturasAContab = True
    Else
        PasarFacturasAContab = False
    End If
End Function

Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim i As Byte

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 0 To Combo1.Count - 1
        Combo1(i).Clear
    Next i
    
    '[Monica]18/02/2013: excluimos las facturas varias
    'tipo de factura
    Sql = "select codtipom, nomtipom from usuarios.stipom where tipodocu > 0 and tipodocu <> 12 "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 1
    While Not Rs.EOF
'        Sql = Replace(Rs.Fields(1).Value, "Factura", "Fac.")
        Sql = Rs.Fields(1).Value
        Sql = Rs.Fields(0).Value & " - " & Sql
        Combo1(0).AddItem Sql 'campo del codigo
        Combo1(0).ItemData(Combo1(0).NewIndex) = i
        i = i + 1
        Rs.MoveNext
    Wend

End Sub

Private Function ComprobarFechasConta(ind As Integer) As Boolean
'comprobar que el periodo de fechas a contabilizar esta dentro del
'periodo de fechas del ejercicio de la contabilidad
Dim FechaIni As String, FechaFin As String
Dim Cad As String
Dim Rs As ADODB.Recordset
    
Dim vSeccion As CSeccion
    
On Error GoTo EComprobar

    ComprobarFechasConta = False
    
    
    If txtCodigo(ind).Text <> "" Then
        FechaIni = "Select fechaini,fechafin From parametros"
        Set Rs = New ADODB.Recordset
        Rs.Open FechaIni, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not Rs.EOF Then
            FechaIni = DBLet(Rs!FechaIni, "F")
            FechaFin = DateAdd("yyyy", 1, DBLet(Rs!FechaFin, "F"))
            'nos guardamos los valores
            Orden1 = FechaIni
            Orden2 = FechaFin
        
            If Not EntreFechas(FechaIni, txtCodigo(ind).Text, FechaFin) Then
                 Cad = "El período de contabilización debe estar dentro del ejercicio:" & vbCrLf & vbCrLf
                 Cad = Cad & "    Desde: " & FechaIni & vbCrLf
                 Cad = Cad & "    Hasta: " & FechaFin
                 MsgBox Cad, vbExclamation
                 txtCodigo(ind).Text = ""
            Else
                ComprobarFechasConta = True
            End If
        End If
        Rs.Close
        Set Rs = Nothing
    Else
        ComprobarFechasConta = True
    End If
    
    '[Monica]20/06/2017: solo para el caso de Montifrut la fecha de recepcion es la de factura, en el resto es la de recepcion
    If ComprobarFechasConta Then
        If vParamAplic.Cooperativa <> 12 Then
            ResultadoFechaContaOK = EsFechaOKConta(CDate(txtCodigo(0)))
            If ResultadoFechaContaOK > 0 Then
                If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
                ComprobarFechasConta = False
            End If
        End If
    End If
            
    
EComprobar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Fechas", Err.Description
End Function


Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam

        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 0
        .Titulo = cadTitulo
        .NombreRPT = cadNomRPT
        .ConSubInforme = conSubRPT
        .Show vbModal
    End With
End Sub


Private Sub ConexionConta()
    
    If vSeccion Is Nothing Then
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(txtCodigo(2).Text) Then
            vSeccion.AbrirConta
        End If
    Else
        ' si el objeto existia: cerramos la conexion y volvemos crearlo
        vSeccion.CerrarConta
        Set vSeccion = Nothing
        
        
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(txtCodigo(2).Text) Then
            vSeccion.AbrirConta
        End If
    End If
End Sub

Private Sub PonerCamposDefecto(Seccion As String)
        
    Select Case CInt(Seccion)
        Case vParamAplic.Seccionhorto <> "" And CInt(ComprobarCero(vParamAplic.Seccionhorto))
            ' formas de pago
            txtCodigo(3).Text = Format(vParamAplic.ForpaPosi, "000")
            If vParamAplic.ContabilidadNueva Then
                txtNombre(3).Text = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtCodigo(3).Text, "N")
                txtCodigo(9).Text = Format(vParamAplic.ForpaNega, "000")
                txtNombre(9).Text = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtCodigo(9).Text, "N")
            Else
                txtNombre(3).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtCodigo(3).Text, "N")
                txtCodigo(9).Text = Format(vParamAplic.ForpaNega, "000")
                txtNombre(9).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtCodigo(9).Text, "N")
            End If
            ' cuentas contables
            txtCodigo(4).Text = vParamAplic.CtaBancoSoc   ' cuenta contable de banco prevista
            txtNombre(4).Text = PonerNombreCuenta(txtCodigo(4), 0)
            txtCodigo(10).Text = vParamAplic.CtaRetenSoc ' cuenta contable de retencion
            txtNombre(10).Text = PonerNombreCuenta(txtCodigo(10), 0)
            txtCodigo(11).Text = vParamAplic.CtaAportaSoc ' cuenta contable de aportacion
            txtNombre(11).Text = PonerNombreCuenta(txtCodigo(11), 0)
            
        Case vParamAplic.SeccionAlmaz <> "" And CInt(ComprobarCero(vParamAplic.SeccionAlmaz))
            ' formas de pago
            txtCodigo(3).Text = Format(vParamAplic.ForpaPosiAlmz, "000")
            If vParamAplic.ContabilidadNueva Then
                txtNombre(3).Text = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtCodigo(3).Text, "N")
                txtCodigo(9).Text = Format(vParamAplic.ForpaNegaAlmz, "000")
                txtNombre(9).Text = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtCodigo(9).Text, "N")
            Else
                txtNombre(3).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtCodigo(3).Text, "N")
                txtCodigo(9).Text = Format(vParamAplic.ForpaNegaAlmz, "000")
                txtNombre(9).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtCodigo(9).Text, "N")
            End If
                
            ' cuentas contables
            txtCodigo(4).Text = vParamAplic.CtaBancoAlmz   ' cuenta contable de banco prevista
            txtNombre(4).Text = PonerNombreCuenta(txtCodigo(4), 0)
            txtCodigo(10).Text = vParamAplic.CtaRetenAlmz ' cuenta contable de retencion
            txtNombre(10).Text = PonerNombreCuenta(txtCodigo(10), 0)
        
        Case vParamAplic.SeccionAlmaz <> "" And CInt(ComprobarCero(vParamAplic.SeccionBodega))
            ' solo podemos poner la cta de banco prevista
            txtCodigo(3).Text = ""
            txtNombre(3).Text = ""
            txtCodigo(9).Text = ""
            txtNombre(9).Text = ""
            ' cuentas contables
            txtCodigo(4).Text = vParamAplic.CtaBancoBOD   ' cuenta contable de banco prevista
            txtNombre(4).Text = PonerNombreCuenta(txtCodigo(4), 0)
            txtCodigo(10).Text = "" ' cuenta contable de retencion
            txtNombre(10).Text = ""
            
        Case Else
            ' Limpiamos forma de pago
            txtCodigo(3).Text = ""
            txtNombre(3).Text = ""
            txtCodigo(9).Text = ""
            txtNombre(9).Text = ""
            ' limpiamos cuentas contables
            txtCodigo(4).Text = ""
            txtNombre(4).Text = ""
            txtCodigo(10).Text = "" ' cuenta contable de retencion
            txtNombre(10).Text = ""
        
    End Select

End Sub


Private Sub AbrirFrmForpa(indice As Integer)
    indCodigo = indice
    Set frmFPa = New frmComFpa
    frmFPa.DeConsulta = True
    frmFPa.DatosADevolverBusqueda = "0|1|"
    frmFPa.CodigoActual = txtCodigo(indCodigo)
    frmFPa.Show vbModal
    Set frmFPa = Nothing
End Sub


