VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAPOListados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6690
   Icon            =   "frmAPOListados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCobros 
      Height          =   5790
      Left            =   0
      TabIndex        =   21
      Top             =   30
      Width           =   6555
      Begin VB.CheckBox chkNegativas 
         Caption         =   "Sólo Negativas"
         Height          =   315
         Left            =   4560
         TabIndex        =   411
         Tag             =   "Correo|N|N|||rsocios|correo||N|"
         Top             =   3450
         Width           =   1425
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Imprimir resumen"
         Enabled         =   0   'False
         Height          =   225
         Left            =   4560
         TabIndex        =   409
         Top             =   3180
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   3240
         TabIndex        =   58
         Top             =   3600
         Width           =   2955
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   12
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   59
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   540
            Width           =   1050
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   6
            Left            =   1050
            Picture         =   "frmAPOListados.frx":000C
            ToolTipText     =   "Buscar fecha"
            Top             =   540
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Certificado"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   18
            Left            =   30
            TabIndex        =   60
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   18
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   4170
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   17
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3420
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1725
         MaxLength       =   10
         TabIndex        =   16
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2625
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1725
         MaxLength       =   10
         TabIndex        =   15
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2265
         Width           =   1050
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   4980
         TabIndex        =   20
         Top             =   4965
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3810
         TabIndex        =   19
         Top             =   4965
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1725
         MaxLength       =   16
         TabIndex        =   13
         Top             =   1215
         Width           =   1035
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1725
         MaxLength       =   16
         TabIndex        =   14
         Top             =   1590
         Width           =   1035
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text5"
         Top             =   1215
         Width           =   3285
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text5"
         Top             =   1590
         Width           =   3285
      End
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   255
         Left            =   330
         TabIndex        =   33
         Top             =   4680
         Visible         =   0   'False
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Precio Disminución Kilos"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   1
         Left            =   450
         TabIndex        =   32
         Top             =   3870
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Precio Aumento Kilos"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   0
         Left            =   450
         TabIndex        =   31
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Informe de Aportaciones"
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
         Index           =   0
         Left            =   495
         TabIndex        =   30
         Top             =   315
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   435
         TabIndex        =   29
         Top             =   1965
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   795
         TabIndex        =   28
         Top             =   2265
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   795
         TabIndex        =   27
         Top             =   2625
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1455
         Picture         =   "frmAPOListados.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   2265
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1455
         Picture         =   "frmAPOListados.frx":0122
         ToolTipText     =   "Buscar fecha"
         Top             =   2625
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   13
         Left            =   825
         TabIndex        =   26
         Top             =   1215
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   12
         Left            =   840
         TabIndex        =   25
         Top             =   1590
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   11
         Left            =   480
         TabIndex        =   24
         Top             =   975
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1440
         MouseIcon       =   "frmAPOListados.frx":01AD
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1215
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1425
         MouseIcon       =   "frmAPOListados.frx":02FF
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1590
         Width           =   240
      End
   End
   Begin VB.Frame FrameInforme 
      Height          =   6300
      Left            =   0
      TabIndex        =   61
      Top             =   30
      Width           =   6555
      Begin VB.CheckBox chkResumen 
         Caption         =   "Resumen"
         Height          =   315
         Index           =   0
         Left            =   4890
         TabIndex        =   69
         Tag             =   "Correo|N|N|||rsocios|correo||N|"
         Top             =   4380
         Width           =   1425
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Tag             =   "Tipo Relacion|N|N|0|2|rsocios|tiporelacion||N|"
         Top             =   4380
         Width           =   1590
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   19
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   83
         Text            =   "Text5"
         Top             =   3645
         Width           =   3285
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   82
         Text            =   "Text5"
         Top             =   3270
         Width           =   3285
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   1710
         MaxLength       =   16
         TabIndex        =   67
         Top             =   3645
         Width           =   1035
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   1710
         MaxLength       =   16
         TabIndex        =   66
         Top             =   3270
         Width           =   1035
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   24
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   73
         Text            =   "Text5"
         Top             =   1590
         Width           =   3285
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   23
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   72
         Text            =   "Text5"
         Top             =   1215
         Width           =   3285
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   1725
         MaxLength       =   16
         TabIndex        =   63
         Top             =   1590
         Width           =   1035
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   23
         Left            =   1725
         MaxLength       =   16
         TabIndex        =   62
         Top             =   1215
         Width           =   1035
      End
      Begin VB.CommandButton CmdAcepListado 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3810
         TabIndex        =   70
         Top             =   5535
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   4980
         TabIndex        =   71
         Top             =   5535
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   21
         Left            =   1725
         MaxLength       =   10
         TabIndex        =   64
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2265
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   22
         Left            =   1725
         MaxLength       =   10
         TabIndex        =   65
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2625
         Width           =   1050
      End
      Begin MSComctlLib.ProgressBar Pb3 
         Height          =   255
         Left            =   420
         TabIndex        =   74
         Top             =   5100
         Visible         =   0   'False
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Relación Cooperativa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   108
         Left            =   420
         TabIndex        =   410
         Top             =   4080
         Width           =   1530
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   1395
         MouseIcon       =   "frmAPOListados.frx":0451
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar aportacion"
         Top             =   3645
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   1410
         MouseIcon       =   "frmAPOListados.frx":05A3
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar aportacion"
         Top             =   3270
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Aportación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   21
         Left            =   450
         TabIndex        =   86
         Top             =   2970
         Width           =   1125
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   20
         Left            =   810
         TabIndex        =   85
         Top             =   3645
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   19
         Left            =   795
         TabIndex        =   84
         Top             =   3270
         Width           =   465
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   1425
         MouseIcon       =   "frmAPOListados.frx":06F5
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1590
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1440
         MouseIcon       =   "frmAPOListados.frx":0847
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1215
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   27
         Left            =   480
         TabIndex        =   81
         Top             =   975
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   26
         Left            =   840
         TabIndex        =   80
         Top             =   1590
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   25
         Left            =   825
         TabIndex        =   79
         Top             =   1215
         Width           =   465
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   9
         Left            =   1455
         Picture         =   "frmAPOListados.frx":0999
         ToolTipText     =   "Buscar fecha"
         Top             =   2625
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   8
         Left            =   1455
         Picture         =   "frmAPOListados.frx":0A24
         ToolTipText     =   "Buscar fecha"
         Top             =   2265
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   24
         Left            =   795
         TabIndex        =   78
         Top             =   2625
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   23
         Left            =   795
         TabIndex        =   77
         Top             =   2265
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   22
         Left            =   435
         TabIndex        =   76
         Top             =   1965
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Informe de Aportaciones"
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
         Index           =   3
         Left            =   495
         TabIndex        =   75
         Top             =   315
         Width           =   5160
      End
   End
   Begin VB.Frame FrameRegularizacion 
      Height          =   7530
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   6555
      Begin VB.Frame Frame3 
         Caption         =   "Datos para la contabilización"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1935
         Left            =   120
         TabIndex        =   48
         Top             =   4350
         Width           =   6315
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   18
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   10
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1440
            Width           =   1050
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   18
            Left            =   3075
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   1440
            Width           =   3045
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   15
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   360
            Width           =   1050
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   16
            Left            =   3075
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   720
            Width           =   3045
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   16
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   8
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   720
            Width           =   1050
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   17
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   9
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1080
            Width           =   1050
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   17
            Left            =   3075
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   1080
            Width           =   3045
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   1710
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   1440
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Banco Prevista"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   24
            Left            =   180
            TabIndex        =   55
            Top             =   1485
            Width           =   1395
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   5
            Left            =   1710
            Picture         =   "frmAPOListados.frx":0AAF
            ToolTipText     =   "Buscar fecha"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Vencimiento"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   17
            Left            =   180
            TabIndex        =   54
            Top             =   405
            Width           =   1425
         End
         Begin VB.Label Label1 
            Caption         =   "F.Pago Positivas"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   4
            Left            =   180
            TabIndex        =   53
            Top             =   765
            Width           =   1515
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   1710
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   720
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   9
            Left            =   1710
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "F.Pago Negativas"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   52
            Top             =   1125
            Width           =   1485
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos de Selección"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3495
         Left            =   120
         TabIndex        =   37
         Top             =   840
         Width           =   6315
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   14
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   6
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   3000
            Width           =   1050
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   11
            Left            =   3090
            Locked          =   -1  'True
            TabIndex        =   39
            Text            =   "Text5"
            Top             =   885
            Width           =   2955
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   10
            Left            =   3090
            Locked          =   -1  'True
            TabIndex        =   38
            Text            =   "Text5"
            Top             =   510
            Width           =   2955
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   11
            Left            =   2010
            MaxLength       =   16
            TabIndex        =   1
            Top             =   885
            Width           =   1050
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   10
            Left            =   2010
            MaxLength       =   16
            TabIndex        =   0
            Top             =   510
            Width           =   1050
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   9
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   3
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   1890
            Width           =   1050
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   8
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   2
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   1530
            Width           =   1050
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   5
            Left            =   4560
            MaxLength       =   10
            TabIndex        =   5
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   2490
            Width           =   1050
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   4
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   2490
            Width           =   1050
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Regularización"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   10
            Left            =   210
            TabIndex        =   57
            Top             =   3000
            Width           =   1545
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   4
            Left            =   1770
            Picture         =   "frmAPOListados.frx":0B3A
            ToolTipText     =   "Buscar fecha"
            Top             =   3000
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   1710
            MouseIcon       =   "frmAPOListados.frx":0BC5
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar socio"
            Top             =   885
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   2
            Left            =   1725
            MouseIcon       =   "frmAPOListados.frx":0D17
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar socio"
            Top             =   510
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Socio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   9
            Left            =   255
            TabIndex        =   47
            Top             =   300
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   8
            Left            =   1125
            TabIndex        =   46
            Top             =   885
            Width           =   420
         End
         Begin VB.Label Label4 
            Caption         =   "Desde"
            Height          =   195
            Index           =   7
            Left            =   1140
            TabIndex        =   45
            Top             =   510
            Width           =   465
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   1
            Left            =   1740
            Picture         =   "frmAPOListados.frx":0E69
            ToolTipText     =   "Buscar fecha"
            Top             =   1890
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   0
            Left            =   1740
            Picture         =   "frmAPOListados.frx":0EF4
            ToolTipText     =   "Buscar fecha"
            Top             =   1530
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   6
            Left            =   1080
            TabIndex        =   44
            Top             =   1890
            Width           =   420
         End
         Begin VB.Label Label4 
            Caption         =   "Desde"
            Height          =   195
            Index           =   5
            Left            =   1080
            TabIndex        =   43
            Top             =   1530
            Width           =   465
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Factura"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   4
            Left            =   210
            TabIndex        =   42
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Precio Aumento Kilos"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   3
            Left            =   225
            TabIndex        =   41
            Top             =   2235
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Precio Disminución Kilos"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   2
            Left            =   3195
            TabIndex        =   40
            Top             =   2235
            Width           =   1815
         End
      End
      Begin VB.CommandButton CmdAcepRegul 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   11
         Top             =   6915
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   5370
         TabIndex        =   12
         Top             =   6915
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar pb2 
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   6630
         Visible         =   0   'False
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Caption         =   "lb1"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   56
         Top             =   6300
         Visible         =   0   'False
         Width           =   6105
      End
      Begin VB.Label Label2 
         Caption         =   "Regularización de Aportaciones"
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
         Left            =   180
         TabIndex        =   36
         Top             =   300
         Width           =   5160
      End
   End
   Begin VB.Frame FrameListAporQua 
      Height          =   5850
      Left            =   30
      TabIndex        =   117
      Top             =   30
      Width           =   6555
      Begin VB.CheckBox Check2 
         Caption         =   "Salta página por socio"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3720
         TabIndex        =   186
         Top             =   4740
         Width           =   1995
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   185
         Tag             =   "Recolectado|N|N|0|1|rcampos|recolect||N|"
         Top             =   4800
         Width           =   1650
      End
      Begin VB.Frame Frame4 
         Caption         =   "Tipo"
         ForeColor       =   &H00972E0B&
         Height          =   780
         Left            =   3480
         TabIndex        =   141
         Top             =   3360
         Width           =   2460
         Begin VB.OptionButton Opcion1 
            Caption         =   "Fecha"
            Height          =   255
            Index           =   1
            Left            =   1290
            TabIndex        =   143
            Top             =   300
            Width           =   930
         End
         Begin VB.OptionButton Opcion1 
            Caption         =   "Año"
            Height          =   255
            Index           =   0
            Left            =   300
            TabIndex        =   142
            Top             =   300
            Width           =   1290
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Imprimir Resumen"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3720
         TabIndex        =   140
         Top             =   4350
         Width           =   1815
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   41
         Left            =   1770
         MaxLength       =   10
         TabIndex        =   123
         Tag             =   "Código Postal|T|S|||clientes|codposta|0000||"
         Top             =   4080
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   39
         Left            =   1785
         MaxLength       =   10
         TabIndex        =   121
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2850
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   38
         Left            =   1785
         MaxLength       =   10
         TabIndex        =   120
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2460
         Width           =   1050
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   4
         Left            =   5040
         TabIndex        =   129
         Top             =   5265
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepListAporQua 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3870
         TabIndex        =   124
         Top             =   5280
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   37
         Left            =   1785
         MaxLength       =   16
         TabIndex        =   119
         Top             =   1725
         Width           =   1035
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   36
         Left            =   1785
         MaxLength       =   16
         TabIndex        =   118
         Top             =   1320
         Width           =   1035
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   36
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   128
         Text            =   "Text5"
         Top             =   1335
         Width           =   3285
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   37
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   127
         Text            =   "Text5"
         Top             =   1710
         Width           =   3285
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   38
         Left            =   2850
         Locked          =   -1  'True
         TabIndex        =   126
         Text            =   "Text5"
         Top             =   2460
         Width           =   3285
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   39
         Left            =   2850
         Locked          =   -1  'True
         TabIndex        =   125
         Text            =   "Text5"
         Top             =   2835
         Width           =   3285
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   35
         Left            =   1770
         MaxLength       =   10
         TabIndex        =   122
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3690
         Width           =   1050
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   0
         Left            =   5760
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   4710
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Situación"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   59
         Left            =   480
         TabIndex        =   184
         Top             =   4560
         Width           =   1185
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   50
         Left            =   840
         TabIndex        =   139
         Top             =   4080
         Width           =   465
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   11
         Left            =   1470
         Picture         =   "frmAPOListados.frx":0F7F
         ToolTipText     =   "Buscar fecha"
         Top             =   4080
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   48
         Left            =   840
         TabIndex        =   138
         Top             =   3690
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Aportación"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   49
         Left            =   480
         TabIndex        =   137
         Top             =   3390
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Listado de Aportaciones"
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
         Index           =   6
         Left            =   495
         TabIndex        =   136
         Top             =   315
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "Clase"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   47
         Left            =   465
         TabIndex        =   135
         Top             =   2145
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   46
         Left            =   825
         TabIndex        =   134
         Top             =   2445
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   45
         Left            =   825
         TabIndex        =   133
         Top             =   2805
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   44
         Left            =   855
         TabIndex        =   132
         Top             =   1335
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   43
         Left            =   870
         TabIndex        =   131
         Top             =   1710
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   42
         Left            =   510
         TabIndex        =   130
         Top             =   1095
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   19
         Left            =   1470
         MouseIcon       =   "frmAPOListados.frx":100A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1725
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   16
         Left            =   1470
         MouseIcon       =   "frmAPOListados.frx":115C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   18
         Left            =   1470
         MouseIcon       =   "frmAPOListados.frx":12AE
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2850
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   17
         Left            =   1470
         MouseIcon       =   "frmAPOListados.frx":1400
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2490
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   10
         Left            =   1470
         Picture         =   "frmAPOListados.frx":1552
         ToolTipText     =   "Buscar fecha"
         Top             =   3690
         Width           =   240
      End
   End
   Begin VB.Frame FrameDevolAporQua 
      Height          =   7140
      Left            =   0
      TabIndex        =   381
      Top             =   0
      Width           =   6555
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   112
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   393
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   5280
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   111
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   391
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3960
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   110
         Left            =   1785
         MaxLength       =   3
         TabIndex        =   389
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2820
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   109
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   388
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2460
         Width           =   1050
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   13
         Left            =   4980
         TabIndex        =   397
         Top             =   6435
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepDevApoQua 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3810
         TabIndex        =   395
         Top             =   6450
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   108
         Left            =   1785
         MaxLength       =   16
         TabIndex        =   387
         Top             =   1650
         Width           =   1035
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   107
         Left            =   1785
         MaxLength       =   16
         TabIndex        =   386
         Top             =   1260
         Width           =   1035
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   107
         Left            =   2865
         Locked          =   -1  'True
         TabIndex        =   385
         Text            =   "Text5"
         Top             =   1275
         Width           =   3285
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   108
         Left            =   2865
         Locked          =   -1  'True
         TabIndex        =   384
         Text            =   "Text5"
         Top             =   1650
         Width           =   3285
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   109
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   383
         Text            =   "Text5"
         Top             =   2460
         Width           =   3285
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   110
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   382
         Text            =   "Text5"
         Top             =   2835
         Width           =   3285
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   106
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   390
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3570
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   98
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   392
         Tag             =   "Código Postal|T|S|||clientes|codposta|0000||"
         Top             =   4560
         Width           =   1050
      End
      Begin MSComctlLib.ProgressBar pb12 
         Height          =   255
         Left            =   420
         TabIndex        =   394
         Top             =   6030
         Visible         =   0   'False
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   129
         Left            =   840
         TabIndex        =   408
         Top             =   3600
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   127
         Left            =   840
         TabIndex        =   407
         Top             =   3960
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Devolución"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   120
         Left            =   510
         TabIndex        =   406
         Top             =   4980
         Width           =   1815
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   34
         Left            =   1530
         Picture         =   "frmAPOListados.frx":15DD
         ToolTipText     =   "Buscar fecha"
         Top             =   5280
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   33
         Left            =   1530
         Picture         =   "frmAPOListados.frx":1668
         ToolTipText     =   "Buscar fecha"
         Top             =   3960
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Aportación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   128
         Left            =   510
         TabIndex        =   405
         Top             =   3270
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Devolución de Aportaciones"
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
         Index           =   28
         Left            =   495
         TabIndex        =   404
         Top             =   315
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "Clase"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   126
         Left            =   540
         TabIndex        =   403
         Top             =   2145
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   125
         Left            =   855
         TabIndex        =   402
         Top             =   2445
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   124
         Left            =   855
         TabIndex        =   401
         Top             =   2805
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   123
         Left            =   885
         TabIndex        =   400
         Top             =   1275
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   122
         Left            =   900
         TabIndex        =   399
         Top             =   1650
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   121
         Left            =   540
         TabIndex        =   398
         Top             =   1035
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   57
         Left            =   1500
         MouseIcon       =   "frmAPOListados.frx":16F3
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   56
         Left            =   1500
         MouseIcon       =   "frmAPOListados.frx":1845
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1260
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   55
         Left            =   1500
         MouseIcon       =   "frmAPOListados.frx":1997
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2820
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   54
         Left            =   1500
         MouseIcon       =   "frmAPOListados.frx":1AE9
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2475
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   32
         Left            =   1530
         Picture         =   "frmAPOListados.frx":1C3B
         ToolTipText     =   "Buscar fecha"
         Top             =   3570
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Ejercicio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   119
         Left            =   510
         TabIndex        =   396
         Top             =   4560
         Width           =   1815
      End
   End
   Begin VB.Frame FrameCalculoAporQua 
      Height          =   7140
      Left            =   30
      TabIndex        =   87
      Top             =   -30
      Width           =   6555
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   32
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   115
         Text            =   "Text5"
         Top             =   1200
         Width           =   3285
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   32
         Left            =   1725
         MaxLength       =   16
         TabIndex        =   88
         Top             =   1200
         Width           =   1035
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   31
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   100
         Tag             =   "Código Postal|T|S|||clientes|codposta|0000||"
         Top             =   5400
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   20
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   97
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   4470
         Width           =   1050
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   28
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   112
         Text            =   "Text5"
         Top             =   3285
         Width           =   3285
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   27
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   111
         Text            =   "Text5"
         Top             =   2910
         Width           =   3285
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   30
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   93
         Text            =   "Text5"
         Top             =   2190
         Width           =   3285
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   29
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   91
         Text            =   "Text5"
         Top             =   1815
         Width           =   3285
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   30
         Left            =   1725
         MaxLength       =   16
         TabIndex        =   90
         Top             =   2190
         Width           =   1035
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   29
         Left            =   1725
         MaxLength       =   16
         TabIndex        =   89
         Top             =   1815
         Width           =   1035
      End
      Begin VB.CommandButton CmdAcepCalApoQua 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3810
         TabIndex        =   102
         Top             =   6450
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   4980
         TabIndex        =   104
         Top             =   6435
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   28
         Left            =   1740
         MaxLength       =   3
         TabIndex        =   94
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3270
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   27
         Left            =   1755
         MaxLength       =   3
         TabIndex        =   92
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2910
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   26
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   96
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3750
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   25
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   99
         Tag             =   "Código Postal|T|S|||clientes|codposta|0000||"
         Top             =   4980
         Width           =   1050
      End
      Begin MSComctlLib.ProgressBar Pb5 
         Height          =   255
         Left            =   420
         TabIndex        =   95
         Top             =   6030
         Visible         =   0   'False
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   15
         Left            =   1440
         MouseIcon       =   "frmAPOListados.frx":1CC6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar seccion"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Sección"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   38
         Left            =   510
         TabIndex        =   116
         Top             =   960
         Width           =   540
      End
      Begin VB.Label Label4 
         Caption         =   "Ejercicio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   37
         Left            =   450
         TabIndex        =   114
         Top             =   5400
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Año"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   28
         Left            =   450
         TabIndex        =   113
         Top             =   4980
         Width           =   1815
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   7
         Left            =   1470
         Picture         =   "frmAPOListados.frx":1E18
         ToolTipText     =   "Buscar fecha"
         Top             =   4470
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   14
         Left            =   1440
         MouseIcon       =   "frmAPOListados.frx":1EA3
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   3285
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   1440
         MouseIcon       =   "frmAPOListados.frx":1FF5
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2910
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   1440
         MouseIcon       =   "frmAPOListados.frx":2147
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   2190
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   11
         Left            =   1440
         MouseIcon       =   "frmAPOListados.frx":2299
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1815
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   36
         Left            =   480
         TabIndex        =   110
         Top             =   1575
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   35
         Left            =   840
         TabIndex        =   109
         Top             =   2190
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   34
         Left            =   825
         TabIndex        =   108
         Top             =   1815
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   33
         Left            =   795
         TabIndex        =   107
         Top             =   3255
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   32
         Left            =   795
         TabIndex        =   106
         Top             =   2895
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Clase"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   31
         Left            =   435
         TabIndex        =   105
         Top             =   2595
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Cálculo de Aportaciones"
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
         Index           =   5
         Left            =   495
         TabIndex        =   103
         Top             =   315
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "Euros/Hanegada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   345
         Index           =   30
         Left            =   450
         TabIndex        =   101
         Top             =   3750
         Width           =   1305
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Aportación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   29
         Left            =   450
         TabIndex        =   98
         Top             =   4170
         Width           =   1815
      End
   End
   Begin VB.Frame FrameDevolAporBol 
      Height          =   6870
      Left            =   0
      TabIndex        =   352
      Top             =   0
      Width           =   6555
      Begin MSComctlLib.ProgressBar Pb11 
         Height          =   255
         Left            =   360
         TabIndex        =   380
         Top             =   5820
         Visible         =   0   'False
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   96
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   377
         Text            =   "Text5"
         Top             =   4080
         Width           =   3285
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   96
         Left            =   1710
         MaxLength       =   3
         TabIndex        =   358
         Top             =   4050
         Width           =   1035
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   105
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   365
         Text            =   "Text5"
         Top             =   3270
         Width           =   3285
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   105
         Left            =   1710
         MaxLength       =   3
         TabIndex        =   357
         Top             =   3270
         Width           =   1035
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   104
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   363
         Text            =   "Text5"
         Top             =   1590
         Width           =   3285
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   103
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   361
         Text            =   "Text5"
         Top             =   1215
         Width           =   3285
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   104
         Left            =   1710
         MaxLength       =   16
         TabIndex        =   354
         Top             =   1590
         Width           =   1035
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   103
         Left            =   1710
         MaxLength       =   16
         TabIndex        =   353
         Top             =   1200
         Width           =   1035
      End
      Begin VB.CommandButton CmdAcepDevApor 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3870
         TabIndex        =   362
         Top             =   6135
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   12
         Left            =   5040
         TabIndex        =   364
         Top             =   6135
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   102
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   356
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2580
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   101
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   355
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2220
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   100
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   359
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   4800
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Height          =   285
         Index           =   99
         Left            =   1710
         MaxLength       =   30
         TabIndex        =   360
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   5370
         Width           =   4320
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   52
         Left            =   1410
         MouseIcon       =   "frmAPOListados.frx":23EB
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar aportacion"
         Top             =   4050
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Tipo Aportación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   107
         Left            =   360
         TabIndex        =   379
         Top             =   3750
         Width           =   1635
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   106
         Left            =   795
         TabIndex        =   378
         Top             =   4050
         Width           =   465
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   51
         Left            =   1410
         MouseIcon       =   "frmAPOListados.frx":253D
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar aportacion"
         Top             =   3270
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Aportación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   118
         Left            =   360
         TabIndex        =   376
         Top             =   2970
         Width           =   1125
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   117
         Left            =   795
         TabIndex        =   375
         Top             =   3270
         Width           =   465
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   50
         Left            =   1440
         MouseIcon       =   "frmAPOListados.frx":268F
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   46
         Left            =   1440
         MouseIcon       =   "frmAPOListados.frx":27E1
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1230
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   116
         Left            =   360
         TabIndex        =   374
         Top             =   975
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   115
         Left            =   840
         TabIndex        =   373
         Top             =   1590
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   114
         Left            =   825
         TabIndex        =   372
         Top             =   1215
         Width           =   465
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   31
         Left            =   1440
         Picture         =   "frmAPOListados.frx":2933
         ToolTipText     =   "Buscar fecha"
         Top             =   2550
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   30
         Left            =   1440
         Picture         =   "frmAPOListados.frx":29BE
         ToolTipText     =   "Buscar fecha"
         Top             =   2250
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   113
         Left            =   795
         TabIndex        =   371
         Top             =   2625
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   112
         Left            =   795
         TabIndex        =   370
         Top             =   2265
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   111
         Left            =   360
         TabIndex        =   369
         Top             =   1965
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Devolución de Aportaciones"
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
         Index           =   27
         Left            =   360
         TabIndex        =   368
         Top             =   315
         Width           =   5160
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   29
         Left            =   1410
         Picture         =   "frmAPOListados.frx":2A49
         ToolTipText     =   "Buscar fecha"
         Top             =   4800
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Devolución"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   110
         Left            =   360
         TabIndex        =   367
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   109
         Left            =   360
         TabIndex        =   366
         Top             =   5250
         Width           =   1815
      End
   End
   Begin VB.Frame FrameIntTesorBol 
      Height          =   7530
      Left            =   0
      TabIndex        =   288
      Top             =   0
      Width           =   6555
      Begin VB.CommandButton CmdAcepIntTesBol 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4230
         TabIndex        =   302
         Top             =   6705
         Width           =   975
      End
      Begin VB.Frame Frame16 
         Caption         =   "Datos para la contabilización"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2085
         Left            =   120
         TabIndex        =   309
         Top             =   3810
         Width           =   6315
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   83
            Left            =   1860
            MaxLength       =   10
            TabIndex        =   300
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1530
            Width           =   1050
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   83
            Left            =   2955
            Locked          =   -1  'True
            TabIndex        =   312
            Top             =   1530
            Width           =   3195
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   86
            Left            =   1860
            MaxLength       =   10
            TabIndex        =   297
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   450
            Width           =   1050
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   85
            Left            =   2955
            Locked          =   -1  'True
            TabIndex        =   311
            Top             =   810
            Width           =   3195
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   85
            Left            =   1860
            MaxLength       =   10
            TabIndex        =   298
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   810
            Width           =   1050
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   84
            Left            =   1860
            MaxLength       =   10
            TabIndex        =   299
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1170
            Width           =   1050
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   84
            Left            =   2955
            Locked          =   -1  'True
            TabIndex        =   310
            Top             =   1170
            Width           =   3195
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   41
            Left            =   1590
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   1530
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Banco Prevista"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   23
            Left            =   180
            TabIndex        =   316
            Top             =   1575
            Width           =   1395
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   26
            Left            =   1590
            Picture         =   "frmAPOListados.frx":2AD4
            ToolTipText     =   "Buscar fecha"
            Top             =   450
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Vencimiento"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   95
            Left            =   180
            TabIndex        =   315
            Top             =   495
            Width           =   1395
         End
         Begin VB.Label Label1 
            Caption         =   "F.Pago Positivas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   22
            Left            =   180
            TabIndex        =   314
            Top             =   855
            Width           =   1275
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   45
            Left            =   1590
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   810
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   44
            Left            =   1590
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   1170
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "F.Pago Negativas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   21
            Left            =   180
            TabIndex        =   313
            Top             =   1215
            Width           =   1365
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Datos de Selección"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2865
         Left            =   120
         TabIndex        =   289
         Top             =   780
         Width           =   6315
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   75
            Left            =   2970
            Locked          =   -1  'True
            TabIndex        =   320
            Text            =   "Text5"
            Top             =   2310
            Width           =   3285
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   75
            Left            =   1860
            MaxLength       =   3
            TabIndex        =   296
            Top             =   2310
            Width           =   1035
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   82
            Left            =   2940
            Locked          =   -1  'True
            TabIndex        =   294
            Text            =   "Text5"
            Top             =   975
            Width           =   3165
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   81
            Left            =   2940
            Locked          =   -1  'True
            TabIndex        =   292
            Text            =   "Text5"
            Top             =   600
            Width           =   3165
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   82
            Left            =   1860
            MaxLength       =   16
            TabIndex        =   291
            Top             =   960
            Width           =   1050
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   81
            Left            =   1860
            MaxLength       =   16
            TabIndex        =   290
            Top             =   600
            Width           =   1050
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   80
            Left            =   1860
            MaxLength       =   10
            TabIndex        =   295
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   1860
            Width           =   1050
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   79
            Left            =   1860
            MaxLength       =   10
            TabIndex        =   293
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   1530
            Width           =   1050
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Aportación"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   82
            Left            =   210
            TabIndex        =   321
            Top             =   2130
            Width           =   1125
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   40
            Left            =   1560
            MouseIcon       =   "frmAPOListados.frx":2B5F
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar aportación"
            Top             =   2340
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   43
            Left            =   1560
            MouseIcon       =   "frmAPOListados.frx":2CB1
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar socio"
            Top             =   960
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   42
            Left            =   1560
            MouseIcon       =   "frmAPOListados.frx":2E03
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar socio"
            Top             =   600
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Socio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   93
            Left            =   225
            TabIndex        =   308
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   92
            Left            =   915
            TabIndex        =   307
            Top             =   975
            Width           =   420
         End
         Begin VB.Label Label4 
            Caption         =   "Desde"
            Height          =   195
            Index           =   91
            Left            =   930
            TabIndex        =   306
            Top             =   600
            Width           =   465
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   24
            Left            =   1560
            Picture         =   "frmAPOListados.frx":2F55
            ToolTipText     =   "Buscar fecha"
            Top             =   1860
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   23
            Left            =   1560
            Picture         =   "frmAPOListados.frx":2FE0
            ToolTipText     =   "Buscar fecha"
            Top             =   1530
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   90
            Left            =   900
            TabIndex        =   305
            Top             =   1890
            Width           =   420
         End
         Begin VB.Label Label4 
            Caption         =   "Desde"
            Height          =   195
            Index           =   89
            Left            =   900
            TabIndex        =   303
            Top             =   1560
            Width           =   465
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Aportación"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   85
            Left            =   210
            TabIndex        =   301
            Top             =   1260
            Width           =   1815
         End
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   10
         Left            =   5370
         TabIndex        =   304
         Top             =   6705
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb10 
         Height          =   255
         Left            =   210
         TabIndex        =   317
         Top             =   6270
         Visible         =   0   'False
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Caption         =   "lb1"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   25
         Left            =   270
         TabIndex        =   319
         Top             =   5940
         Visible         =   0   'False
         Width           =   6105
      End
      Begin VB.Label Label7 
         Caption         =   "Integración Aportaciones Tesoreria"
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
         Left            =   180
         TabIndex        =   318
         Top             =   270
         Width           =   5160
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameRegBajaSocios 
      Height          =   5400
      Left            =   0
      TabIndex        =   208
      Top             =   0
      Width           =   6555
      Begin VB.Frame Frame11 
         Caption         =   "Datos para la contabilización"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1935
         Left            =   120
         TabIndex        =   210
         Top             =   2130
         Width           =   6315
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   58
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   224
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1440
            Width           =   1050
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   58
            Left            =   3075
            Locked          =   -1  'True
            TabIndex        =   213
            Top             =   1440
            Width           =   3045
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   57
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   221
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   360
            Width           =   1050
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   56
            Left            =   3075
            Locked          =   -1  'True
            TabIndex        =   212
            Top             =   720
            Width           =   3045
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   56
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   222
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   720
            Width           =   1050
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   55
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   223
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1080
            Width           =   1050
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   55
            Left            =   3075
            Locked          =   -1  'True
            TabIndex        =   211
            Top             =   1080
            Width           =   3045
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   32
            Left            =   1710
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   1440
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Banco Prevista"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   17
            Left            =   180
            TabIndex        =   217
            Top             =   1485
            Width           =   1395
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   17
            Left            =   1710
            Picture         =   "frmAPOListados.frx":306B
            ToolTipText     =   "Buscar fecha"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Vencimiento"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   62
            Left            =   180
            TabIndex        =   216
            Top             =   405
            Width           =   1425
         End
         Begin VB.Label Label1 
            Caption         =   "F.Pago Positivas"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   16
            Left            =   180
            TabIndex        =   215
            Top             =   765
            Width           =   1515
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   31
            Left            =   1710
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   720
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   30
            Left            =   1710
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "F.Pago Negativas"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   15
            Left            =   180
            TabIndex        =   214
            Top             =   1125
            Width           =   1485
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Datos para la selección"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1245
         Left            =   120
         TabIndex        =   209
         Top             =   780
         Width           =   6315
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   59
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   219
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   360
            Width           =   1050
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   59
            Left            =   3075
            Locked          =   -1  'True
            TabIndex        =   229
            Top             =   360
            Width           =   3045
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   54
            Left            =   2010
            MaxLength       =   10
            TabIndex        =   220
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   810
            Width           =   1050
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   33
            Left            =   1710
            ToolTipText     =   "Buscar socio"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Socio"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   18
            Left            =   180
            TabIndex        =   230
            Top             =   405
            Width           =   1515
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   18
            Left            =   1710
            Picture         =   "frmAPOListados.frx":30F6
            ToolTipText     =   "Buscar fecha"
            Top             =   780
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Devolución"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   61
            Left            =   180
            TabIndex        =   227
            Top             =   765
            Width           =   1425
         End
      End
      Begin VB.CommandButton CmdAcepRegBajaSocios 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   225
         Top             =   4755
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   7
         Left            =   5370
         TabIndex        =   226
         Top             =   4755
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb7 
         Height          =   255
         Left            =   210
         TabIndex        =   228
         Top             =   4320
         Visible         =   0   'False
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label6 
         Caption         =   "Devolución por Baja Socios"
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
         Left            =   180
         TabIndex        =   218
         Top             =   270
         Width           =   5160
      End
   End
   Begin VB.Frame FrameCertificadoBol 
      Height          =   7530
      Left            =   0
      TabIndex        =   322
      Top             =   0
      Width           =   6555
      Begin VB.TextBox txtcodigo 
         Height          =   975
         Index           =   95
         Left            =   1710
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   335
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   5640
         Width           =   4320
      End
      Begin VB.TextBox txtcodigo 
         Height          =   285
         Index           =   94
         Left            =   1710
         MaxLength       =   30
         TabIndex        =   334
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   5220
         Width           =   4320
      End
      Begin VB.TextBox txtcodigo 
         Height          =   285
         Index           =   93
         Left            =   1710
         MaxLength       =   30
         TabIndex        =   333
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   4830
         Width           =   4320
      End
      Begin VB.TextBox txtcodigo 
         Height          =   285
         Index           =   92
         Left            =   1710
         MaxLength       =   30
         TabIndex        =   332
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   4440
         Width           =   4320
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   76
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   331
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3930
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   91
         Left            =   1725
         MaxLength       =   10
         TabIndex        =   329
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2625
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   90
         Left            =   1725
         MaxLength       =   10
         TabIndex        =   328
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2265
         Width           =   1050
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   11
         Left            =   4980
         TabIndex        =   337
         Top             =   6855
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepCertBol 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3810
         TabIndex        =   336
         Top             =   6855
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   89
         Left            =   1710
         MaxLength       =   16
         TabIndex        =   327
         Top             =   1590
         Width           =   1035
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   88
         Left            =   1710
         MaxLength       =   16
         TabIndex        =   326
         Top             =   1200
         Width           =   1035
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   88
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   325
         Text            =   "Text5"
         Top             =   1215
         Width           =   3285
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   89
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   324
         Text            =   "Text5"
         Top             =   1590
         Width           =   3285
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   87
         Left            =   1710
         MaxLength       =   3
         TabIndex        =   330
         Top             =   3270
         Width           =   1035
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   87
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   323
         Text            =   "Text5"
         Top             =   3270
         Width           =   3285
      End
      Begin VB.Label Label4 
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   105
         Left            =   480
         TabIndex        =   351
         Top             =   5640
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Tesorero"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   104
         Left            =   480
         TabIndex        =   350
         Top             =   5220
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Secretario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   103
         Left            =   480
         TabIndex        =   349
         Top             =   4830
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Presidente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   102
         Left            =   480
         TabIndex        =   348
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Certificado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   84
         Left            =   450
         TabIndex        =   347
         Top             =   3630
         Width           =   1815
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   28
         Left            =   1470
         Picture         =   "frmAPOListados.frx":3181
         ToolTipText     =   "Buscar fecha"
         Top             =   3930
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Certificado de Aportaciones"
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
         Index           =   26
         Left            =   495
         TabIndex        =   346
         Top             =   315
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   101
         Left            =   435
         TabIndex        =   345
         Top             =   1965
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   100
         Left            =   795
         TabIndex        =   344
         Top             =   2265
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   99
         Left            =   795
         TabIndex        =   343
         Top             =   2625
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   27
         Left            =   1440
         Picture         =   "frmAPOListados.frx":320C
         ToolTipText     =   "Buscar fecha"
         Top             =   2610
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   25
         Left            =   1440
         Picture         =   "frmAPOListados.frx":3297
         ToolTipText     =   "Buscar fecha"
         Top             =   2250
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   98
         Left            =   825
         TabIndex        =   342
         Top             =   1215
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   97
         Left            =   840
         TabIndex        =   341
         Top             =   1590
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   96
         Left            =   480
         TabIndex        =   340
         Top             =   975
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   49
         Left            =   1440
         MouseIcon       =   "frmAPOListados.frx":3322
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1590
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   48
         Left            =   1440
         MouseIcon       =   "frmAPOListados.frx":3474
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1230
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   94
         Left            =   795
         TabIndex        =   339
         Top             =   3270
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Aportación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   83
         Left            =   450
         TabIndex        =   338
         Top             =   2970
         Width           =   1125
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   47
         Left            =   1410
         MouseIcon       =   "frmAPOListados.frx":35C6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar aportacion"
         Top             =   3270
         Width           =   240
      End
   End
   Begin VB.Frame FrameRegAltaSocios 
      Height          =   5400
      Left            =   0
      TabIndex        =   187
      Top             =   0
      Width           =   6555
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   6
         Left            =   5370
         TabIndex        =   203
         Top             =   4755
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepRegAltaSocios 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   202
         Top             =   4755
         Width           =   975
      End
      Begin VB.Frame Frame9 
         Caption         =   "Datos de Selección"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1005
         Left            =   120
         TabIndex        =   196
         Top             =   840
         Width           =   6315
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   60
            Left            =   2010
            MaxLength       =   10
            TabIndex        =   197
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   450
            Width           =   1050
         End
         Begin VB.Label Label4 
            Caption         =   "Precio Kilo"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   68
            Left            =   195
            TabIndex        =   204
            Top             =   465
            Width           =   1815
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Datos para la contabilización"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1935
         Left            =   120
         TabIndex        =   188
         Top             =   1890
         Width           =   6315
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   53
            Left            =   3075
            Locked          =   -1  'True
            TabIndex        =   191
            Top             =   1080
            Width           =   3045
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   53
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   200
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1080
            Width           =   1050
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   52
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   199
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   720
            Width           =   1050
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   52
            Left            =   3075
            Locked          =   -1  'True
            TabIndex        =   190
            Top             =   720
            Width           =   3045
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   51
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   198
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   360
            Width           =   1050
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   50
            Left            =   3075
            Locked          =   -1  'True
            TabIndex        =   189
            Top             =   1440
            Width           =   3045
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   50
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   201
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1440
            Width           =   1050
         End
         Begin VB.Label Label1 
            Caption         =   "F.Pago Negativas"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   13
            Left            =   180
            TabIndex        =   195
            Top             =   1125
            Width           =   1485
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   29
            Left            =   1710
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   1080
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   28
            Left            =   1710
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   720
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "F.Pago Positivas"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   12
            Left            =   180
            TabIndex        =   194
            Top             =   765
            Width           =   1515
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Vencimiento"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   60
            Left            =   180
            TabIndex        =   193
            Top             =   405
            Width           =   1425
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   16
            Left            =   1710
            Picture         =   "frmAPOListados.frx":3718
            ToolTipText     =   "Buscar fecha"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Banco Prevista"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   11
            Left            =   180
            TabIndex        =   192
            Top             =   1485
            Width           =   1395
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   27
            Left            =   1710
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   1440
            Width           =   240
         End
      End
      Begin MSComctlLib.ProgressBar Pb6 
         Height          =   255
         Left            =   210
         TabIndex        =   205
         Top             =   4320
         Visible         =   0   'False
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label5 
         Caption         =   "Regularización por Alta Socios"
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
         Left            =   180
         TabIndex        =   207
         Top             =   270
         Width           =   5160
      End
      Begin VB.Label Label1 
         Caption         =   "lb1"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   14
         Left            =   270
         TabIndex        =   206
         Top             =   3990
         Visible         =   0   'False
         Width           =   6105
      End
   End
   Begin VB.Frame FrameInsertarApoBol 
      Height          =   7470
      Left            =   0
      TabIndex        =   231
      Top             =   60
      Width           =   6555
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   2505
         Left            =   150
         TabIndex        =   256
         Top             =   4080
         Width           =   6165
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   68
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   264
            Text            =   "Text5"
            Top             =   300
            Width           =   3285
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   68
            Left            =   1560
            MaxLength       =   3
            TabIndex        =   263
            Top             =   285
            Width           =   1035
         End
         Begin VB.TextBox txtcodigo 
            Height          =   285
            Index           =   63
            Left            =   1560
            MaxLength       =   30
            TabIndex        =   249
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   840
            Width           =   4350
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   69
            Left            =   1560
            MaxLength       =   30
            TabIndex        =   250
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   1470
            Width           =   1020
         End
         Begin MSComctlLib.ProgressBar Pb8 
            Height          =   255
            Left            =   210
            TabIndex        =   257
            Top             =   1890
            Visible         =   0   'False
            Width           =   6045
            _ExtentX        =   10663
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Aportación"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   73
            Left            =   240
            TabIndex        =   265
            Top             =   0
            Width           =   1125
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   36
            Left            =   1230
            MouseIcon       =   "frmAPOListados.frx":37A3
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar aportación"
            Top             =   300
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Descripción"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   65
            Left            =   270
            TabIndex        =   259
            Top             =   630
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Porcentaje de Aportación"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   76
            Left            =   270
            TabIndex        =   258
            Top             =   1200
            Width           =   1875
         End
      End
      Begin VB.Frame Frame12 
         BorderStyle     =   0  'None
         Height          =   765
         Left            =   390
         TabIndex        =   260
         Top             =   4080
         Width           =   3135
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   70
            Left            =   1290
            MaxLength       =   10
            TabIndex        =   261
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   315
            Width           =   1050
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Recibo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   77
            Left            =   0
            TabIndex        =   262
            Top             =   60
            Width           =   1815
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   19
            Left            =   975
            Picture         =   "frmAPOListados.frx":38F5
            ToolTipText     =   "Buscar fecha"
            Top             =   300
            Width           =   240
         End
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   62
         Left            =   1680
         MaxLength       =   7
         TabIndex        =   244
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   1500
         Width           =   830
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   61
         Left            =   1680
         MaxLength       =   7
         TabIndex        =   243
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   1110
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   67
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   233
         Text            =   "Text5"
         Top             =   3510
         Width           =   3285
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   66
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   232
         Text            =   "Text5"
         Top             =   3135
         Width           =   3285
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   67
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   248
         Top             =   3510
         Width           =   1035
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   66
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   247
         Top             =   3135
         Width           =   1035
      End
      Begin VB.CommandButton CmdAcepInsApoBol 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   251
         Top             =   6660
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   8
         Left            =   4980
         TabIndex        =   252
         Top             =   6645
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   65
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   246
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2640
         Width           =   1050
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   64
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   245
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2280
         Width           =   1050
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1110
         Index           =   0
         Left            =   2940
         TabIndex        =   241
         Top             =   1110
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1958
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   75
         Left            =   810
         TabIndex        =   255
         Top             =   1500
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   64
         Left            =   810
         TabIndex        =   254
         Top             =   1140
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nro.Factura"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   63
         Left            =   390
         TabIndex        =   253
         Top             =   870
         Width           =   870
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de Factura"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   74
         Left            =   2970
         TabIndex        =   242
         Top             =   870
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   3
         Left            =   5820
         Picture         =   "frmAPOListados.frx":3980
         ToolTipText     =   "Desmarcar todos"
         Top             =   840
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   2
         Left            =   5580
         Picture         =   "frmAPOListados.frx":4382
         ToolTipText     =   "Marcar todos"
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   35
         Left            =   1365
         MouseIcon       =   "frmAPOListados.frx":ABD4
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   3510
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   34
         Left            =   1365
         MouseIcon       =   "frmAPOListados.frx":AD26
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   3135
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   72
         Left            =   390
         TabIndex        =   240
         Top             =   2895
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   71
         Left            =   810
         TabIndex        =   239
         Top             =   3510
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   70
         Left            =   810
         TabIndex        =   238
         Top             =   3135
         Width           =   465
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   21
         Left            =   1365
         Picture         =   "frmAPOListados.frx":AE78
         ToolTipText     =   "Buscar fecha"
         Top             =   2625
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   20
         Left            =   1365
         Picture         =   "frmAPOListados.frx":AF03
         ToolTipText     =   "Buscar fecha"
         Top             =   2265
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   69
         Left            =   795
         TabIndex        =   237
         Top             =   2625
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   67
         Left            =   810
         TabIndex        =   236
         Top             =   2265
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   66
         Left            =   390
         TabIndex        =   235
         Top             =   1965
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Traspaso de Aportaciones"
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
         Index           =   19
         Left            =   375
         TabIndex        =   234
         Top             =   315
         Width           =   5160
      End
   End
   Begin VB.Frame FrameIntTesorQua 
      Height          =   7530
      Left            =   0
      TabIndex        =   144
      Top             =   0
      Width           =   6555
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   5
         Left            =   5370
         TabIndex        =   168
         Top             =   7005
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepIntTesQua 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4230
         TabIndex        =   167
         Top             =   7005
         Width           =   975
      End
      Begin VB.Frame Frame7 
         Caption         =   "Datos de Selección"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3315
         Left            =   120
         TabIndex        =   153
         Top             =   780
         Width           =   6315
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   48
            Left            =   2925
            Locked          =   -1  'True
            TabIndex        =   179
            Text            =   "Text5"
            Top             =   1950
            Width           =   3165
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   43
            Left            =   2925
            Locked          =   -1  'True
            TabIndex        =   178
            Text            =   "Text5"
            Top             =   1575
            Width           =   3165
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   48
            Left            =   1860
            MaxLength       =   10
            TabIndex        =   159
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   1935
            Width           =   1050
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   43
            Left            =   1860
            MaxLength       =   10
            TabIndex        =   158
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   1575
            Width           =   1050
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   47
            Left            =   1860
            MaxLength       =   10
            TabIndex        =   161
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   2880
            Width           =   1050
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   46
            Left            =   1860
            MaxLength       =   10
            TabIndex        =   160
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   2520
            Width           =   1050
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   45
            Left            =   1860
            MaxLength       =   16
            TabIndex        =   157
            Top             =   930
            Width           =   1050
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   44
            Left            =   1860
            MaxLength       =   16
            TabIndex        =   156
            Top             =   570
            Width           =   1050
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   44
            Left            =   2940
            Locked          =   -1  'True
            TabIndex        =   155
            Text            =   "Text5"
            Top             =   570
            Width           =   3165
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   45
            Left            =   2940
            Locked          =   -1  'True
            TabIndex        =   154
            Text            =   "Text5"
            Top             =   945
            Width           =   3165
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Clase"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   40
            Left            =   240
            TabIndex        =   182
            Top             =   1260
            Width           =   390
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   26
            Left            =   1575
            MouseIcon       =   "frmAPOListados.frx":AF8E
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar clase"
            Top             =   1935
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   25
            Left            =   1575
            MouseIcon       =   "frmAPOListados.frx":B0E0
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar clase"
            Top             =   1575
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   57
            Left            =   900
            TabIndex        =   181
            Top             =   1920
            Width           =   420
         End
         Begin VB.Label Label4 
            Caption         =   "Desde"
            Height          =   195
            Index           =   56
            Left            =   900
            TabIndex        =   180
            Top             =   1560
            Width           =   465
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Aportacion"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   55
            Left            =   210
            TabIndex        =   174
            Top             =   2250
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Desde"
            Height          =   195
            Index           =   54
            Left            =   900
            TabIndex        =   173
            Top             =   2550
            Width           =   465
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   53
            Left            =   900
            TabIndex        =   172
            Top             =   2880
            Width           =   420
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   15
            Left            =   1560
            Picture         =   "frmAPOListados.frx":B232
            ToolTipText     =   "Buscar fecha"
            Top             =   2880
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   14
            Left            =   1560
            Picture         =   "frmAPOListados.frx":B2BD
            ToolTipText     =   "Buscar fecha"
            Top             =   2520
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Desde"
            Height          =   195
            Index           =   52
            Left            =   930
            TabIndex        =   171
            Top             =   570
            Width           =   465
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   51
            Left            =   915
            TabIndex        =   170
            Top             =   945
            Width           =   420
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Socio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   41
            Left            =   225
            TabIndex        =   169
            Top             =   330
            Width           =   375
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   24
            Left            =   1575
            MouseIcon       =   "frmAPOListados.frx":B348
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar socio"
            Top             =   960
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   23
            Left            =   1590
            MouseIcon       =   "frmAPOListados.frx":B49A
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar socio"
            Top             =   570
            Width           =   240
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Datos para la contabilización"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2235
         Left            =   120
         TabIndex        =   145
         Top             =   4110
         Width           =   6315
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   49
            Left            =   1860
            MaxLength       =   10
            TabIndex        =   162
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   390
            Width           =   1050
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   42
            Left            =   2955
            Locked          =   -1  'True
            TabIndex        =   148
            Top             =   1470
            Width           =   3195
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   42
            Left            =   1860
            MaxLength       =   10
            TabIndex        =   165
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1470
            Width           =   1050
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   40
            Left            =   1860
            MaxLength       =   10
            TabIndex        =   164
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1110
            Width           =   1050
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   40
            Left            =   2955
            Locked          =   -1  'True
            TabIndex        =   147
            Top             =   1110
            Width           =   3195
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   34
            Left            =   1860
            MaxLength       =   10
            TabIndex        =   163
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   750
            Width           =   1050
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   33
            Left            =   2955
            Locked          =   -1  'True
            TabIndex        =   146
            Top             =   1830
            Width           =   3195
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   33
            Left            =   1860
            MaxLength       =   10
            TabIndex        =   166
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1830
            Width           =   1050
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   13
            Left            =   1590
            Picture         =   "frmAPOListados.frx":B5EC
            ToolTipText     =   "Buscar fecha"
            Top             =   390
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Aportación"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   58
            Left            =   180
            TabIndex        =   183
            Top             =   435
            Width           =   1395
         End
         Begin VB.Label Label1 
            Caption         =   "F.Pago Negativas"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   9
            Left            =   180
            TabIndex        =   152
            Top             =   1515
            Width           =   1365
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   22
            Left            =   1590
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   1470
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   21
            Left            =   1590
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   1110
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "F.Pago Positivas"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   8
            Left            =   180
            TabIndex        =   151
            Top             =   1155
            Width           =   1275
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Vencimiento"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   39
            Left            =   180
            TabIndex        =   150
            Top             =   795
            Width           =   1395
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   12
            Left            =   1590
            Picture         =   "frmAPOListados.frx":B677
            ToolTipText     =   "Buscar fecha"
            Top             =   750
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Banco Prevista"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   7
            Left            =   180
            TabIndex        =   149
            Top             =   1875
            Width           =   1395
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   20
            Left            =   1590
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   1830
            Width           =   240
         End
      End
      Begin MSComctlLib.ProgressBar Pb4 
         Height          =   255
         Left            =   120
         TabIndex        =   175
         Top             =   6720
         Visible         =   0   'False
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label3 
         Caption         =   "Integración Aportaciones Tesoreria"
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
         Left            =   180
         TabIndex        =   177
         Top             =   270
         Width           =   5160
      End
      Begin VB.Label Label1 
         Caption         =   "lb1"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   180
         TabIndex        =   176
         Top             =   6390
         Visible         =   0   'False
         Width           =   6105
      End
   End
   Begin VB.Frame FrameAporObligatoria 
      Height          =   6330
      Left            =   -30
      TabIndex        =   266
      Top             =   30
      Width           =   6555
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   74
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   279
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1245
         Width           =   1050
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   9
         Left            =   4980
         TabIndex        =   287
         Top             =   5415
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepApoObli 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   286
         Top             =   5430
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   78
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   282
         Top             =   2220
         Width           =   1035
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   77
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   281
         Top             =   1860
         Width           =   1035
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   77
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   274
         Text            =   "Text5"
         Top             =   1875
         Width           =   3285
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   78
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   273
         Text            =   "Text5"
         Top             =   2250
         Width           =   3285
      End
      Begin VB.Frame Frame14 
         BorderStyle     =   0  'None
         Height          =   2565
         Left            =   150
         TabIndex        =   267
         Top             =   2730
         Width           =   6165
         Begin MSComctlLib.ProgressBar Pb9 
            Height          =   255
            Left            =   150
            TabIndex        =   269
            Top             =   1980
            Visible         =   0   'False
            Width           =   6045
            _ExtentX        =   10663
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   73
            Left            =   1560
            MaxLength       =   12
            TabIndex        =   285
            Top             =   1500
            Width           =   1020
         End
         Begin VB.TextBox txtcodigo 
            Height          =   285
            Index           =   72
            Left            =   1560
            MaxLength       =   30
            TabIndex        =   284
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   840
            Width           =   4380
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   71
            Left            =   1560
            MaxLength       =   3
            TabIndex        =   283
            Top             =   270
            Width           =   1035
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   71
            Left            =   2670
            Locked          =   -1  'True
            TabIndex        =   268
            Text            =   "Text5"
            Top             =   270
            Width           =   3285
         End
         Begin VB.Label Label4 
            Caption         =   "Importe Aportación"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   80
            Left            =   300
            TabIndex        =   272
            Top             =   1200
            Width           =   1875
         End
         Begin VB.Label Label4 
            Caption         =   "Descripción"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   79
            Left            =   300
            TabIndex        =   271
            Top             =   630
            Width           =   1815
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   37
            Left            =   1230
            MouseIcon       =   "frmAPOListados.frx":B702
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar aportación"
            Top             =   300
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Aportación"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   78
            Left            =   300
            TabIndex        =   270
            Top             =   0
            Width           =   1125
         End
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   22
         Left            =   1365
         Picture         =   "frmAPOListados.frx":B854
         ToolTipText     =   "Buscar fecha"
         Top             =   1230
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Aportación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   81
         Left            =   450
         TabIndex        =   280
         Top             =   990
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Aportación Obligatoria"
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
         Index           =   20
         Left            =   375
         TabIndex        =   278
         Top             =   315
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   88
         Left            =   840
         TabIndex        =   277
         Top             =   1875
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   87
         Left            =   840
         TabIndex        =   276
         Top             =   2250
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   86
         Left            =   420
         TabIndex        =   275
         Top             =   1635
         Width           =   375
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   39
         Left            =   1380
         MouseIcon       =   "frmAPOListados.frx":B8DF
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   38
         Left            =   1380
         MouseIcon       =   "frmAPOListados.frx":BA31
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1860
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmAPOListados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public OpcionListado As Byte

'1 = Listado de aportaciones
'2 = Regularizacion de aportaciones
'3 = Certificado de aportaciones
'4 = Informe de aportaciones desde el mantenimineto de aportaciones

' APORTACIONES DE QUATRETONDA
'
'5 = Actualizaciones de aportaciones (dentro del mto de aportaciones de Quatretonda)
'6 = Informes de aportaciones (dentro del mto de aportaciones de Quatretonda)
'7 = Borrado masivo de aportaciones (dentro del mto de aportaciones de Quatretonda)
'8 = Integracion en tesoreria (dentro del mto de aportaciones de Quatretonda)
'17= Devolucion de aportaciones (dentro del mto de aportaciones de Quatretonda)

' OPERACIONES SOLO PARA MOGENTE
'
'9= Alta de socios (dentro del mantenimiento)
'10= Baja de socios (dentro del mantenimiento)


' APORTACIONES DE BOLBAITE
'
'11= Insercion de aportaciones de Bolbaite
'12= impresion de recibos de bolbaite
'13= Generación de aportación obligatoria
'14= Integracion a tesoreria de aportaciones en bolbaite
'15= Certificado de aportaciones
'16= Devolucion de aportaciones

Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmSoc As frmManSocios 'Socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCtas As frmCtasConta 'cuentas de contabilidad
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmFpa As frmForpaConta 'formas de pago de la contabilidad
Attribute frmFpa.VB_VarHelpID = -1
Private WithEvents frmApo As frmAPOTipos 'Tipo de Aportaciones
Attribute frmApo.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'para marcar que aportaciones queremos
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmSec As frmManSeccion 'para seleccionar
Attribute frmSec.VB_VarHelpID = -1
Private WithEvents frmVar As frmComVar 'Variedad
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmCla As frmComercial 'Clase
Attribute frmCla.VB_VarHelpID = -1
Private WithEvents frmMens1 As frmMensajes 'para marcar que variedades queremos
Attribute frmMens1.VB_VarHelpID = -1
Private WithEvents frmMens2 As frmMensajes 'para marcar que variedades queremos en informe de aportaciones de quatretonda
Attribute frmMens2.VB_VarHelpID = -1


 
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
Dim Tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

Dim vSeccion As CSeccion


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub Check3_Click()
    chkNegativas.Enabled = (Check3.Value = 1)
    If Not chkNegativas.Enabled Then chkNegativas.Value = 0
End Sub



Private Sub CmdAcepApoObli_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim I As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim Tipos As String
Dim Sql As String
Dim Sql2 As String

    
    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H socio
    cDesde = Trim(txtcodigo(77).Text)
    cHasta = Trim(txtcodigo(78).Text)
    nDesde = txtNombre(78).Text
    nHasta = txtNombre(78).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    Sql = "rsocios.fechabaja is null"
    If Not AnyadirAFormula(cadSelect, Sql) Then Exit Sub
    
    
    If HayRegistros(Tabla, cadSelect) Then
        Sql2 = "select * from raportacion where (fecaport, codaport, codsocio) in (select " & DBSet(txtcodigo(74).Text, "F") & "," & DBSet(txtcodigo(71).Text, "N") & ", codsocio from "
        Sql2 = Sql2 & Tabla
        If cadSelect <> "" Then Sql2 = Sql2 & " where " & cadSelect & ")"
        
        If TotalRegistros(Sql2) <> 0 Then
            If MsgBox("Existen aportaciones para algún socio/s de este tipo para esta fecha. " & vbCrLf & vbCrLf & " ¿ Desea continuar ? " & vbCrLf, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                Exit Sub
            End If
        End If
        If InsertarTemporal2(Tabla, cadSelect) Then
            indRPT = 83 ' "rManAportacion.rpt"
            
            cadTitulo = "Aportación Obligatoria"
        
            If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
            
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            
            CadParam = CadParam & "pTitulo=""Aportación Obligatoria""|"
            numParam = numParam + 1
            
            cadNombreRPT = nomDocu
            LlamarImprimir
            If MsgBox(" ¿ Desea continuar con el proceso ? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
    
                If InsertarAportacionesObligatoriasBolbaite(Tabla, cadSelect) Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                    cmdCancel_Click (0)
                End If
                
            End If
        End If
    End If
        

End Sub

Private Sub CmdAcepCertBol_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim I As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim vSQL As String

    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H socio
    cDesde = Trim(txtcodigo(88).Text)
    cHasta = Trim(txtcodigo(89).Text)
    nDesde = txtNombre(88).Text
    nHasta = txtNombre(89).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H Fecha
    cDesde = Trim(txtcodigo(90).Text)
    cHasta = Trim(txtcodigo(91).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.fecaport}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    'Tipo de Aportacion
    If Not AnyadirAFormula(cadFormula, "{raportacion.codaport} = " & DBSet(txtcodigo(87).Text, "N")) Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "{raportacion.codaport} = " & DBSet(txtcodigo(87).Text, "N")) Then Exit Sub
    
    Tabla = "raportacion INNER JOIN rsocios ON raportacion.codsocio = rsocios.codsocio"
    
    If HayRegistros(Tabla, cadSelect) Then
        CadParam = CadParam & "pPresi=""" & txtcodigo(92).Text & """|"
        CadParam = CadParam & "pSecre=""" & txtcodigo(93).Text & """|"
        CadParam = CadParam & "pTesor=""" & txtcodigo(94).Text & """|"
        CadParam = CadParam & "pObser=""" & txtcodigo(95).Text & """|"
        CadParam = CadParam & "pFecha=""" & txtcodigo(76).Text & """|"
        CadParam = CadParam & "pHastaFecha=""" & txtcodigo(91).Text & """|"
        numParam = numParam + 6
        
        indRPT = 74 ' "rManAportacion.rpt"
        
        cadTitulo = "Certificado de Aportaciones"
    
        If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
        
        cadNombreRPT = nomDocu
        LlamarImprimir
        If MsgBox(" ¿ Impresión correcta para actualizar ? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            If ActualizarTipo(Tabla, cadSelect) Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
            End If
        End If
    End If
End Sub

Private Function ActualizarTipo(Tabla As String, cadSelect As String) As Boolean
Dim Sql As String
Dim Nregs As Long

    On Error GoTo eActualizarTipo

    ActualizarTipo = False

    Sql = "select distinct rsocios.codsocio from " & Tabla
    Sql = Sql & " where " & cadSelect
    
    Nregs = TotalRegistrosConsulta(Sql)
    
    Sql = "update rtipoapor set numero = numero + " & DBSet(Nregs, "N")
    Sql = Sql & " where codaport = " & DBSet(txtcodigo(87).Text, "N")
    
    conn.Execute Sql
    
    ActualizarTipo = True
    Exit Function
    
eActualizarTipo:
    MuestraError Err.Number, "Actualizar Tipo", Err.Description
End Function

Private Sub CmdAcepDevApoQua_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim I As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim vSQL As String

    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'Seleccionar registros de la empresa conectada
    'D/H socio
    cDesde = Trim(txtcodigo(107).Text)
    cHasta = Trim(txtcodigo(108).Text)
    nDesde = txtNombre(107).Text
    nHasta = txtNombre(108).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raporhco.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H clase
    cDesde = Trim(txtcodigo(109).Text)
    cHasta = Trim(txtcodigo(110).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{variedades.codclase}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHClase= """) Then Exit Sub
    End If
    
    vSQL = ""
    If txtcodigo(109).Text <> "" Then vSQL = vSQL & " and clases.codclase >= " & DBSet(txtcodigo(109).Text, "N")
    If txtcodigo(110).Text <> "" Then vSQL = vSQL & " and clases.codclase <= " & DBSet(txtcodigo(110).Text, "N")
    
                
    Set frmMens2 = New frmMensajes
    
    frmMens2.OpcionMensaje = 16
    frmMens2.cadWhere = vSQL
    frmMens2.Show vbModal
    
    Set frmMens2 = Nothing
    
    
    'D/H Fecha
    cDesde = Trim(txtcodigo(106).Text)
    cHasta = Trim(txtcodigo(111).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raporhco.fecaport}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    'Ejercicio
    If Not AnyadirAFormula(cadFormula, "{raporhco.ejercicio} = " & DBSet(txtcodigo(98).Text, "N")) Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "{raporhco.ejercicio} = " & DBSet(txtcodigo(98).Text, "N")) Then Exit Sub
    
    
    Tabla = "raporhco INNER JOIN variedades ON raporhco.codvarie = variedades.codvarie "
    
    
    If HayRegistros(Tabla, cadSelect) Then
            indRPT = 83 ' "rManAportacion.rpt"
            
            cadTitulo = "Resumen Devolución de Aportaciones"
            
            If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
            
            cadNombreRPT = Replace(nomDocu, "APOInf.rpt", "APOInfAnyo.rpt")
            
            CadParam = CadParam & "pResumen=1|"
            numParam = numParam + 1

'            cadNombreRPT = nomDocu
            LlamarImprimir
            If MsgBox(" ¿ Desea continuar con el proceso ? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                If InsertarDevolucionesQua(Tabla, cadSelect) Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                End If
            End If
    End If


End Sub

Private Sub CmdAcepDevApor_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim I As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim vSQL As String

    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H socio
    cDesde = Trim(txtcodigo(103).Text)
    cHasta = Trim(txtcodigo(104).Text)
    nDesde = txtNombre(103).Text
    nHasta = txtNombre(104).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H Fecha
    cDesde = Trim(txtcodigo(101).Text)
    cHasta = Trim(txtcodigo(102).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.fecaport}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    'Tipo de Aportacion
    If Not AnyadirAFormula(cadFormula, "{raportacion.codaport} = " & DBSet(txtcodigo(105).Text, "N")) Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "{raportacion.codaport} = " & DBSet(txtcodigo(105).Text, "N")) Then Exit Sub
    
    
    
    'DAVID Agosto 2014
    'QUITAMOS del join   and rsocios.fechabaja is null
    Tabla = "raportacion INNER JOIN rsocios ON raportacion.codsocio = rsocios.codsocio " 'and rsocios.fechabaja is null "
    
    If HayRegistros(Tabla, cadSelect) Then
        If InsertarTemporal(Tabla, cadSelect) Then
            indRPT = 83 ' "rManAportacion.rpt"
            
            cadTitulo = "Devolución de Aportaciones"
            
            CadParam = CadParam & "pTitulo=""Devolución de Aportaciones""|"
            numParam = numParam + 1
        
            If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
            
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            
            cadNombreRPT = nomDocu
            LlamarImprimir
            If MsgBox(" ¿ Desea continuar con el proceso ? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                If ActualizarDevoluciones(Tabla, cadSelect) Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                End If
            End If
        End If
    End If

End Sub

Private Function ActualizarDevoluciones(vTabla As String, vSelect As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim RS As ADODB.Recordset
Dim b As Boolean
Dim SqlValues As String
Dim SqlExiste As String
Dim Importe As Currency
    
    On Error GoTo eActualizarDevoluciones
    
    ActualizarDevoluciones = False
    
    Sql = "DEVAPO" 'devolucion aportaciones
    'Bloquear para que nadie mas pueda realizarlo
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se puede realizar el proceso de Devolución de Aportaciones. Hay otro usuario realizándolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
        
    conn.BeginTrans

    Sql = "select codigo1, sum(importe2) importe from tmpinformes where codusu = " & vUsu.Codigo
    Sql = Sql & " group by codigo1 "
    Sql = Sql & " order by codigo1 "
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Sql2 = "insert into raportacion (codsocio,fecaport,codaport,descripcion,campanya,kilos,importe,codtipom,numfactu,intconta,porcaport) values "


    b = True

    Pb11.visible = True
    Pb11.Max = TotalRegistrosConsulta(Sql)
    Pb11.Value = 0
    
    SqlValues = ""
    
    While Not RS.EOF And b
        IncrementarProgresNew Pb11, 1
    
        SqlExiste = "select count(*) from raportacion where codsocio = " & DBSet(RS!Codigo1, "N") & " and codaport=" & DBSet(txtcodigo(96).Text, "N") & " and fecaport=" & DBSet(txtcodigo(100).Text, "F") & " and numfactu = 0"
        b = (TotalRegistros(SqlExiste) = 0)
        
        If Not b Then
            MsgBox "El registro para el socio " & Format(DBLet(RS!Codigo1, "N"), "000000") & " de fecha " & txtcodigo(100).Text & " y tipo " & DBSet(txtcodigo(96).Text, "N") & " existe. Revise.", vbExclamation
        Else
            Importe = DBLet(RS!Importe, "N") * (-1)
        
            SqlValues = "(" & DBSet(RS!Codigo1, "N") & "," & DBSet(txtcodigo(100).Text, "F") & "," & DBSet(txtcodigo(96).Text, "N") & "," & DBSet(txtcodigo(99).Text, "T") & ",'',0,"
            SqlValues = SqlValues & DBSet(Importe, "N") & "," & ValorNulo & ",0,0,0)"
            
            conn.Execute Sql2 & SqlValues
        End If
    
        RS.MoveNext
    Wend
    Set RS = Nothing
    
eActualizarDevoluciones:
    If Err.Number <> 0 Or Not b Then
        ActualizarDevoluciones = False
        conn.RollbackTrans
    Else
        ActualizarDevoluciones = True
        conn.CommitTrans
    End If
    
    DesBloqueoManual ("DEVAPO") 'devolucion de aportaciones
    
    Screen.MousePointer = vbDefault
    
End Function

Private Function InsertarTemporal(vTabla As String, vSelect As String) As Boolean
Dim Sql As String

    On Error GoTo eInsertarTemporal
    
    InsertarTemporal = False
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
                                            'socio, fecaport,codaport,numfactu, codtipom, importe
    Sql = "insert into tmpinformes (codusu, codigo1, fecha1, campo1, importe1, nombre1, importe2)"
    Sql = Sql & " select " & vUsu.Codigo & ", raportacion.codsocio, fecaport, codaport, numfactu, codtipom, importe "
    Sql = Sql & " from " & vTabla
    Sql = Sql & " where " & vSelect
    
    conn.Execute Sql

    InsertarTemporal = True
    Exit Function

eInsertarTemporal:
    MuestraError Err.Number, "Insertar Temporal", Err.Description
End Function



Private Function InsertarTemporal2(vTabla As String, vSelect As String) As Boolean
Dim Sql As String

    On Error GoTo eInsertarTemporal
    
    InsertarTemporal2 = False
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
                                            'socio, fecaport,codaport,numfactu, codtipom, importe
    Sql = "insert into tmpinformes (codusu, codigo1, fecha1, campo1, importe1, nombre1, importe2)"
    Sql = Sql & " select " & vUsu.Codigo & ", codsocio, " & DBSet(txtcodigo(74).Text, "F") & "," & DBSet(txtcodigo(71).Text, "N") & ", 0, null," & DBSet(txtcodigo(73).Text, "N")
    Sql = Sql & " from " & vTabla
    Sql = Sql & " where " & vSelect
    
    conn.Execute Sql

    InsertarTemporal2 = True
    Exit Function

eInsertarTemporal:
    MuestraError Err.Number, "Insertar Temporal", Err.Description
End Function




Private Sub CmdAcepInsApoBol_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim I As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim Tipos As String
Dim Sql As String

    
    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'Tipo de movimiento:
    Tipos = ""
    For I = 1 To ListView1(0).ListItems.Count
        If ListView1(0).ListItems(I).Checked Then
            Tipos = Tipos & DBSet(ListView1(0).ListItems(I).Key, "T") & ","
        End If
    Next I
    
    If Tipos = "" Then
        MsgBox "Debe seleccionar al menos un tipo de factura.", vbExclamation
        Exit Sub
    Else
        ' quitamos la ultima coma
        Tipos = "{" & Tabla & ".codtipom} in (" & Mid(Tipos, 1, Len(Tipos) - 1) & ")"
        If Not AnyadirAFormula(cadSelect, Tipos) Then Exit Sub
        Tipos = Replace(Replace(Tipos, "(", "["), ")", "]")
        If Not AnyadirAFormula(cadFormula, Tipos) Then Exit Sub
    End If
    
    'D/H socio
    cDesde = Trim(txtcodigo(66).Text)
    cHasta = Trim(txtcodigo(67).Text)
    nDesde = txtNombre(66).Text
    nHasta = txtNombre(67).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H Nro de Factura
    cDesde = Trim(txtcodigo(61).Text)
    cHasta = Trim(txtcodigo(62).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFactura= """) Then Exit Sub
    End If
    
    
    Select Case OpcionListado
    Case 11 'Insercion de aportaciones
        
        'D/H Fecha factura
        cDesde = Trim(txtcodigo(64).Text)
        cHasta = Trim(txtcodigo(65).Text)
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".fecfactu}"
            TipCod = "F"
            If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
        End If
        
        
        Sql = " not (rfactsoc.codtipom, rfactsoc.fecfactu, rfactsoc.numfactu) in (select codtipom, fecaport, numfactu from raportacion) "
        If Not AnyadirAFormula(cadSelect, Sql) Then Exit Sub
        
        If HayRegistros(Tabla, cadSelect) Then
            If InsertarAportacionesBolbaite(Tabla, cadSelect) Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                cmdCancel_Click (0)
            End If
        End If
        
    Case 12 'Impresion de recibos
        'D/H Fecha factura
        cDesde = Trim(txtcodigo(64).Text)
        cHasta = Trim(txtcodigo(65).Text)
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".fecaport}"
            TipCod = "F"
            If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
        End If
        
        CadParam = CadParam & "pFecha=""" & txtcodigo(70).Text & """|"
        numParam = numParam + 1
        
        If HayRegistros(Tabla, cadSelect) Then
            indRPT = 100 'Impresion de Recibos de aportaciones
            ConSubInforme = True
            
            cadTitulo = "Impresión de Recibos Aportaciones"
            If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
            
            cadNombreRPT = nomDocu
              
              
            LlamarImprimir
        End If
    End Select

End Sub

Private Sub CmdAcepCalApoQua_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim I As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim vSQL As String

InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
     '======== FORMULA  ====================================
    'Seleccionar registros de la empresa conectada
    'SECCION
    Codigo = "{rsocios_seccion.codsecci}=" & txtcodigo(32).Text
    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
    If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
    
    'D/H socio
    cDesde = Trim(txtcodigo(29).Text)
    cHasta = Trim(txtcodigo(30).Text)
    nDesde = txtNombre(29).Text
    nHasta = txtNombre(30).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rcampos.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'hasta el año de plantacion
    Codigo = "{rcampos.anoplant}<=" & txtcodigo(25).Text
    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
    If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
    
    
    'D/H clase
    cDesde = Trim(txtcodigo(27).Text)
    cHasta = Trim(txtcodigo(28).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{variedades.codclase}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHClase= """) Then Exit Sub
    End If
    
    vSQL = ""
    If txtcodigo(27).Text <> "" Then vSQL = vSQL & " and clases.codclase >= " & DBSet(txtcodigo(27).Text, "N")
    If txtcodigo(28).Text <> "" Then vSQL = vSQL & " and clases.codclase <= " & DBSet(txtcodigo(28).Text, "N")
    
                
    Set frmMens1 = New frmMensajes
    
    frmMens1.OpcionMensaje = 16
    frmMens1.cadWhere = vSQL
    frmMens1.Show vbModal
    
    Set frmMens1 = Nothing
    
    
    Tabla = "((rsocios INNER JOIN rcampos ON rsocios.codsocio = rcampos.codsocio and rcampos.fecbajas is null and rsocios.fechabaja is null) "
    Tabla = Tabla & " INNER JOIN rsocios_seccion ON rcampos.codsocio = rsocios_seccion.codsocio and rsocios_seccion.fecbaja is null) "
    Tabla = Tabla & " INNER JOIN variedades ON rcampos.codvarie = variedades.codvarie "
    
    If HayRegistros(Tabla, cadSelect) Then
        If CalculoAportacionQuatretonda(Tabla, cadSelect) Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
            cmdCancel_Click (3)
        End If
    End If

End Sub

Private Function CalculoAportacionQuatretonda(vTabla As String, vWhere As String) As Boolean
Dim Sql As String
Dim Importe As Currency
Dim RS As ADODB.Recordset
Dim cadErr As String
Dim NumApor As Long
Dim vTipoMov As CTiposMov
Dim b As Boolean
Dim SqlInsert As String
Dim CadValues As String
Dim CodTipoMov As String
Dim Sql2 As String
Dim devuelve As String
Dim Existe As Boolean

    On Error GoTo eCalculoAportacionQuatretonda
    
    conn.BeginTrans
    
    CalculoAportacionQuatretonda = False
    
    b = True
    
    '[Monica]15/09/2014: las aportaciones de cualquier campaña se insertarán siempre en la campaña actual
    SqlInsert = "insert into ariagro.raporhco (numaport,codsocio,codcampo,poligono,parcela,codparti,codvarie,impaport," & _
                "fecaport,anoplant,observac,supcoope,ejercicio,intconta) values "
    
    Sql = "select rcampos.* from " & vTabla
    Sql = Sql & " where " & vWhere
    
    CargarProgres Pb5, TotalRegistrosConsulta(Sql)
    Pb5.visible = True
    
    
    CadValues = ""
    CodTipoMov = "APO"
    
    Set vTipoMov = New CTiposMov
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF And b
        Sql2 = "select count(*) from ariagro.raporhco where codsocio = " & DBSet(RS!Codsocio, "N")
        Sql2 = Sql2 & " and codcampo = " & DBSet(RS!CodCampo, "N") & " and codvarie = " & DBSet(RS!codvarie, "N")
        Sql2 = Sql2 & " and fecaport = " & DBSet(txtcodigo(20).Text, "F")
        
        IncrementarProgres Pb5, 1
        DoEvents
        
        
        If TotalRegistros(Sql2) > 0 Then
            b = False
            cadErr = "Ya existe la aportación para el socio " & DBLet(RS!Codsocio, "N") & ", campo " & _
                    DBLet(RS!CodCampo, "N") & ", variedad " & DBLet(RS!codvarie, "N") & _
                    " y fecha de aportación " & txtcodigo(20).Text & ". Revise."
        Else
            Importe = Round2(Round2(DBLet(RS!supcoope, "N") / vParamAplic.Faneca, 2) * CCur(ImporteSinFormato(txtcodigo(26).Text)), 2)
        
            If Importe <> 0 Then ' no insertamos una aportacion 0
                NumApor = vTipoMov.ConseguirContador(CodTipoMov)
            
                Do
                    devuelve = DevuelveDesdeBDNew(cAgro, "raporhco", "numaport", "numaport", CStr(NumApor), "N")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (CodTipoMov)
                        NumApor = vTipoMov.ConseguirContador(CodTipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
                
                CadValues = "(" & DBSet(NumApor, "N") & "," & DBSet(RS!Codsocio, "N") & "," & DBSet(RS!CodCampo, "N") & ","
                CadValues = CadValues & DBSet(RS!poligono, "N") & "," & DBSet(RS!Parcela, "N") & "," & DBSet(RS!codparti, "N") & ","
                CadValues = CadValues & DBSet(RS!codvarie, "N") & "," & DBSet(Importe, "N") & "," & DBSet(txtcodigo(20).Text, "F") & ","
                CadValues = CadValues & DBSet(RS!anoplant, "N") & "," & ValorNulo & "," & DBSet(RS!supcoope, "N") & ","
                CadValues = CadValues & DBSet(txtcodigo(31).Text, "N") & ",0)"
                
                conn.Execute SqlInsert & CadValues
                
                b = vTipoMov.IncrementarContador(CodTipoMov)
           End If
        End If
        
        RS.MoveNext
    Wend
    
    Set RS = Nothing
    Set vTipoMov = Nothing
    
    If b Then
        CalculoAportacionQuatretonda = True
        Pb5.visible = False
        conn.CommitTrans
        Exit Function
    End If
    

eCalculoAportacionQuatretonda:
    conn.RollbackTrans
    Pb5.visible = False
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Calculo de Aportaciones de Quatretonda", Err.Description
    End If
    If Not b Then
        MsgBox "Cálculo de Aportaciones de Quatretonda:" & vbCrLf & vbCrLf & cadErr, vbExclamation
    End If
End Function


Private Sub CmdAcepIntTesBol_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim I As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim b As Boolean
Dim vSQL As String

InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    'D/H socio
    cDesde = Trim(txtcodigo(81).Text)
    cHasta = Trim(txtcodigo(82).Text)
    nDesde = txtNombre(81).Text
    nHasta = txtNombre(82).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H Fecha aportacion
    cDesde = Trim(txtcodigo(79).Text)
    cHasta = Trim(txtcodigo(80).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.fecaport}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    ' del tipo de aportacion
    If Not AnyadirAFormula(cadFormula, "{raportacion.codaport} = " & DBSet(txtcodigo(75).Text, "N")) Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "{raportacion.codaport} = " & DBSet(txtcodigo(75).Text, "N")) Then Exit Sub
    
    ' Condicion de que no esten contabilizados
    If Not AnyadirAFormula(cadFormula, "{raportacion.intconta} = 0") Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "{raportacion.intconta} = 0") Then Exit Sub
    
    Tabla = "raportacion"
    
    If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub
        
    If CargarTemporalBol(Tabla, cadSelect) Then
    
        TerminaBloquear
        
        Tabla = Tabla & " INNER JOIN tmpinformes ON raportacion.codsocio = tmpinformes.codigo1 and tmpinformes.codusu = " & vUsu.Codigo
        Tabla = Tabla & " and raportacion.fecaport = tmpinformes.fecha1 and raportacion.numfactu = tmpinformes.importe1 and (raportacion.codtipom = tmpinformes.nombre1 or raportacion.codtipom is null) "
        
        If Not BloqueaRegistro(Tabla, cadSelect) Then
            MsgBox "No se pueden Integrar en Tesoreria Aportaciones. Hay registros bloqueados.", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        ' comprobacion de que todos los socios tienen que estar en la seccion de almazara
        b = SociosEnSeccion("tmpinformes", "codusu = " & vUsu.Codigo, vParamAplic.Seccionhorto)
        If b Then b = ComprobarCtaContable_new("tmpinformes", 2, vParamAplic.Seccionhorto)
    
        If b Then
            If IntegracionAportacionesTesoreriaBolbaite(Tabla, cadSelect) Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                cmdCancel_Click (0)
            End If
        End If
        'Desbloqueamos ya no estamos contabilizando facturas
        DesBloqueoManual ("INTAPO") 'CONtabilizar facturas SOCios

    End If
    

End Sub

Private Sub CmdAcepIntTesQua_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim I As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim b As Boolean
Dim vSQL As String

InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    'D/H socio
    cDesde = Trim(txtcodigo(44).Text)
    cHasta = Trim(txtcodigo(45).Text)
    nDesde = txtNombre(44).Text
    nHasta = txtNombre(45).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raporhco.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H Fecha aportacion
    cDesde = Trim(txtcodigo(46).Text)
    cHasta = Trim(txtcodigo(47).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raporhco.fecaport}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    'D/H Clase
    cDesde = Trim(txtcodigo(43).Text)
    cHasta = Trim(txtcodigo(48).Text)
    nDesde = txtNombre(43).Text
    nHasta = txtNombre(48).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{variedades.codclase}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase= """) Then Exit Sub
    End If
    
    ' Condicion de que no esten contabilizados
    If Not AnyadirAFormula(cadFormula, "{raporhco.intconta} = 0") Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "{raporhco.intconta} = 0") Then Exit Sub
    
    vSQL = ""
    If txtcodigo(43).Text <> "" Then vSQL = vSQL & " and clases.codclase >= " & DBSet(txtcodigo(43).Text, "N")
    If txtcodigo(48).Text <> "" Then vSQL = vSQL & " and clases.codclase <= " & DBSet(txtcodigo(48).Text, "N")
    
                
    Set frmMens2 = New frmMensajes
    
    frmMens2.OpcionMensaje = 16
    frmMens2.cadWhere = vSQL
    frmMens2.Show vbModal
    
    Set frmMens2 = Nothing
    
    
    Tabla = "raporhco INNER JOIN variedades ON raporhco.codvarie = variedades.codvarie "

    If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub
        
    If CargarTemporalQua(Tabla, cadSelect) Then
    
        'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
    '    TerminaBloquear
        Tabla = "(" & Tabla & ") INNER JOIN tmpinformes ON raporhco.numaport = tmpinformes.importe1 and tmpinformes.codusu = " & vUsu.Codigo
        If Not BloqueaRegistro(Tabla, cadSelect) Then
            MsgBox "No se pueden Integrar en Tesoreria Aportaciones. Hay registros bloqueados.", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        ' Comprobaciones
        ' comprobacion de que todos los socios tienen que estar en la seccion de almazara
        b = SociosEnSeccion("tmpinformes", "codusu = " & vUsu.Codigo, vParamAplic.Seccionhorto)
        If b Then b = ComprobarCtaContable_new("tmpinformes", 1, vParamAplic.Seccionhorto)
    
        If b Then
            If IntegracionAportacionesTesoreria(Tabla, cadSelect) Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                cmdCancel_Click (0)
            End If
        End If
        'Desbloqueamos ya no estamos contabilizando facturas
        DesBloqueoManual ("INTAPO") 'CONtabilizar facturas SOCios

    End If
    
End Sub

Private Sub CmdAcepListado_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim I As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim vSQL As String

InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H socio
    cDesde = Trim(txtcodigo(23).Text)
    cHasta = Trim(txtcodigo(24).Text)
    nDesde = txtNombre(23).Text
    nHasta = txtNombre(24).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H Fecha
    cDesde = Trim(txtcodigo(21).Text)
    cHasta = Trim(txtcodigo(22).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.fecaport}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    'D/H Tipo de Aportacion
    cDesde = Trim(txtcodigo(13).Text)
    cHasta = Trim(txtcodigo(19).Text)
    nDesde = txtNombre(13).Text
    nHasta = txtNombre(19).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.codaport}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHAportacion= """) Then Exit Sub
    End If
    
    vSQL = ""
    If txtcodigo(13).Text <> "" Then vSQL = vSQL & " and rtipoapor.codaport >= " & DBSet(txtcodigo(13).Text, "N")
    If txtcodigo(19).Text <> "" Then vSQL = vSQL & " and rtipoapor.codaport <= " & DBSet(txtcodigo(19).Text, "N")
    
                
    Set frmMens = New frmMensajes
    
    frmMens.OpcionMensaje = 32
    frmMens.cadWhere = vSQL
    frmMens.Show vbModal
    
    Set frmMens = Nothing
    
    '[Monica]18/01/2016: añadimos la relacion con cooperativa
    Select Case Combo1(1).ListIndex
        Case 0 ' todos
            
        Case Else
            If Not AnyadirAFormula(cadFormula, "{rsocios.tiporelacion} = " & Combo1(1).ListIndex - 1) Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "rsocios.tiporelacion=" & Combo1(1).ListIndex - 1) Then Exit Sub
            CadParam = CadParam & "pRelacion=" & Combo1(1).ListIndex & "|"
            numParam = numParam + 1
    End Select
    
    CadParam = CadParam & "pResumen=" & chkResumen(0).Value & "|"
    numParam = numParam + 1
    
    Tabla = "raportacion INNER JOIN rsocios ON raportacion.codsocio = rsocios.codsocio and rsocios.fechabaja is null "
    
    If HayRegistros(Tabla, cadSelect) Then
        indRPT = 101 ' "rManAportacion.rpt"
        
        cadTitulo = "Informe Aportaciones"
    
        If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
        
        cadNombreRPT = nomDocu
        LlamarImprimir
    
    End If

End Sub

Private Sub CmdAcepListAporQua_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim I As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim vSQL As String

InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H socio
    cDesde = Trim(txtcodigo(36).Text)
    cHasta = Trim(txtcodigo(37).Text)
    nDesde = txtNombre(36).Text
    nHasta = txtNombre(37).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raporhco.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H Fecha
    cDesde = Trim(txtcodigo(35).Text)
    cHasta = Trim(txtcodigo(41).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raporhco.fecaport}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    If OpcionListado = 6 Then
        'D/H Clase
        cDesde = Trim(txtcodigo(38).Text)
        cHasta = Trim(txtcodigo(39).Text)
        nDesde = txtNombre(38).Text
        nHasta = txtNombre(39).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{variedades.codclase}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase= """) Then Exit Sub
        End If
    End If
    
    vSQL = ""
    If txtcodigo(38).Text <> "" Then vSQL = vSQL & " and variedades.codclase >= " & DBSet(txtcodigo(38).Text, "N")
    If txtcodigo(39).Text <> "" Then vSQL = vSQL & " and variedades.codclase <= " & DBSet(txtcodigo(39).Text, "N")
    
                
    Set frmMens2 = New frmMensajes
    
    frmMens2.OpcionMensaje = 16
    frmMens2.cadWhere = vSQL
    frmMens2.Show vbModal
    
    Set frmMens2 = Nothing
    
    
    If OpcionListado = 6 Then ' solo en el caso del listado
        Select Case Combo1(0).ListIndex
            Case 0
                ' Condicion de que no esten contabilizados
                If Not AnyadirAFormula(cadFormula, "{raporhco.intconta} = 0") Then Exit Sub
                If Not AnyadirAFormula(cadSelect, "{raporhco.intconta} = 0") Then Exit Sub
            Case 1
                ' Condicion de que esten contabilizados
                If Not AnyadirAFormula(cadFormula, "{raporhco.intconta} = 1") Then Exit Sub
                If Not AnyadirAFormula(cadSelect, "{raporhco.intconta} = 1") Then Exit Sub
            Case 2
            
        End Select
    End If
    
    Tabla = "(raporhco INNER JOIN variedades ON raporhco.codvarie = variedades.codvarie) "
    
    If HayRegistros(Tabla, cadSelect) Then
        Select Case OpcionListado
            Case 6
                indRPT = 83 'informe de APORTACIONES para Quatretonda
            
                If Not PonerParamRPT(indRPT, CadParam, 1, nomDocu) Then Exit Sub
                                   
                cadNombreRPT = nomDocu '"rAPOInf.rpt"
                
                cadTitulo = "Informe Aportaciones"
                
                
                '[Monica]24/01/2012: salta página por socio, nuevo report
                If Check2.Value Then
                    cadNombreRPT = Replace(cadNombreRPT, "APOInf.rpt", "APOInfSocio.rpt")
                    cadTitulo = cadTitulo & " por Socio "
                    '[Monica]18/09/2014: añado lo del resumen cuando es por socio por las devoluciones
                    CadParam = CadParam & "pResumen=" & Me.Check1.Value & "|"
                    numParam = numParam + 1
                Else
                    If Me.Opcion1(0).Value Then
                        cadNombreRPT = Replace(cadNombreRPT, "APOInf.rpt", "APOInfAnyo.rpt")
                        cadTitulo = cadTitulo & " por Año "
                        
                        CadParam = CadParam & "pResumen=" & Me.Check1.Value & "|"
                        numParam = numParam + 1
                    End If
                End If
                
                
                frmImprimir.NombreRPT = cadNombreRPT
                cadTitulo = "Informe Aportaciones"
                LlamarImprimir
            
            Case 7 ' borrado masivo de aportaciones
                If BorradoMasivoAporQua(Tabla, cadSelect) Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                    cmdCancel_Click (0)
                End If
        End Select
    End If
End Sub

Private Sub CmdAcepRegBajaSocios_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim I As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim b As Boolean
Dim Sql As String

Dim vCampAnt As CCampAnt

    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    ' socios dados de alta durante la campaña anterior
    Sql = "rsocios.codsocio = " & DBSet(txtcodigo(59).Text, "N")
    If Not AnyadirAFormula(cadSelect, Sql) Then Exit Sub
    
    Tabla = "rsocios"
    
    If HayRegistros(Tabla, cadSelect) Then
        Me.Label1(1).Caption = "Cargando tabla temporal"
        Me.Label1(1).visible = True
        Me.Refresh
        DoEvents
        If CargarTablaTemporal3(Tabla, cadSelect, "0", Me.Pb7) Then
            Label1(1).Caption = "Comprobando Socios en Sección"
            Label1(1).visible = True
            Me.Refresh
            DoEvents
            ' comprobacion de que todos los socios tienen que estar en la seccion de almazara
            b = ComprobarCtaContable_new("tmpinformes", 1, vParamAplic.SeccionAlmaz)
            If b Then
                Me.Label1(1).visible = True
                Me.Label1(1).Caption = "Actualizando Regularización"
                Me.Refresh
                DoEvents
                If ActualizarRegularizacionBajaSocio() Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                    cmdCancel_Click (0)
                End If
            End If
        End If
     End If
    
End Sub

Private Sub CmdAcepRegul_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim I As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim b As Boolean

InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    'D/H socio
    cDesde = Trim(txtcodigo(10).Text)
    cHasta = Trim(txtcodigo(11).Text)
    nDesde = txtNombre(10).Text
    nHasta = txtNombre(11).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H Fecha factura
    cDesde = Trim(txtcodigo(8).Text)
    cHasta = Trim(txtcodigo(9).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.fecaport}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    Tabla = "raportacion INNER JOIN rsocios ON raportacion.codsocio = rsocios.codsocio and rsocios.fechabaja is null "

    If HayRegistros(Tabla, cadSelect) Then
        Me.Label1(1).Caption = "Cargando tabla temporal"
        Me.Label1(1).visible = True
        Me.Refresh
        DoEvents
        If CargarTablaTemporal(Tabla, cadSelect, txtcodigo(4).Text, txtcodigo(5).Text, Me.pb2) Then
            Label1(1).Caption = "Comprobando Socios en Sección"
            Label1(1).visible = True
            Me.Refresh
            DoEvents
            ' comprobacion de que todos los socios tienen que estar en la seccion de almazara
            b = SociosEnSeccion("tmpinformes", "tmpinformes.codusu=" & vUsu.Codigo, vParamAplic.SeccionAlmaz)
            If b Then b = ComprobarCtaContable_new("tmpinformes", 1, vParamAplic.SeccionAlmaz)
            If b Then
                Me.Label1(1).visible = True
                Me.Label1(1).Caption = "Actualizando Regularización"
                Me.Refresh
                DoEvents
                If ActualizarRegularizacion Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                    cmdCancel_Click (0)
                End If
            End If
        End If
    End If

End Sub

Private Function SociosEnSeccion(vTabla As String, vWhere As String, Seccion As Integer) As Boolean
Dim Sql As String
Dim RS As ADODB.Recordset
Dim b As Boolean

    On Error GoTo ESocSec

    SociosEnSeccion = False

    'Seleccionamos los distintos socios, cuentas que vamos a facturar
    Sql = "SELECT DISTINCT " & vTabla & ".codigo1 codsocio"
    Sql = Sql & " from " & vTabla
    If vWhere <> "" Then Sql = Sql & " where " & vWhere
    Sql = Sql & " order by 1 "

    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    b = True

    While Not RS.EOF And b
        Sql = "select * from rsocios_seccion where codsocio = " & DBSet(RS!Codsocio, "N") & " and codsecci = " & DBSet(Seccion, "N")

        If Not (RegistrosAListar(Sql, cAgro) > 0) Then
        'si no lo encuentra
            b = False 'no encontrado
            Sql = "El Socio " & Format(RS!Codsocio, "000000") & " no tiene registro en la seccion " & Seccion
        End If

        RS.MoveNext
    Wend

    If Not b Then
        Sql = "Comprobando Socios en Seccion.. " & vbCrLf & vbCrLf & Sql

        MsgBox Sql, vbExclamation
        SociosEnSeccion = False
    Else
        SociosEnSeccion = True
    End If
    
    Exit Function

ESocSec:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Socios en Sección", Err.Description
    End If
End Function

Private Function ActualizarRegularizacion()
Dim Sql As String
Dim RS As ADODB.Recordset
Dim Sql2 As String
Dim SqlValues As String
Dim Descripc As String
Dim Campanya As String
Dim ImporIni As Currency
Dim Importe As Currency
Dim b As Boolean
Dim MensError As String
Dim SqlExiste As String

    On Error GoTo eActualizarRegularizacion
        
        
    Sql = "REGAPO" 'regularizacion de aportaciones
    'Bloquear para que nadie mas pueda realizarlo
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se puede realizar el proceso de Regularización de Aportaciones. Hay otro usuario realizándolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
        
    conn.BeginTrans
    ConnConta.BeginTrans

    Sql = "select * from tmpinformes where codusu = " & vUsu.Codigo
    Sql = Sql & " order by codigo1 "
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Sql2 = "insert into raportacion (codsocio,fecaport,codaport,descripcion,campanya,kilos,importe) values "

    Campanya = Mid(Format(Year(CDate(txtcodigo(8).Text)) + 1, "0000"), 3, 2) & "/" & Mid(Format(Year(CDate(txtcodigo(9).Text)), "0000"), 3, 2)
    Descripc = "ACUMULADA " & Campanya

    b = True

    pb2.visible = True
    pb2.Max = TotalRegistrosConsulta(Sql)
    pb2.Value = 0
    
    While Not RS.EOF And b
        IncrementarProgresNew pb2, 1
    
        SqlValues = ""
        
        Sql = "select importe from raportacion where codsocio=" & DBSet(RS!Codigo1, "N") & " and codaport=0 and fecaport=" & DBSet(txtcodigo(8).Text, "F")
    
        ImporIni = DevuelveValor(Sql)
        Importe = ImporIni + DBLet(RS!importe4, "N")
    
        SqlExiste = "select count(*) from raportacion where codsocio = " & DBSet(RS!Codigo1, "N") & " and codaport=0 and fecaport=" & DBSet(txtcodigo(14).Text, "F")
        b = (TotalRegistros(SqlExiste) = 0)
        
        If Not b Then
            MsgBox "El registro para el socio " & Format(DBLet(RS!Codigo1, "N"), "000000") & " de fecha " & DBSet(txtcodigo(9).Text, "F") & " y tipo 0 existe. Revise.", vbExclamation
        Else
            
            '[Monica]27/10/2015: en el caso de que el socio no quiera devolucion grabamos el registro de acumulado anterior
            If NoDevolverAporSocio(CStr(RS!Codigo1)) Then
                SqlValues = "(" & DBSet(RS!Codigo1, "N") & "," & DBSet(txtcodigo(14).Text, "F") & ",0," & DBSet(Descripc, "T") & ","
                SqlValues = SqlValues & DBSet(Campanya, "T") & "," & DBSet(RS!importe1, "N") & "," & DBSet(ImporIni, "N") & ")"
            Else
                ' como estaba
                SqlValues = "(" & DBSet(RS!Codigo1, "N") & "," & DBSet(txtcodigo(14).Text, "F") & ",0," & DBSet(Descripc, "T") & ","
                SqlValues = SqlValues & DBSet(Campanya, "T") & "," & DBSet(RS!importe2, "N") & "," & DBSet(Importe, "N") & ")"
            End If
            
            conn.Execute Sql2 & SqlValues
            
            MensError = "Insertando cobro en tesoreria"
            If NoDevolverAporSocio(CStr(RS!Codigo1)) Then
                b = True
            Else
                b = InsertarEnTesoreriaNewAPO(MensError, RS!Codigo1, DBLet(RS!importe4, "N"), txtcodigo(15).Text, txtcodigo(17).Text, txtcodigo(16).Text, txtcodigo(18).Text, txtcodigo(14).Text, 0)
            End If
        End If
    
    
        RS.MoveNext
    Wend
    Set RS = Nothing
    
eActualizarRegularizacion:
    If Err.Number <> 0 Or Not b Then
        ActualizarRegularizacion = False
        conn.RollbackTrans
        ConnConta.RollbackTrans
    Else
        ActualizarRegularizacion = True
        conn.CommitTrans
        ConnConta.CommitTrans
    End If
    
    DesBloqueoManual ("REGAPO") 'regularizacion de aportaciones
    
    Screen.MousePointer = vbDefault
    
End Function

Private Function NoDevolverAporSocio(Socio As String) As Boolean
Dim Sql As String

    Sql = "select nodevolverapor from rsocios where codsocio = " & DBSet(Socio, "N")
    NoDevolverAporSocio = (DevuelveValor(Sql) = 1)

End Function

Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim I As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H socio
    cDesde = Trim(txtcodigo(0).Text)
    cHasta = Trim(txtcodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H Fecha factura
    cDesde = Trim(txtcodigo(2).Text)
    cHasta = Trim(txtcodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.fecaport}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    
    Tabla = "raportacion INNER JOIN rsocios ON raportacion.codsocio = rsocios.codsocio and rsocios.fechabaja is null "
    
    
    If HayRegistros(Tabla, cadSelect) Then
        If CargarTablaTemporal(Tabla, cadSelect, txtcodigo(6).Text, txtcodigo(7).Text, Me.pb1) Then
            '[Monica]20/01/2016: si es mogente solo las de regularizacion negaiva
            If chkNegativas.Value = 1 Then BorrarPositivas
        
            cadFormula = ""
            If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
            
            If Not HayRegistros("tmpinformes", "tmpinformes.codusu = " & vUsu.Codigo) Then Exit Sub
            
            Select Case OpcionListado
                Case 1 'Informe de aportaciones
                    'Nombre fichero .rpt a Imprimir
                    indRPT = 70 'informe de APORTACIONES
                
                    If Not PonerParamRPT(indRPT, "", 1, nomDocu) Then Exit Sub
                                       
                    cadNombreRPT = nomDocu '"rAPOInforme.rpt"
                    
                    frmImprimir.NombreRPT = cadNombreRPT
                    
                    cadTitulo = "Informe Aportaciones"
                    LlamarImprimir
                Case 3 ' Certificado de aportaciones
                    CadParam = CadParam & "pDesdeFecha=""" & txtcodigo(2).Text & """|"
                    CadParam = CadParam & "pHastaFecha=""" & txtcodigo(3).Text & """|"
                    CadParam = CadParam & "pFecha=""" & txtcodigo(12).Text & """|"
                    numParam = numParam + 3
                    '[Monica]11/03/2015:imprimimos el resumen
                    If vParamAplic.Cooperativa = 3 Then
                        CadParam = CadParam & "pResumen=" & Check3.Value & "|"
                        numParam = numParam + 1
                    End If
                    indRPT = 74 'certificado de APORTACIONES
                
                    If Not PonerParamRPT(indRPT, "", 1, nomDocu) Then Exit Sub
                                       
                    cadNombreRPT = nomDocu '"rAPOCertificado.rpt"
                    
                    frmImprimir.NombreRPT = cadNombreRPT
                    
                    cadTitulo = "Certificado de Aportaciones"
                    LlamarImprimir
            End Select
        End If
    End If
End Sub

Private Sub BorrarPositivas()
Dim Sql As String

    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo & " and importe4 > 0 "
    conn.Execute Sql
    

End Sub


Private Function CargarTablaTemporal(nTabla1 As String, nSelect1 As String, Precio1 As String, Precio2 As String, ByRef pb1 As ProgressBar) As Boolean
Dim RS As ADODB.Recordset
Dim Sql As String
Dim Sql2 As String
Dim cValues As String
Dim AcumAnt As Long
Dim Kilos As Long
Dim KilosMed As Long
Dim Nregs As Integer
Dim SocioAnt As Long
Dim NombreAnt As String
Dim Diferencia As Long
Dim Entro As Boolean
Dim Importe As Currency


    On Error GoTo eCargarTablaTemporal

    If ExistenRegistrosAcumulados(nTabla1, nSelect1) Then
        CargarTablaTemporal = False
        Exit Function
    End If

    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql



    Sql = "insert into tmpinformes (codusu, codigo1, nombre1, importe1, importe2, importe3, precio1, importe4) values "
    
    Sql2 = " select raportacion.codsocio, nomsocio, fecaport, codaport, kilos "
    Sql2 = Sql2 & " from " & nTabla1
    
    If nSelect1 <> "" Then Sql2 = Sql2 & " where  " & nSelect1
    Sql2 = Sql2 & " order by 1, 3, 4"
    
    
    pb1.visible = True
    pb1.Max = TotalRegistrosConsulta(Sql2)
    pb1.Value = 0
    
    
    cValues = ""
    
    Set RS = New ADODB.Recordset
    RS.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        SocioAnt = RS.Fields(0).Value
        NombreAnt = RS.Fields(1).Value
        
        Kilos = 0
        Nregs = 0
        AcumAnt = 0
    End If
    
    Entro = False
    
    While Not RS.EOF
        Entro = True
        
        pb1.Value = pb1.Value + 1
        DoEvents
        
        If SocioAnt <> RS.Fields(0).Value Then
            cValues = cValues & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NombreAnt, "T") & ","
            
            If Nregs <> 0 Then
                KilosMed = Round2(Kilos / Nregs, 0)
            Else
                KilosMed = 0
            End If
            
            cValues = cValues & DBSet(AcumAnt, "N") & "," & DBSet(KilosMed, "N") & ","
        
            Diferencia = KilosMed - AcumAnt
            
            cValues = cValues & DBSet(Diferencia, "N") & ","
            
            If Diferencia > 0 Then
                Importe = Round2(Diferencia * ImporteSinFormato(Precio1), 2)
                cValues = cValues & DBSet(ImporteSinFormato(Precio1), "N") & ","
            Else
                Importe = Round2(Diferencia * ImporteSinFormato(Precio2), 2)
                cValues = cValues & DBSet(ImporteSinFormato(Precio2), "N") & ","
            End If
            cValues = cValues & DBSet(Importe, "N") & "),"
            
            Kilos = 0
            Nregs = 0
            AcumAnt = 0
            
            SocioAnt = RS.Fields(0).Value
            NombreAnt = RS.Fields(1).Value
        
        End If
    
        If RS!Codaport = 0 Then
            AcumAnt = RS!Kilos
            Nregs = 0
        Else
            Kilos = Kilos + RS!Kilos
            Nregs = Nregs + 1
        End If
        
        RS.MoveNext
    Wend
    ' el ultimo registro no se ha grabado
    
    If Entro Then
        cValues = cValues & "(" & vUsu.Codigo & "," & DBSet(SocioAnt, "N") & "," & DBSet(NombreAnt, "T") & ","
        If Nregs <> 0 Then
            KilosMed = Round2(Kilos / Nregs, 0)
        Else
            KilosMed = 0
        End If
        
        cValues = cValues & DBSet(AcumAnt, "N") & "," & DBSet(KilosMed, "N") & ","
    
        Diferencia = KilosMed - AcumAnt
        cValues = cValues & DBSet(Diferencia, "N") & ","
        
        If Diferencia > 0 Then
            Importe = Round2(Diferencia * ImporteSinFormato(Precio1), 2)
            cValues = cValues & DBSet(ImporteSinFormato(Precio1), "N") & ","
        Else
            Importe = Round2(Diferencia * ImporteSinFormato(Precio2), 2)
            cValues = cValues & DBSet(ImporteSinFormato(Precio2), "N") & ","
        End If
        cValues = cValues & DBSet(Importe, "N") & "),"
    
        Kilos = 0
        Nregs = 0
        AcumAnt = 0
    End If

    If cValues <> "" Then
        cValues = Mid(cValues, 1, Len(cValues) - 1)
        conn.Execute Sql & cValues
    End If

    Set RS = Nothing

    CargarTablaTemporal = True
    pb1.visible = False
    
    Exit Function
    
eCargarTablaTemporal:
    MuestraError Err.Number, "Cargar Tabla Temporal", Err.Description
End Function

Private Function ExistenRegistrosAcumulados(nTabla As String, nWhere As String) As Boolean
Dim RS As ADODB.Recordset
Dim Sql As String
Dim I As Long
Dim cadMen As String
Dim cad As String


    On Error GoTo eExistenRegistrosAcumulados
    
    ExistenRegistrosAcumulados = False
    
    Sql = "select raportacion.codsocio, count(*) from " & nTabla
    Sql = Sql & " where codaport = 0 "
    If nWhere <> "" Then
        Sql = Sql & " and " & nWhere
    End If
    Sql = Sql & " group by 1 "
    Sql = Sql & " having count(*) > 1"
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        cadMen = "Los siguientes socios tienen más de un registro de acumulado anterior entre las fechas: "
        I = 0
        While Not RS.EOF
            I = I + 1
            cad = cad & Format(RS.Fields(0), "000000") & ","
            
            If I = 10 Then
                cad = cad & vbCrLf
                I = 0
            End If
            
            RS.MoveNext
        Wend
        
    End If
    Set RS = Nothing
    
    ExistenRegistrosAcumulados = False
    
    Exit Function
    
eExistenRegistrosAcumulados:
    MuestraError Err.Number, "Existen Registros Acumulados", Err.Description
End Function



Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub


Private Sub CmdAcepRegAltaSocios_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim I As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim b As Boolean
Dim Sql As String


InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    ' socios dados de alta durante la campaña
    Sql = "((rsocios.fechaalta between " & DBSet(vParam.FecIniCam, "F") & " and " & DBSet(vParam.FecFinCam, "F") & ") or "
    Sql = Sql & " rsocios.codsocio in (select codsocio from rsocios_seccion where codsecci = " & DBSet(vParamAplic.SeccionAlmaz, "N")
    Sql = Sql & " and fecalta between " & DBSet(vParam.FecIniCam, "F") & " and " & DBSet(vParam.FecFinCam, "F")
    Sql = Sql & " and fecbaja is null)) "
    
    If Not AnyadirAFormula(cadSelect, Sql) Then Exit Sub
    
    
    Sql = "rsocios.fechabaja is null"
    If Not AnyadirAFormula(cadSelect, Sql) Then Exit Sub
    
    Sql = "rsocios.codsocio in (select codsocio from (rcampos inner join variedades on rcampos.codvarie = variedades.codvarie) "
    Sql = Sql & " inner join productos on variedades.codprodu = productos.codprodu "
    Sql = Sql & " where productos.codgrupo = 5) "
    If Not AnyadirAFormula(cadSelect, Sql) Then Exit Sub
    
    
    Tabla = "rsocios"
    
    
    If HayRegistros(Tabla, cadSelect) Then
        Me.Label1(1).Caption = "Cargando tabla temporal"
        Me.Label1(1).visible = True
        Me.Refresh
        DoEvents
        If CargarTablaTemporal2(Tabla, cadSelect, txtcodigo(60).Text, Me.Pb6) Then
            Label1(1).Caption = "Comprobando Socios en Sección"
            Label1(1).visible = True
            Me.Refresh
            DoEvents
            ' comprobacion de que todos los socios tienen que estar en la seccion de almazara
            b = SociosEnSeccion("tmpinformes", "tmpinformes.codusu=" & vUsu.Codigo, vParamAplic.SeccionAlmaz)
            If b Then b = ComprobarCtaContable_new("tmpinformes", 1, vParamAplic.SeccionAlmaz)
            If b Then
                Me.Label1(1).visible = True
                Me.Label1(1).Caption = "Actualizando Regularización"
                Me.Refresh
                DoEvents
                If ActualizarRegularizacionAltaSocio(txtcodigo(60).Text) Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                    cmdCancel_Click (0)
                End If
            End If
        End If
     End If


End Sub



Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
    
        Select Case OpcionListado
            Case 1 ' informe de aportaciones
                PonerFoco txtcodigo(0)
                txtcodigo(3).Text = Format(vParam.FecFinCam, "dd/mm/yyyy")
            Case 2 ' regularizacion
                txtcodigo(9).Text = Format(vParam.FecFinCam, "dd/mm/yyyy")
                txtcodigo(14).Text = Format(DateAdd("d", 1, vParam.FecFinCam), "dd/mm/yyyy")
            
                '[Monica]30/01/2014: valores por defecto de las formas de pago
                txtcodigo(16).Text = Format(vParamAplic.ForpaNega, "000")
                txtCodigo_LostFocus (16)
                txtcodigo(17).Text = Format(vParamAplic.ForpaPosi, "000")
                txtCodigo_LostFocus (17)
            
                PonerFoco txtcodigo(10)
            Case 3 ' Certificado de Aportaciones
                PonerFoco txtcodigo(0)
                txtcodigo(3).Text = Format(vParam.FecFinCam, "dd/mm/yyyy")
                txtcodigo(12).Text = Format(Now, "dd/mm/yyyy")
            Case 4 ' Informe de Aportaciones en el mantenimiento
                PonerFoco txtcodigo(23)
                '[Monica]18/01/2016: Añadimos la relacion con la cooperativa
                Combo1(1).ListIndex = 0
            Case 5 ' calculo de aportaciones de quatretonda
                PonerFoco txtcodigo(32)
            Case 6 ' listado de aportaciones para quatretonda
                Opcion1(0).Value = True
                PonerFoco txtcodigo(33)
                Combo1(0).ListIndex = 0
            Case 7 ' borrrado masivo de aportaciones de quatretonda
                PonerFoco txtcodigo(44)
            Case 8 ' integracion en tesoreria de quatretonda
                PonerFoco txtcodigo(44)
                
                '[Monica]30/01/2014: valores por defecto de las formas de pago
                txtcodigo(40).Text = Format(vParamAplic.ForpaNega, "000")
                txtCodigo_LostFocus (40)
                txtcodigo(42).Text = Format(vParamAplic.ForpaPosi, "000")
                txtCodigo_LostFocus (42)

            Case 9 ' integracion en tesoreria alta de socios moixent
                PonerFoco txtcodigo(60)
            
                '[Monica]30/01/2014: valores por defecto de las formas de pago
                txtcodigo(52).Text = Format(vParamAplic.ForpaNega, "000")
                txtCodigo_LostFocus (52)
                txtcodigo(53).Text = Format(vParamAplic.ForpaPosi, "000")
                txtCodigo_LostFocus (53)
            
            Case 10 ' integracion en tesoreria baja de socios moixent
                PonerFoco txtcodigo(59)
                
                '[Monica]30/01/2014: valores por defecto de las formas de pago
                txtcodigo(56).Text = Format(vParamAplic.ForpaNega, "000")
                txtCodigo_LostFocus (56)
                txtcodigo(55).Text = Format(vParamAplic.ForpaPosi, "000")
                txtCodigo_LostFocus (55)
                
            Case 11 ' Insercion de aportaciones de Bolbaite
                PonerFoco txtcodigo(61)
                txtcodigo(69).Text = vParamAplic.PorcenAFO ' por defecto
                If txtcodigo(69).Text <> "" Then txtcodigo(69).Text = Format(txtcodigo(69).Text, "##0.00")
            
            Case 12 ' Impresion de Recibos de Bolbaite
                PonerFoco txtcodigo(61)
                txtcodigo(70).Text = Format(Now, "dd/mm/yyyy")
                
            Case 13 ' Aportacion obligatoria
                PonerFoco txtcodigo(74)
                txtcodigo(74).Text = Format(Now, "dd/mm/yyyy")
                
                
            Case 14 ' integracion a contabilidad de aportaciones bolbaite
                PonerFoco txtcodigo(81)
                txtcodigo(86).Text = Format(Now, "dd/mm/yyyy")
                
                '[Monica]30/01/2014:
                txtcodigo(85).Text = Format(vParamAplic.ForpaNega, "000")
                txtCodigo_LostFocus (85)
                txtcodigo(84).Text = Format(vParamAplic.ForpaPosi, "000")
                txtCodigo_LostFocus (84)
                
            Case 15 ' certificado de retenciones
                PonerFoco txtcodigo(88)
                
            Case 16 ' devolucion de aportaciones
                PonerFoco txtcodigo(103)
                
            Case 17 ' devolucion de aportaciones de quatretonda
                PonerFoco txtcodigo(107)
                
        End Select
        Screen.MousePointer = vbDefault
    
    End If
End Sub

Private Sub Form_Load()
Dim h As Integer, w As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
    For h = 0 To 29
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next h
    For h = 33 To 52
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next h
    For h = 54 To 57
        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next h
    
    
    For h = 0 To imgAyuda.Count - 1
        imgAyuda(h).Picture = frmPpal.ImageListB.ListImages(10).Picture
    Next h


    indFrame = 5
    Me.FrameCobros.visible = False
    Me.FrameRegularizacion.visible = False
    Me.FrameInforme.visible = False
    Me.FrameCalculoAporQua.visible = False
    Me.FrameListAporQua.visible = False
    Me.FrameIntTesorQua.visible = False
    Me.FrameRegAltaSocios.visible = False
    Me.FrameRegBajaSocios.visible = False
    Me.FrameInsertarApoBol.visible = False
    Me.FrameAporObligatoria.visible = False
    Me.FrameCertificadoBol.visible = False
    Me.FrameIntTesorBol.visible = False
    Me.FrameDevolAporBol.visible = False
    Me.FrameDevolAporQua.visible = False
    
    Select Case OpcionListado
        Case 1 ' rendimiento por articulo
            FrameCobrosVisible True, h, w
            Tabla = "raportacion"
            Me.pb1.visible = False
            Frame1.visible = False
            Frame1.Enabled = False
            Label1(0).Caption = "Informe de Aportaciones"
        
        Case 2 ' regularizacion
            ConexionConta vParamAplic.SeccionAlmaz
        
            FrameRegularizacionVisible True, h, w
            Tabla = "raportacion"
            Me.pb1.visible = False
            
        Case 3 ' Certificado de aportaciones
            FrameCobrosVisible True, h, w
            Tabla = "raportacion"
            Me.pb1.visible = False
            Frame1.visible = True
            Frame1.Enabled = True
            Label1(0).Caption = "Certificado de Aportaciones"
            
            '[Monica]11/03/2015: para el caso de Mogente imprimen o no resumen
            Me.Check3.Enabled = (vParamAplic.Cooperativa = 3)
            Me.Check3.visible = (vParamAplic.Cooperativa = 3)
            '[Monica]20/01/2016: para el caso de Mogente cuando es resumen pueden o no imprimir solo las que
            '                    tienen una regularizacion negativa
            Me.chkNegativas.Enabled = False
            Me.chkNegativas.visible = (vParamAplic.Cooperativa = 3)
            
        Case 4 ' Informe de aportaciones
            FrameInformesVisible True, h, w
            Tabla = "raportacion"
            Me.pb1.visible = False
            Label1(0).Caption = "Certificado de Aportaciones"
            
            CargaCombo
                
        Case 5 ' Cálculo de Aportaciones de Quatretonda
            FrameCalculoAporQuaVisible True, h, w
            Tabla = "rcampos"
            Me.pb1.visible = False
            Label1(0).Caption = "Cálculo de Aportaciones"
    
        Case 6 ' Listado de aportaciones para quatretonda
            FrameListAporQuaVisible True, h, w
            Tabla = "raporhco"
            Me.pb1.visible = False
            CargaCombo
                    
        Case 7 ' borrado masivo
            FrameListAporQuaVisible True, h, w
            Tabla = "raporhco"
            Label1(6).Caption = "Borrado Masivo de Aportaciones"
            
            Frame4.visible = False
            Frame4.Enabled = False
            Check1.visible = False
            Check1.Enabled = False
            Label4(59).visible = False
            Combo1(0).visible = False
            Combo1(0).Enabled = False
            Check2.Enabled = False
            Check2.visible = False
            imgAyuda(0).Enabled = False
            imgAyuda(0).visible = False
            
            
        Case 17 ' Devolucion de aportaciones para quatretonda
            FrameDevolAporQuaVisible True, h, w
            Tabla = "raporhco"
            Me.pb1.visible = False
            CargaCombo
            txtcodigo(112).Text = Format(Now, "dd/mm/yyyy")
            
            
        Case 8 ' integracion en tesoreria
            ConexionConta vParamAplic.Seccionhorto
            FrameIntTesorQuaVisible True, h, w
            Tabla = "raporhco"
            Me.Pb4.visible = False
            
        Case 9 ' integracion en tesoresia del alta de socios de mogente
            ConexionConta vParamAplic.SeccionAlmaz
        
            FrameRegAltaSociosVisible True, h, w
            Tabla = "rsocios"
            Me.Pb6.visible = False
            
        Case 10 ' integracion en tesoresia del alta de socios de mogente
            ConexionConta vParamAplic.SeccionAlmaz
        
            FrameRegBajaSociosVisible True, h, w
            Tabla = "rsocios"
            Me.Pb7.visible = False
            
        Case 11 ' insercion de aportaciones para bolbaite
            FrameInsertarApoBolVisible True, h, w
            Tabla = "rfactsoc"
            Me.Pb8.visible = False
            Frame12.visible = False
            Frame12.Enabled = False
            
            CargarListView 0
            
        Case 12 ' Impresion de recibos de bolbaite
            FrameInsertarApoBolVisible True, h, w
            
            Label1(19).Caption = "Impresión de Recibos"
            Tabla = "raportacion"
            Me.Pb8.visible = False
            Frame5.visible = False
            Frame5.Enabled = False
            Me.CmdAcepInsApoBol.Top = 5100
            Me.CmdCancel(8).Top = 5100
            
            CargarListView 0
            
        Case 13 ' aportacion obligatoria de bolbaite
            FrameAportacionObligatoriaVisible True, h, w
            
            Tabla = "rsocios"
            Me.Pb9.visible = False
            
        Case 14
            FrameIntTesorBolVisible True, h, w
            
            ConexionConta vParamAplic.Seccionhorto
            Tabla = "raportacion"
            Me.Pb10.visible = False
            
        Case 15 ' certificado de aportacion bolbaite
            FrameCertificadoBolVisible True, h, w
            
            Tabla = "raportacion"
        
        Case 16 ' devolucion de aportaciones de bolbaite
            FrameDevolAporBolVisible True, h, w
            Tabla = "raportacion"
        
    End Select
    
    Me.Width = w + 70
    Me.Height = h + 350
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not vSeccion Is Nothing Then
        vSeccion.CerrarConta
        Set vSeccion = Nothing
    End If
End Sub

Private Sub frmApo_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de tipo de aportaciones
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtcodigo(indCodigo).Text = Format(txtcodigo(indCodigo).Text, "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtcodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmCla_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de clases
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Cuentas contables
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFPa_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de formas de pago de contabilidad
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtcodigo(indCodigo).Text = Format(txtcodigo(indCodigo).Text, "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " {raportacion.codaport} in (" & CadenaSeleccion & ")"
        Sql2 = " {raportacion.codaport} in [" & CadenaSeleccion & "]"
    Else
        Sql = " {raportacion.codaport} = -1 "
    End If
    If Not AnyadirAFormula(cadSelect, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub

Private Sub frmMens1_DatoSeleccionado(CadenaSeleccion As String)
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

Private Sub frmMens2_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " {raporhco.codvarie} in (" & CadenaSeleccion & ")"
        Sql2 = " {raporhco.codvarie} in [" & CadenaSeleccion & "]"
    Else
        Sql = " {raporhco.codvarie} = -1 "
    End If
    If Not AnyadirAFormula(cadSelect, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub

Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub Image1_Click(Index As Integer)
Dim I As Integer
Dim TotalCant As Currency
Dim TotalImporte As Currency

    Screen.MousePointer = vbHourglass
    
    Select Case Index
        ' informe de resultados y listado de retenciones
        Case 2
            For I = 1 To ListView1(0).ListItems.Count
                ListView1(0).ListItems(I).Checked = True
            Next I
        Case 3
            For I = 1 To ListView1(0).ListItems.Count
                ListView1(0).ListItems(I).Checked = False
            Next I
    End Select
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "Salta página por socio saca un informe para cada socio de las  " & vbCrLf & _
                      "aportaciones que se pasan al Arimoney.  " & vbCrLf & vbCrLf & _
                      "Es independiente del tipo de informe que se seleccione y no se " & vbCrLf & _
                      "imprime resumen. " & vbCrLf
                      
            
            vCadena = vCadena & "" & vbCrLf & vbCrLf
                      
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"

End Sub

Private Sub imgFec_Click(Index As Integer)
Dim indice As Integer

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
    
    Select Case Index
        Case 0, 1
            indice = Index + 8
        Case 6
            indice = Index + 6
        Case 8, 9
            indice = Index + 13
        Case 7
            indice = 20
        Case 10
            indice = 35
        Case 11
            indice = 41
        Case 14, 15
            indice = Index + 32
        Case 12
            indice = Index + 22
        Case 13
            indice = 49
        Case 16
            indice = 51
        Case 18
            indice = 54
        Case 19
            indice = 70
        Case 17
            indice = 57
        Case 20, 21
            indice = Index + 44
        Case 22
            indice = 74
        Case 23, 24
            indice = Index + 56
        Case 26
            indice = 86
        Case 25
            indice = 90
        Case 27
            indice = 91
        Case 28
            indice = 76
        Case 30, 31
            indice = Index + 71
        Case 29
            indice = 100
            
        ' devolucion de aportaciones quatretonda
        Case 32
            indice = 106
        Case 33
            indice = 111
        Case 34
            indice = 112
            
            
        Case Else
            indice = Index
    End Select
    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
    imgFec(0).Tag = indice 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtcodigo(indice).Text <> "" Then frmC.NovaData = txtcodigo(indice).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtcodigo(CByte(imgFec(0).Tag))
    ' ***************************
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 'Socios
            AbrirFrmSocios (Index)
        
        Case 2, 3  'Socios
            AbrirFrmSocios (Index)
            
        Case 4 ' formas de pago positiva
            AbrirFrmForpaConta (Index + 12)
        
        Case 9
            AbrirFrmForpaConta (Index + 8)
        
        Case 5 ' cuenta de banco prevista
            AbrirFrmCuentas (Index + 13)
    
        Case 6, 7  'Socios
            AbrirFrmSocios (Index + 17)
    
        Case 8 ' tipo de aportacion
            AbrirFrmTipoAportacion (Index + 5)
        Case 10 ' tipo de aportacion
            AbrirFrmTipoAportacion (Index + 9)
        
        'calculo de aportaciones para quatretonda
        Case 15 ' seccion
            AbrirFrmSeccion (32)
        Case 11, 12 ' socios
            AbrirFrmSocios (Index + 18)
        Case 13, 14 'clases
            AbrirFrmClase (Index + 14)
            
        ' informe de aportaciones para Quatretonda
        Case 16 'socio desde
            AbrirFrmSocios (Index + 20)
        Case 19 ' socio hasta
            AbrirFrmSocios (Index + 18)
        Case 17, 18 'clase
            AbrirFrmClase (Index + 21)
        
        ' integracion en tesoreria de quatretonda
        Case 23, 24 'socio desde hasta
            AbrirFrmSocios (Index + 21)
        Case 25 'clase
            AbrirFrmClase (Index + 18)
        Case 26 'clase
            AbrirFrmClase (Index + 22)
        Case 21 ' forma de pago
             AbrirFrmForpaConta (40)
        Case 22 ' forma de pago
             AbrirFrmForpaConta (42)
        Case 20 ' cta de banco prevista
            AbrirFrmCuentas (Index + 13)
        
        '[Monica]15/09/2014
        ' devolucion de aportaciones de quatretonda
        Case 56, 57  'Socios
            AbrirFrmSocios (Index + 51)
        Case 54, 55 'clase
            AbrirFrmClase (Index + 55)
                
        
        ' integracion tesoreria alta de socios mogente
        Case 28, 29 ' formas de pago positiva y negativa
            AbrirFrmForpaConta (Index + 24)
        Case 27 ' cta de banco prevista
            AbrirFrmCuentas (Index + 23)
        
        ' integracion en tesoreria baja de socios de mogente
        Case 30, 31 ' formas de pago positiva y negativa
            AbrirFrmForpaConta (Index + 25)
        Case 32 ' cta de banco prevista
            AbrirFrmCuentas (Index + 26)
        Case 33 ' socios
            AbrirFrmSocios (Index + 26)
        
        ' insercion de aportaciones de bolbaite
        Case 34, 35
            AbrirFrmSocios (Index + 32)
        Case 36
            AbrirFrmTipoAportacion (Index + 32)
        
        'obligatorias
        Case 38, 39
            AbrirFrmSocios (Index + 39)
        Case 37
            AbrirFrmTipoAportacion (Index + 34)
        
        'integracion tesoreria
        Case 42, 43
            AbrirFrmSocios (Index + 39)
        Case 40
            AbrirFrmTipoAportacion (Index + 35)
        Case 44, 45 ' formas de pago positiva y negativa
            AbrirFrmForpaConta (Index + 40)
        Case 41 ' cta de banco prevista
            AbrirFrmCuentas (Index + 42)
            
        'certificado de aportaciones
        Case 48, 49
            AbrirFrmSocios (Index + 40)
        Case 47
            AbrirFrmTipoAportacion (Index + 40)
        'devolucion de aportaciones bolbaite
        Case 46
            AbrirFrmSocios (Index + 57)
        Case 50
            AbrirFrmSocios (Index + 54)
        Case 51
            AbrirFrmTipoAportacion (Index + 54)
        Case 52
            AbrirFrmTipoAportacion (Index + 44)
                
    End Select
    
    PonerFoco txtcodigo(indCodigo)
End Sub

Private Sub AbrirFrmCuentas(indice As Integer)
    indCodigo = indice
    Set frmCtas = New frmCtasConta
    frmCtas.DatosADevolverBusqueda = "0|1|"
    frmCtas.CodigoActual = txtcodigo(indCodigo)
    frmCtas.Show vbModal
    Set frmCtas = Nothing
End Sub

Private Sub AbrirFrmForpaConta(indice As Integer)
    indCodigo = indice
    Set frmFpa = New frmForpaConta
    frmFpa.DatosADevolverBusqueda = "0|1|"
    frmFpa.CodigoActual = txtcodigo(indCodigo)
    frmFpa.Show vbModal
    Set frmFpa = Nothing
End Sub



Private Sub Opcion1_Click(Index As Integer)
    Check1.Enabled = Opcion1(0).Value
    If Not Check1.Enabled Then Check1.Value = 0
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtcodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
Dim cerrar As Boolean
'14/02/2007 antes
'    KEYpress KeyAscii
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'socio desde
            Case 1: KEYBusqueda KeyAscii, 1 'socio hasta
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
        
            Case 10: KEYBusqueda KeyAscii, 2 'socio desde
            Case 11: KEYBusqueda KeyAscii, 3 'socio hasta
            
            Case 8: KEYFecha KeyAscii, 0 'fecha desde
            Case 9: KEYFecha KeyAscii, 1 'fecha hasta
            
            Case 14: KEYFecha KeyAscii, 4 'fecha regularizacion
            Case 15: KEYFecha KeyAscii, 5 'fecha vto
        
            Case 16: KEYBusqueda KeyAscii, 4 'forma de pago positivas
            Case 17: KEYBusqueda KeyAscii, 9 'forma de pago negativas
        
            Case 18: KEYBusqueda KeyAscii, 5 'cta banco
            
            Case 12: KEYFecha KeyAscii, 6 'fecha de certificado
            
            Case 23: KEYBusqueda KeyAscii, 6 'socio desde
            Case 24: KEYBusqueda KeyAscii, 7 'socio hasta
            Case 21: KEYFecha KeyAscii, 8 'fecha desde
            Case 22: KEYFecha KeyAscii, 9 'fecha hasta
            Case 13: KEYBusqueda KeyAscii, 8 'tipo aportacion desde
            Case 19: KEYBusqueda KeyAscii, 10 'tipo aportacion hasta
            ' calculo de aportaciones de quatretonda
            Case 29: KEYBusqueda KeyAscii, 11 'socio desde
            Case 30: KEYBusqueda KeyAscii, 12 'socio hasta
            Case 27: KEYBusqueda KeyAscii, 13 'variedad desde
            Case 28: KEYBusqueda KeyAscii, 14 'variedad hasta
            Case 20: KEYFecha KeyAscii, 7 'fecha aportacion
            ' Listado de aportaciones para quatretonda
            Case 36: KEYBusqueda KeyAscii, 16 'socio desde
            Case 37: KEYBusqueda KeyAscii, 19 'socio hasta
            Case 38: KEYBusqueda KeyAscii, 17 'clase desde
            Case 39: KEYBusqueda KeyAscii, 18 'clase hasta
            Case 35: KEYFecha KeyAscii, 10 'fecha aportacion desde
            Case 41: KEYFecha KeyAscii, 11 'fecha aportacion hasta
            ' Integracion a tesoreria de aportaciones de quatretonda
            Case 44: KEYBusqueda KeyAscii, 23 'socio desde
            Case 45: KEYBusqueda KeyAscii, 24 'socio hasta
            Case 43: KEYBusqueda KeyAscii, 25 'clase desde
            Case 48: KEYBusqueda KeyAscii, 26 'clase hasta
            Case 46: KEYFecha KeyAscii, 14 'fecha aportacion desde
            Case 47: KEYFecha KeyAscii, 15 'fecha aportacion hasta
            
            'devolucion de aportaciones para quatretonda
            Case 107: KEYBusqueda KeyAscii, 56 'socio desde
            Case 108: KEYBusqueda KeyAscii, 57 'socio hasta
            Case 109: KEYBusqueda KeyAscii, 54 'clase desde
            Case 110: KEYBusqueda KeyAscii, 55 'clase hasta
            Case 106: KEYFecha KeyAscii, 32 'fecha aportacion desde
            Case 111: KEYFecha KeyAscii, 33 'fecha aportacion hasta
            Case 112: KEYFecha KeyAscii, 34 'fecha devolucion
            
            Case 34: KEYFecha KeyAscii, 12 'fecha de vencimiento
            Case 40: KEYBusqueda KeyAscii, 21 'f.pago positiva
            Case 42: KEYBusqueda KeyAscii, 22 'f.pago negativa
            
            Case 33: KEYBusqueda KeyAscii, 20 'cta banco prevista
            Case 49: KEYFecha KeyAscii, 13 'fecha de aportacion
            ' borrado masivo de apotaciones de quatretonda
            
            ' alta de socios de mogente
            Case 51: KEYFecha KeyAscii, 16 'fecha vto
            Case 52: KEYBusqueda KeyAscii, 28 'f.pago positiva
            Case 53: KEYBusqueda KeyAscii, 29 'f.pago negativa
            Case 50: KEYBusqueda KeyAscii, 27 'cta banco prevista
            
            ' baja de socios de mogente
            Case 54: KEYFecha KeyAscii, 18 'fecha devolucion
            Case 57: KEYFecha KeyAscii, 17 'fecha vto
            Case 56: KEYBusqueda KeyAscii, 31 'f.pago positiva
            Case 55: KEYBusqueda KeyAscii, 30 'f.pago negativa
            Case 58: KEYBusqueda KeyAscii, 32 'cta banco prevista
            Case 59: KEYBusqueda KeyAscii, 33 'codigo de socio
        
            ' insercion de aportaciones de bolbaite e impresion de recibos
            Case 70: KEYFecha KeyAscii, 19 'fecha recibo
            Case 64: KEYFecha KeyAscii, 20 'fecha desde
            Case 65: KEYFecha KeyAscii, 21 'fecha hasta
            Case 66: KEYBusqueda KeyAscii, 34 'socio desde
            Case 67: KEYBusqueda KeyAscii, 35 'socio hasta
            
            Case 68: KEYBusqueda KeyAscii, 36 'tipo de aportacion
        
            ' aportacion obligatoria de bolbaite
            Case 74: KEYFecha KeyAscii, 22 'fecha aportacion
            Case 77: KEYBusqueda KeyAscii, 38 'socio desde
            Case 78: KEYBusqueda KeyAscii, 39 'socio hasta
            
            Case 71: KEYBusqueda KeyAscii, 37 'tipo de aportacion
        
            ' integracion contable tesorieria de bolbaite
            Case 81: KEYBusqueda KeyAscii, 42 'socio desde
            Case 82: KEYBusqueda KeyAscii, 43 'socio hasta
            Case 79: KEYFecha KeyAscii, 23 'fecha desde
            Case 80: KEYFecha KeyAscii, 24 'fecha hasta
            Case 71: KEYBusqueda KeyAscii, 40 'tipo de aportacion
            Case 86: KEYFecha KeyAscii, 26 'fecha vto
            Case 85: KEYBusqueda KeyAscii, 45 'f.pago positiva
            Case 84: KEYBusqueda KeyAscii, 44 'f.pago negativa
            Case 83: KEYBusqueda KeyAscii, 41 'cta banco prevista
        
            ' certificado de aportacion de bolbaite
            Case 88: KEYBusqueda KeyAscii, 48 'socio desde
            Case 89: KEYBusqueda KeyAscii, 49 'socio hasta
            Case 90: KEYFecha KeyAscii, 25 'fecha desde
            Case 91: KEYFecha KeyAscii, 27 'fecha hasta
            Case 87: KEYBusqueda KeyAscii, 47 'tipo de aportacion
            Case 76: KEYFecha KeyAscii, 28 'fecha vto
        
            'devolucion de aportaciones
            Case 103: KEYBusqueda KeyAscii, 46 'socio desde
            Case 104: KEYBusqueda KeyAscii, 50 'socio hasta
            Case 101: KEYFecha KeyAscii, 30 'fecha desde
            Case 102: KEYFecha KeyAscii, 31 'fecha hasta
            Case 105: KEYBusqueda KeyAscii, 51 'tipo de aportacion
            Case 96: KEYBusqueda KeyAscii, 52 'tipo de aportacion
            Case 100: KEYFecha KeyAscii, 29 'fecha devolucion
            
        End Select
    Else
        KEYpressGnral KeyAscii, 0, cerrar
        If cerrar Then Unload Me
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
Dim cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 0, 1, 10, 11, 23, 24, 29, 30, 36, 37, 44, 45, 59, 66, 67, 77, 78, 81, 82, 88, 89, 103, 104, 107, 108 'socios
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
        
        Case 8, 9, 12, 14, 15, 20, 35, 41, 46, 47, 34, 49, 51, 54, 57, 64, 65, 74, 86, 79, 80, 100 'FECHAS
            If txtcodigo(Index).Text <> "" Then PonerFormatoFecha txtcodigo(Index), True
            
        Case 2, 3, 21, 22, 76, 90, 91, 101, 102, 106, 111, 112
            If txtcodigo(Index).Text <> "" Then PonerFormatoFecha txtcodigo(Index), False
            
        Case 6, 7, 60 'precios
            PonerFormatoDecimal txtcodigo(Index), 7
            
        Case 16, 17, 40, 42, 52, 53, 55, 56, 84, 85 ' forma de pago
            If vSeccion Is Nothing Then Exit Sub
            
            If txtcodigo(Index).Text <> "" Then txtNombre(Index).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtcodigo(Index).Text, "N")
            If txtNombre(Index).Text = "" Then
                MsgBox "Forma de Pago  no existe en la contabilidad. Reintroduzca.", vbExclamation
            End If
        
        Case 18, 33, 50, 58, 83 ' cta de banco
            If vSeccion Is Nothing Then Exit Sub
        
            If txtcodigo(Index).Text <> "" Then txtNombre(Index).Text = PonerNombreCuenta(txtcodigo(Index), 2)
            
        Case 4, 5 ' importes
            PonerFormatoDecimal txtcodigo(Index), 7
            
        Case 13, 19, 68, 71, 75, 87, 96, 105 ' codigo de aportaciones
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rtipoapor", "nomaport", "codaport", "N")
        
        Case 27, 28, 38, 39, 43, 48, 109, 110 'CLASES
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "clases", "nomclase", "codclase", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000")
        
        Case 32, 33 'SECCIONES
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rseccion", "nomsecci", "codsecci", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000")
    
        Case 25 'Año
            PonerFormatoEntero txtcodigo(Index)
        
        Case 26 ' Euros/hanegada para el calculo de aportaciones quatetonda
            PonerFormatoDecimal txtcodigo(Index), 3
        
        Case 31 'Ejercicio
            PonerFormatoEntero txtcodigo(Index)
        
        Case 69 'porcentaje de aportacion
            PonerFormatoDecimal txtcodigo(Index), 4
            
        Case 61, 62 'numero de factura
            PonerFormatoEntero txtcodigo(Index)
            
        Case 73 ' importe de la aportacion obligatoria
            PonerFormatoDecimal txtcodigo(Index), 3
            
        Case 92, 93, 94
            txtcodigo(Index).Text = UCase(txtcodigo(Index))
        
    End Select
End Sub


Private Sub FrameCalculoAporQuaVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCalculoAporQua.visible = visible
    If visible = True Then
        Me.FrameCalculoAporQua.Top = -90
        Me.FrameCalculoAporQua.Left = 0
        Me.FrameCalculoAporQua.Height = 7140
        Me.FrameCalculoAporQua.Width = 6555
        w = Me.FrameCalculoAporQua.Width
        h = Me.FrameCalculoAporQua.Height
    End If
End Sub


Private Sub FrameCobrosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 5790
        Me.FrameCobros.Width = 6555
        w = Me.FrameCobros.Width
        h = Me.FrameCobros.Height
    End If
End Sub


Private Sub FrameInformesVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameInforme.visible = visible
    If visible = True Then
        Me.FrameInforme.Top = -90
        Me.FrameInforme.Left = 0
        Me.FrameInforme.Height = 6300
        Me.FrameInforme.Width = 6555
        w = Me.FrameInforme.Width
        h = Me.FrameInforme.Height
    End If
End Sub

Private Sub FrameListAporQuaVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameListAporQua.visible = visible
    If visible = True Then
        Me.FrameListAporQua.Top = -90
        Me.FrameListAporQua.Left = 0
        Me.FrameListAporQua.Height = 6660
        Me.FrameListAporQua.Width = 6555
        w = Me.FrameListAporQua.Width
        h = Me.FrameListAporQua.Height
    End If
End Sub

Private Sub FrameDevolAporQuaVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameDevolAporQua.visible = visible
    If visible = True Then
        Me.FrameDevolAporQua.Top = -90
        Me.FrameDevolAporQua.Left = 0
        Me.FrameDevolAporQua.Height = 7140
        Me.FrameDevolAporQua.Width = 6555
        w = Me.FrameDevolAporQua.Width
        h = Me.FrameDevolAporQua.Height
    End If
End Sub


Private Sub FrameIntTesorQuaVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameIntTesorQua.visible = visible
    If visible = True Then
        Me.FrameIntTesorQua.Top = -90
        Me.FrameIntTesorQua.Left = 0
        Me.FrameIntTesorQua.Height = 7530
        Me.FrameIntTesorQua.Width = 6555
        w = Me.FrameIntTesorQua.Width
        h = Me.FrameIntTesorQua.Height
    End If
End Sub

Private Sub FrameRegularizacionVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameRegularizacion.visible = visible
    If visible = True Then
        Me.FrameRegularizacion.Top = -90
        Me.FrameRegularizacion.Left = 0
        Me.FrameRegularizacion.Height = 7530
        Me.FrameRegularizacion.Width = 6555
        w = Me.FrameRegularizacion.Width
        h = Me.FrameRegularizacion.Height
    End If
End Sub

Private Sub FrameInsertarApoBolVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameInsertarApoBol.visible = visible
    If visible = True Then
        Me.FrameInsertarApoBol.Top = -90
        Me.FrameInsertarApoBol.Left = 0
        Me.FrameInsertarApoBol.Height = 7530
        
        If OpcionListado = 12 Then Me.FrameInsertarApoBol.Height = 6460
        
        Me.FrameInsertarApoBol.Width = 6555
        w = Me.FrameInsertarApoBol.Width
        h = Me.FrameInsertarApoBol.Height
    End If
End Sub


Private Sub FrameAportacionObligatoriaVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameAporObligatoria.visible = visible
    If visible = True Then
        Me.FrameAporObligatoria.Top = -90
        Me.FrameAporObligatoria.Left = 0
        Me.FrameAporObligatoria.Height = 6330
        Me.FrameAporObligatoria.Width = 6555
        w = Me.FrameAporObligatoria.Width
        h = Me.FrameAporObligatoria.Height
    End If
End Sub

Private Sub FrameIntTesorBolVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameIntTesorBol.visible = visible
    If visible = True Then
        Me.FrameIntTesorBol.Top = -90
        Me.FrameIntTesorBol.Left = 0
        Me.FrameIntTesorBol.Height = 7530
        Me.FrameIntTesorBol.Width = 6555
        w = Me.FrameIntTesorBol.Width
        h = Me.FrameIntTesorBol.Height
    End If
End Sub

Private Sub FrameCertificadoBolVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameCertificadoBol.visible = visible
    If visible = True Then
        Me.FrameCertificadoBol.Top = -90
        Me.FrameCertificadoBol.Left = 0
        Me.FrameCertificadoBol.Height = 7530
        Me.FrameCertificadoBol.Width = 6555
        w = Me.FrameCertificadoBol.Width
        h = Me.FrameCertificadoBol.Height
    End If
End Sub

Private Sub FrameDevolAporBolVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameDevolAporBol.visible = visible
    If visible = True Then
        Me.FrameDevolAporBol.Top = -90
        Me.FrameDevolAporBol.Left = 0
        Me.FrameDevolAporBol.Height = 6900
        Me.FrameDevolAporBol.Width = 6555
        w = Me.FrameDevolAporBol.Width
        h = Me.FrameDevolAporBol.Height
    End If
End Sub





Private Sub FrameRegAltaSociosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameRegAltaSocios.visible = visible
    If visible = True Then
        Me.FrameRegAltaSocios.Top = -90
        Me.FrameRegAltaSocios.Left = 0
        Me.FrameRegAltaSocios.Height = 5400
        Me.FrameRegAltaSocios.Width = 6555
        w = Me.FrameRegAltaSocios.Width
        h = Me.FrameRegAltaSocios.Height
    End If
End Sub


Private Sub FrameRegBajaSociosVisible(visible As Boolean, ByRef h As Integer, ByRef w As Integer)
    Me.FrameRegBajaSocios.visible = visible
    If visible = True Then
        Me.FrameRegBajaSocios.Top = -90
        Me.FrameRegBajaSocios.Left = 0
        Me.FrameRegBajaSocios.Height = 5400
        Me.FrameRegBajaSocios.Width = 6555
        w = Me.FrameRegBajaSocios.Width
        h = Me.FrameRegBajaSocios.Height
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
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .ConSubInforme = True
        .EnvioEMail = False
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmSeccion(indice As Integer)
    indCodigo = indice
    Set frmSec = New frmManSeccion
    frmSec.DatosADevolverBusqueda = "0|1|"
    frmSec.Show vbModal
    Set frmSec = Nothing
End Sub

Private Sub AbrirFrmClase(indice As Integer)
    indCodigo = indice
    Set frmCla = New frmComercial
    AyudaClasesCom frmCla, txtcodigo(indice).Text
    Set frmCla = Nothing
End Sub

Private Sub AbrirFrmSocios(indice As Integer)
    indCodigo = indice
    Set frmSoc = New frmManSocios
    frmSoc.DatosADevolverBusqueda = "0|1|"
    frmSoc.Show vbModal
    Set frmSoc = Nothing
End Sub

Private Sub AbrirFrmTipoAportacion(indice As Integer)
    indCodigo = indice
    Set frmApo = New frmAPOTipos
    frmApo.DatosADevolverBusqueda = "0|1|"
    frmApo.Show vbModal
    Set frmApo = Nothing
End Sub

Private Sub AbrirFrmVariedades(indice As Integer)
    indCodigo = indice
    Set frmVar = New frmComVar
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.CodigoActual = txtcodigo(indCodigo)
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub

Private Function HayRegistros(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
Dim RS As ADODB.Recordset

    Sql = "Select * FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If RS.EOF Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        HayRegistros = False
    Else
        HayRegistros = True
    End If

End Function

Private Function ProcesarCambios(cadWhere As String) As Boolean
Dim Sql As String
Dim Sql1 As String
Dim I As Integer
Dim HayReg As Integer
Dim b As Boolean

On Error GoTo eProcesarCambios

    HayReg = 0
    
    conn.Execute "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
        
    If cadWhere <> "" Then
        cadWhere = QuitarCaracterACadena(cadWhere, "{")
        cadWhere = QuitarCaracterACadena(cadWhere, "}")
        cadWhere = QuitarCaracterACadena(cadWhere, "_1")
    End If
        
    Sql = "insert into tmpinformes (codusu, codigo1) select " & DBSet(vUsu.Codigo, "N")
    Sql = Sql & ", albaran.numalbar from albaran, albaran_variedad where albaran.numalbar not in (select numalbar from tcafpa) "
    Sql = Sql & " and albaran.numalbar = albaran_variedad.numalbar "
    
    If cadWhere <> "" Then Sql = Sql & " and " & cadWhere
    
    
    conn.Execute Sql
        
    ProcesarCambios = HayRegistros("tmpinformes", "codusu = " & vUsu.Codigo)

eProcesarCambios:
    If Err.Number <> 0 Then
        ProcesarCambios = False
    End If
End Function

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim vDevuelve As String
Dim Sql As String


    b = True

    Select Case OpcionListado
        Case 1
            If txtcodigo(6).Text = "" Then
                MsgBox "Debe introducir un valor en Precio Aumento de Kilos. Revise.", vbExclamation
                b = False
            End If
            If b Then
                If txtcodigo(7).Text = "" Then
                    MsgBox "Debe introducir un valor en Precio Disminución de Kilos. Revise.", vbExclamation
                    b = False
                End If
            End If
        Case 2
            If txtcodigo(4).Text = "" Then
                MsgBox "Debe introducir un valor en Precio Aumento de Kilos. Revise.", vbExclamation
                b = False
            End If
            If b Then
                If txtcodigo(5).Text = "" Then
                    MsgBox "Debe introducir un valor en Precio Disminución de Kilos. Revise.", vbExclamation
                    b = False
                End If
            End If
            If b Then
                If txtcodigo(8).Text = "" Or txtcodigo(9).Text = "" Then
                    MsgBox "Debe introducir valor en desde/hasta Fecha Factura. Revise.", vbExclamation
                    b = False
                End If
            End If
        Case 5 ' calculo de aportaciones de quatretonda
            If txtcodigo(32).Text = "" Then
                MsgBox "Debe introducir una sección. Reintroduzca.", vbExclamation
                PonerFoco txtcodigo(32)
                b = False
            End If
            ' debe introducir todos los datos para el calculo de aportaciones
            ' importe por hda
            If b Then
                If CDbl(ComprobarCero(txtcodigo(26).Text)) = "0" Then
                    MsgBox "Debe introducir el importe por hanegada. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(26)
                    b = False
                End If
            End If
            ' fecha de aportacion
            If b Then
                If txtcodigo(20).Text = "" Then
                    MsgBox "Debe introducir la Fecha de Aportación. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(20)
                    b = False
                End If
            End If
            ' año
            If b Then
                If txtcodigo(25).Text = "" Then
                    MsgBox "Debe introducir el Año. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(25)
                    b = False
                End If
            End If
            ' Ejercicio
            If b Then
                If txtcodigo(31).Text = "" Then
                    MsgBox "Debe introducir el Ejercicio. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(31)
                    b = False
                End If
            End If
            
        '[Monica]15/09/2014
        Case 17 ' Devoluciones de aportaciones de quatretonda
            ' debe introducir todos los datos para el calculo de aportaciones
            ' fecha de aportacion
            If b Then
                If txtcodigo(112).Text = "" Then
                    MsgBox "Debe introducir la Fecha de Devolución de Aportación. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(112)
                    b = False
                End If
            End If
            ' Ejercicio
            If b Then
                If txtcodigo(98).Text = "" Then
                    MsgBox "Debe introducir el Ejercicio de devolución. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(98)
                    b = False
                End If
            End If
            
            
        Case 8 ' Integracion de aportaciones en tesoreria
            If txtcodigo(34).Text = "" Then
                MsgBox "Debe introducir la Fecha de Vencimiento. Reintroduzca.", vbExclamation
                PonerFoco txtcodigo(34)
                b = False
            End If
            
            If b Then
                If txtcodigo(33).Text = "" Then
                    MsgBox "Debe introducir la Cuenta de Banco Prevista. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(33)
                    b = False
                Else
                    If PonerNombreCuenta(txtcodigo(33), 2) = "" Then
'                        MsgBox "La Cuenta de Banco Prevista no existe en Contabilidad. Reintroduzca.", vbExclamation
                        PonerFoco txtcodigo(33)
                        b = False
                    End If
                End If
            End If
            
            If b Then
                If txtcodigo(40).Text = "" Then
                    MsgBox "Debe introducir la Forma de Pago Positiva. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(40)
                    b = False
                Else
                    vDevuelve = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtcodigo(40).Text, "N")
                    If vDevuelve = "" Then
                        MsgBox "La Forma de Pago Positiva no existe en Contabilidad. Reintroduzca.", vbExclamation
                        PonerFoco txtcodigo(40)
                        b = False
                    End If
                End If
            End If
            
            If b Then
                If txtcodigo(42).Text = "" Then
                    MsgBox "Debe introducir la Forma de Pago Negativa. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(42)
                    b = False
                Else
                    vDevuelve = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtcodigo(42).Text, "N")
                    If vDevuelve = "" Then
                        MsgBox "La Forma de Pago Negativa no existe en Contabilidad. Reintroduzca.", vbExclamation
                        PonerFoco txtcodigo(42)
                        b = False
                    End If
                End If
            End If
            
        Case 9 ' integracion en tesoreria de alta de socios solo para mogente
            If txtcodigo(60).Text = "" Then
                MsgBox "Debe introducir el precio kilo. Reintroduzca.", vbExclamation
                PonerFoco txtcodigo(60)
                b = False
            End If
            
            If b Then
                If txtcodigo(51).Text = "" Then
                    MsgBox "Debe introducir la Fecha de Vencimiento. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(51)
                    b = False
                End If
            End If
            
            If b Then
                If txtcodigo(50).Text = "" Then
                    MsgBox "Debe introducir la Cuenta de Banco Prevista. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(50)
                    b = False
                Else
                    If PonerNombreCuenta(txtcodigo(50), 2) = "" Then
                        PonerFoco txtcodigo(50)
                        b = False
                    End If
                End If
            End If
            
            If b Then
                If txtcodigo(52).Text = "" Then
                    MsgBox "Debe introducir la Forma de Pago Positiva. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(52)
                    b = False
                Else
                    vDevuelve = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtcodigo(52).Text, "N")
                    If vDevuelve = "" Then
                        MsgBox "La Forma de Pago Positiva no existe en Contabilidad. Reintroduzca.", vbExclamation
                        PonerFoco txtcodigo(52)
                        b = False
                    End If
                End If
            End If
            
            If b Then
                If txtcodigo(53).Text = "" Then
                    MsgBox "Debe introducir la Forma de Pago Negativa. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(53)
                    b = False
                Else
                    vDevuelve = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtcodigo(53).Text, "N")
                    If vDevuelve = "" Then
                        MsgBox "La Forma de Pago Negativa no existe en Contabilidad. Reintroduzca.", vbExclamation
                        PonerFoco txtcodigo(53)
                        b = False
                    End If
                End If
            End If
        
        Case 10 ' integracion en tesoreria de baja de socios solo para mogente
            If txtcodigo(54).Text = "" Then
                MsgBox "Debe introducir la Fecha de Devolución. Reintroduzca.", vbExclamation
                PonerFoco txtcodigo(54)
                b = False
            End If
            
            If b Then
                If txtcodigo(57).Text = "" Then
                    MsgBox "Debe introducir la Fecha de Vencimiento. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(57)
                    b = False
                End If
            End If
            
            If b Then
                If txtcodigo(58).Text = "" Then
                    MsgBox "Debe introducir la Cuenta de Banco Prevista. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(58)
                    b = False
                Else
                    If PonerNombreCuenta(txtcodigo(58), 2) = "" Then
                        PonerFoco txtcodigo(58)
                        b = False
                    End If
                End If
            End If
            
            If b Then
                If txtcodigo(56).Text = "" Then
                    MsgBox "Debe introducir la Forma de Pago Positiva. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(56)
                    b = False
                Else
                    vDevuelve = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtcodigo(56).Text, "N")
                    If vDevuelve = "" Then
                        MsgBox "La Forma de Pago Positiva no existe en Contabilidad. Reintroduzca.", vbExclamation
                        PonerFoco txtcodigo(56)
                        b = False
                    End If
                End If
            End If
            
            If b Then
                If txtcodigo(55).Text = "" Then
                    MsgBox "Debe introducir la Forma de Pago Negativa. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(55)
                    b = False
                Else
                    vDevuelve = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtcodigo(55).Text, "N")
                    If vDevuelve = "" Then
                        MsgBox "La Forma de Pago Negativa no existe en Contabilidad. Reintroduzca.", vbExclamation
                        PonerFoco txtcodigo(55)
                        b = False
                    End If
                End If
            End If
        
            ' vemos si el socio al que vamos a dar de baja tiene concepto de aportacion 0
            If b Then
                Sql = "select * from raportacion where raportacion.codsocio = " & DBSet(txtcodigo(59).Text, "N")
                If TotalRegistrosConsulta(Sql) = 0 Then
                    MsgBox "El socio a dar de baja no tiene registro de regularizacion. Revise.", vbExclamation
                    PonerFoco txtcodigo(59)
                    b = False
                End If
                ' vemos si el socio tiene fecha de baja
                If b Then
                    Sql = "select * from rsocios  "
                    Sql = Sql & " where codsocio = " & DBSet(txtcodigo(59).Text, "N") & " and not fechabaja is null "
                    If TotalRegistrosConsulta(Sql) = 0 Then
                        MsgBox "El socio a dar de baja no tiene fecha de baja. Revise.", vbExclamation
                        PonerFoco txtcodigo(59)
                        b = False
                    End If
                End If
                ' vemos si el socio esta en la seccion de almazara
                If b Then
                    Sql = "select * from rsocios_seccion where codsocio = " & DBSet(txtcodigo(59).Text, "N")
                    Sql = Sql & " and codsecci = " & DBSet(vParamAplic.SeccionAlmaz, "N")
                    If TotalRegistrosConsulta(Sql) = 0 Then
                        MsgBox "El socio a dar de baja no es de la sección de almazara. Revise.", vbExclamation
                        PonerFoco txtcodigo(59)
                        b = False
                    End If
                End If
                ' comprobamos que a este socio no se le haya hecho ya la devolucion
                If b Then
                    Sql = "select sum(importe) from raportacion where codsocio = " & DBSet(txtcodigo(59).Text, "N")
                    Sql = Sql & " and fecaport >= (select max(fecaport) from raportacion where codsocio = " & DBSet(txtcodigo(59).Text, "N")
                    Sql = Sql & " and codaport = 0) "
                    If DevuelveValor(Sql) = 0 Then
                        MsgBox "Al socio ya se le ha hecho la devolución de la aportación. Revise.", vbExclamation
                        PonerFoco txtcodigo(59)
                        b = False
                    End If
                End If
            End If
        
        Case 11 ' insercion de aportaciones Bolbaite
            ' descripcion
            If txtcodigo(63).Text = "" Then
                MsgBox "Debe introducir la descripción. Revise.", vbExclamation
                PonerFoco txtcodigo(63)
                b = False
            End If
            ' tipo de aportacion
            If b Then
                If txtcodigo(68).Text = "" Then
                    MsgBox "Debe introducir el Tipo de Aportación. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(68)
                    b = False
                Else
                    vDevuelve = DevuelveDesdeBDNew(cAgro, "rtipoapor", "nomaport", "codaport", txtcodigo(68).Text, "N")
                    If vDevuelve = "" Then
                        MsgBox "El tipo de Aportación no existe. Reintroduzca.", vbExclamation
                        PonerFoco txtcodigo(68)
                        b = False
                    End If
                End If
            End If
        
        Case 12 ' Impresion de recibos de aportaciones de bolbaite
            If txtcodigo(70).Text = "" Then
                MsgBox "Debe introducir la fecha de Impresión de Recibo. Reintroduzca.", vbExclamation
                PonerFoco txtcodigo(70)
                b = False
            End If
        
        Case 13 ' Aportacion obligatoria de bolbaite
            If txtcodigo(74).Text = "" Then
                MsgBox "Debe introducir la fecha de Aportación. Reintroduzca.", vbExclamation
                PonerFoco txtcodigo(74)
                b = False
            End If
            ' descripcion
            If txtcodigo(72).Text = "" Then
                MsgBox "Debe introducir la descripción. Revise.", vbExclamation
                PonerFoco txtcodigo(72)
                b = False
            End If
            ' tipo de aportacion
            If b Then
                If txtcodigo(71).Text = "" Then
                    MsgBox "Debe introducir el Tipo de Aportación. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(71)
                    b = False
                Else
                    vDevuelve = DevuelveDesdeBDNew(cAgro, "rtipoapor", "nomaport", "codaport", txtcodigo(71).Text, "N")
                    If vDevuelve = "" Then
                        MsgBox "El tipo de Aportación no existe. Reintroduzca.", vbExclamation
                        PonerFoco txtcodigo(71)
                        b = False
                    End If
                End If
            End If
        
        Case 14 ' integracion en tesoreria de bolbaite
            If txtcodigo(86).Text = "" Then
                MsgBox "Debe introducir la fecha de Vencimiento. Reintroduzca.", vbExclamation
                PonerFoco txtcodigo(86)
                b = False
            End If
        
            ' tipo de aportacion
            If b Then
                If txtcodigo(75).Text = "" Then
                    MsgBox "Debe introducir el Tipo de Aportación. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(75)
                    b = False
                End If
            End If
        
            If b Then
                If txtcodigo(83).Text = "" Then
                    MsgBox "Debe introducir la Cuenta de Banco Prevista. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(83)
                    b = False
                Else
                    If PonerNombreCuenta(txtcodigo(83), 2) = "" Then
                        PonerFoco txtcodigo(83)
                        b = False
                    End If
                End If
            End If
            
            If b Then
                If txtcodigo(85).Text = "" Then
                    MsgBox "Debe introducir la Forma de Pago Positiva. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(85)
                    b = False
                Else
                    vDevuelve = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtcodigo(85).Text, "N")
                    If vDevuelve = "" Then
                        MsgBox "La Forma de Pago Positiva no existe en Contabilidad. Reintroduzca.", vbExclamation
                        PonerFoco txtcodigo(85)
                        b = False
                    End If
                End If
            End If
            
            If b Then
                If txtcodigo(84).Text = "" Then
                    MsgBox "Debe introducir la Forma de Pago Negativa. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(84)
                    b = False
                Else
                    vDevuelve = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtcodigo(84).Text, "N")
                    If vDevuelve = "" Then
                        MsgBox "La Forma de Pago Negativa no existe en Contabilidad. Reintroduzca.", vbExclamation
                        PonerFoco txtcodigo(84)
                        b = False
                    End If
                End If
            End If
        
        Case 15 ' Certificado
            If b Then
                If txtcodigo(90).Text = "" Then
                    MsgBox "Debe introducir la fecha desde de Aportación. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(90)
                    b = False
                End If
            End If
            If b Then
                If txtcodigo(91).Text = "" Then
                    MsgBox "Debe introducir la fecha hasta de Aportación. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(91)
                    b = False
                End If
            End If
            If b Then
                If txtcodigo(76).Text = "" Then
                    MsgBox "Debe introducir la fecha de Certificado. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(76)
                    b = False
                End If
            End If
            ' tipo de aportacion
            If b Then
                If txtcodigo(87).Text = "" Then
                    MsgBox "Debe introducir el Tipo de Aportación. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(87)
                    b = False
                Else
                    '[Monica]05/12/2013
                    If txtNombre(87).Text = "" Then
                        MsgBox "El Tipo de Aportación no existe. Reintroduzca.", vbExclamation
                        PonerFoco txtcodigo(87)
                        b = False
                    End If
                End If
            End If
                    
            ' Presidente
            If b Then
                If txtcodigo(92).Text = "" Then
                    MsgBox "Debe introducir el Presidente. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(92)
                    b = False
                End If
            End If
            ' Secretario
            If b Then
                If txtcodigo(93).Text = "" Then
                    MsgBox "Debe introducir el Secretario. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(93)
                    b = False
                End If
            End If
            ' Tesorero
            If b Then
                If txtcodigo(94).Text = "" Then
                    MsgBox "Debe introducir el Tesorero. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(94)
                    b = False
                End If
            End If
            
        Case 16 ' devolucion de aportacion
            If b Then
                If txtcodigo(100).Text = "" Then
                    MsgBox "Debe introducir la fecha de Devolución. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(100)
                    b = False
                End If
            End If
            ' tipo de aportacion origen
            If b Then
                If txtcodigo(105).Text = "" Then
                    MsgBox "Debe introducir el Tipo de Aportación. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(105)
                    b = False
                End If
            End If
            ' tipo de aportacion destino
            If b Then
                If txtcodigo(96).Text = "" Then
                    MsgBox "Debe introducir el Nuevo Tipo de Aportación. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(96)
                    b = False
                End If
            End If
            ' descripcion
            If b Then
                If txtcodigo(99).Text = "" Then
                    MsgBox "Debe introducir la descripción. Revise.", vbExclamation
                    PonerFoco txtcodigo(99)
                    b = False
                End If
            End If
    End Select
    
    DatosOk = b

End Function




'======================================================================
'GRABAR EN TESORERIA
'======================================================================
' ### [Monica] 16/01/2008
Private Function InsertarEnTesoreriaNewAPO(MenError As String, Socio As Long, Importe As Currency, FVenci As String, FPNeg As String, FPPos As String, CtaBanco As String, FecFac As String, Tipo As Byte) As Boolean
'Guarda datos de Tesoreria en tablas: conta.scobros
' Tipo: 0 = Regularizacion
'       1 = Alta Socio
'       2 = Baja Socio

Dim b As Boolean
Dim Sql As String, Text33csb As String, Text41csb As String
Dim Sql4 As String
Dim Rs4 As ADODB.Recordset

Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim I As Byte
Dim DigConta As String
Dim CC As String
Dim vrefer As String
Dim letraser As String
Dim Rsx As ADODB.Recordset
Dim vSocio As cSocio
Dim Seccion As Integer
Dim FecVen As String
Dim ForpaNeg As String
Dim ForpaPos As String
Dim CtaBan As String
Dim fecfactu As String
Dim numfactu As String


    On Error GoTo EInsertarTesoreriaNew

    b = False
    InsertarEnTesoreriaNewAPO = False
    CadValues = ""
    CadValues2 = ""
    
    Seccion = vParamAplic.SeccionAlmaz
    
    
    Set vSocio = New cSocio
    If vSocio.LeerDatos(CStr(Socio)) Then
        If vSocio.LeerDatosSeccion(CStr(Socio), CStr(Seccion)) Then
            
            FecVen = FVenci 'txtcodigo(15).Text
            ForpaNeg = FPNeg 'txtcodigo(17).Text
            ForpaPos = FPPos 'txtcodigo(16).Text
            CtaBan = CtaBanco 'txtcodigo(18).Text
            fecfactu = FecFac 'txtcodigo(14).Text
            numfactu = Format(vSocio.Codigo, "000000")
            
            
            If DBLet(Importe, "N") >= 0 Then
                ' si el importe de la regularizacion
                letraser = ""
                letraser = DevuelveDesdeBDNew(cAgro, "usuarios.stipom", "letraser", "codtipom", "RAP", "T")
    
                Select Case Tipo
                    Case 0 ' Regularizacion
                        Text33csb = "'Regularización Aportación de " & Format(DBLet(fecfactu, "F"), "dd/mm/yy") & "'"
                        Text41csb = "de " & DBSet(Importe, "N")
                    Case 1 ' Alta Socios
                        Text33csb = "'Aportación de Alta Socio de " & Format(DBLet(fecfactu, "F"), "dd/mm/yy") & "'"
                        Text41csb = "de " & DBSet(Importe, "N")
                    Case 2 ' Baja Socios
                        Text33csb = "'Aportación de Baja Socio de " & Format(DBLet(fecfactu, "F"), "dd/mm/yy") & "'"
                        Text41csb = "de " & DBSet(Importe, "N")
                End Select
                        
                CC = DBLet(vSocio.Digcontrol, "T")
                If DBLet(vSocio.Digcontrol, "T") = "**" Then CC = "00"
                
                '[Monica]03/07/2013: añado trim(codmacta)
                CadValuesAux2 = "(" & DBSet(letraser, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(fecfactu, "F") & ", 1," & DBSet(Trim(vSocio.CtaClien), "T") & ","
                CadValues2 = CadValuesAux2 & DBSet(ForpaPos, "N") & "," & DBSet(FecVen, "F") & "," & DBSet(Importe, "N") & ","
                CadValues2 = CadValues2 & DBSet(CtaBan, "T") & "," & DBSet(vSocio.Banco, "N", "S") & "," & DBSet(vSocio.Sucursal, "N", "S") & ","
                CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(vSocio.CuentaBan, "T", "S") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & ",1" ')"
                
                '[Monica]22/11/2013: Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    CadValues2 = CadValues2 & ", " & DBSet(vSocio.Iban, "T", "S") & ") "
                Else
                    CadValues2 = CadValues2 & ") "
                End If
                
    
                'Insertamos en la tabla scobro de la CONTA
                Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
                Sql = Sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
                Sql = Sql & " text33csb, text41csb, agente" ') "
                '[Monica]22/11/2013: Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    Sql = Sql & ", iban) "
                Else
                    Sql = Sql & ") "
                End If
                
                Sql = Sql & " VALUES " & CadValues2
                ConnConta.Execute Sql
            
            Else
                '********** si la factura es negativa se inserta en la spago con valor positivo
                CadValues2 = ""
            
                CadValuesAux2 = "('" & vSocio.CtaProv & "', " & DBSet(numfactu, "N") & ", '" & Format(fecfactu, FormatoFecha) & "', "
                '------------------------------------------------------------
                
                CC = DBLet(vSocio.Digcontrol, "T")
                If DBLet(vSocio.Digcontrol, "T") = "**" Then CC = "00"
                
                I = 1
                CadValues2 = CadValuesAux2 & I
                CadValues2 = CadValues2 & ", " & DBSet(ForpaNeg, "N") & ", " & DBSet(FecVen, "F") & ", "
                CadValues2 = CadValues2 & DBSet(DBLet(Importe, "N") * (-1), "N") & ", " & DBSet(CtaBan, "T") & ","
            
                'David. Para que ponga la cuenta bancaria (SI LA tiene)
                CadValues2 = CadValues2 & DBSet(vSocio.Banco, "T", "S") & "," & DBSet(vSocio.Sucursal, "T", "S") & ","
                CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(vSocio.CuentaBan, "T", "S") & ","
            
                'David. JUNIO 07.   Los dos textos de grabacion de datos de csb
                Select Case Tipo
                    Case 0
                        Sql = "Regularización de Aportación"
                    Case 1
                        Sql = "Aportación de Alta Socio"
                    Case 2
                        Sql = "Devolución Aportación Baja Socio"
                End Select
                    
                CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "',"
                
                Sql = " de " & Format(DBLet(fecfactu, "F"), "dd/mm/yyyy")
                CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "'" ')"

                '[Monica]22/11/2013: Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    CadValues2 = CadValues2 & ", " & DBSet(vSocio.Iban, "T", "S") & ") "
                Else
                    CadValues2 = CadValues2 & ") "
                End If
            
                'Grabar tabla spagop de la CONTABILIDAD
                '-------------------------------------------------
                If CadValues2 <> "" Then
                    'Insertamos en la tabla spagop de la CONTA
                    'David. Cuenta bancaria y descripcion textos
                    Sql = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,entidad,oficina,cc,cuentaba,text1csb,text2csb" ' ) "
                    '[Monica]22/11/2013: Tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                        Sql = Sql & ", iban) "
                    Else
                        Sql = Sql & ") "
                    End If
                    
                    Sql = Sql & " VALUES " & CadValues2
                    ConnConta.Execute Sql
                End If
                '*******
            End If
        End If
    End If

    b = True
    
EInsertarTesoreriaNew:
    If Err.Number <> 0 Then
        b = False
        MenError = MenError & " " & Err.Description
    End If
    InsertarEnTesoreriaNewAPO = b
End Function


Private Sub ConexionConta(Seccion As Integer)
    
    If vSeccion Is Nothing Then
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(CStr(Seccion)) Then
            vSeccion.AbrirConta
        End If
    Else
        ' si el objeto existia: cerramos la conexion y volvemos crearlo
        vSeccion.CerrarConta
        Set vSeccion = Nothing
        
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(CStr(Seccion)) Then
            vSeccion.AbrirConta
        End If
    End If
End Sub


Private Function ComprobarCtaContable_new(cadTabla As String, Opcion As Byte, Optional Seccion As Integer, Optional cuenta As String) As Boolean
'Comprobar que todas las ctas contables de los distintos clientes de las facturas
'que vamos a contabilizar existan en la contabilidad
Dim Sql As String
Dim RS As ADODB.Recordset
'Dim RSconta As ADODB.Recordset
Dim b As Boolean
Dim cadG As String
Dim SQLcuentas As String
Dim CadCampo1 As String
Dim numnivel As String
Dim NumDigit As String
Dim NumDigit3 As String


    On Error GoTo ECompCta

    ComprobarCtaContable_new = False

    Label1(1).Caption = "Comprobando Cuentas Contables "
    Label1(1).visible = True
    Me.Refresh
    DoEvents

    SQLcuentas = "SELECT count(*) FROM cuentas WHERE apudirec='S' "
    If cadG <> "" Then SQLcuentas = SQLcuentas & cadG
    Select Case Opcion
        Case 1
            'Seleccionamos los distintos socios, cuentas que vamos a facturar
            Sql = "SELECT DISTINCT tmpinformes.codigo1 codsocio, rsocios_seccion.codmaccli as codmacta "
            Sql = Sql & " FROM (tmpinformes INNER JOIN rsocios_seccion ON tmpinformes.codigo1=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(Seccion, "N")
            Sql = Sql & " and tmpinformes.codusu = " & vUsu.Codigo & ") "
            Sql = Sql & " ORDER BY 1 "
        
        Case 2
            'Seleccionamos los distintos socios proveedor, cuentas que vamos a facturar
            Sql = "SELECT DISTINCT tmpinformes.codigo1 codsocio, rsocios_seccion.codmacpro as codmacta "
            Sql = Sql & " FROM (tmpinformes INNER JOIN rsocios_seccion ON tmpinformes.codigo1=rsocios_seccion.codsocio and rsocios_seccion.codsecci = " & DBSet(Seccion, "N")
            Sql = Sql & " and tmpinformes.codusu = " & vUsu.Codigo & ") "
            Sql = Sql & " ORDER BY 1 "
        
        
        
    End Select
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    b = True

    While Not RS.EOF And b
        If Opcion < 4 Then
            Sql = SQLcuentas & " AND codmacta= " & DBSet(RS!Codmacta, "T")
        End If

        If Not (RegistrosAListar(Sql, cConta) > 0) Then
        'si no lo encuentra
            b = False 'no encontrado
            If Opcion = 1 Or Opcion = 2 Then
                Sql = DBLet(RS!Codmacta, "T") & " del Socio " & Format(RS!Codsocio, "000000")
            End If
        End If

        RS.MoveNext
    Wend

    If Not b Then
        Sql = "Comprobando Ctas Contables en contabilidad... " & vbCrLf & vbCrLf & Sql

        MsgBox Sql, vbExclamation
        ComprobarCtaContable_new = False
    Else
        ComprobarCtaContable_new = True
    End If
    
    Exit Function

ECompCta:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Ctas Contables", Err.Description
    End If
End Function


Private Function IntegracionAportacionesTesoreria(Tabla As String, vWhere As String)
Dim Sql As String
Dim RS As ADODB.Recordset
Dim Sql2 As String
Dim SqlValues As String
Dim Descripc As String
Dim Campanya As String
Dim ImporIni As Currency
Dim Importe As Currency
Dim b As Boolean
Dim MensError As String
Dim SqlExiste As String

    On Error GoTo eIntegracionAportacionesTesoreria
        
        
    Sql = "INTAPO" 'Integracion de aportaciones
    'Bloquear para que nadie mas pueda realizarlo
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se puede realizar el proceso de Integración de Aportaciones. Hay otro usuario realizándolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
        
    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    LOG.Insertar 3, vUsu, "Int.Tesoreria Aportaciones: " & vbCrLf & Tabla & vbCrLf & vWhere
    Set LOG = Nothing
    '-----------------------------------------------------------------------------
        
    conn.BeginTrans
    ConnConta.BeginTrans

    Sql = "select raporhco.codsocio, sum(impaport) as importe from " & Tabla
    If vWhere <> "" Then Sql = Sql & " WHERE " & vWhere
    Sql = Sql & " group by 1 "
    Sql = Sql & " order by 1 "
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    b = True

    Pb4.visible = True
    Pb4.Max = TotalRegistrosConsulta(Sql)
    Pb4.Value = 0
    
    While Not RS.EOF And b
        IncrementarProgresNew Pb4, 1
    
        MensError = "Insertando cobro en tesoreria" & vbCrLf & vbCrLf
        b = InsertarEnTesoreriaAPOQua(MensError, RS!Codsocio, DBLet(RS!Importe, "N"))
    
        RS.MoveNext
    Wend
    Set RS = Nothing
    
    If b Then
        MensError = "Actualizando Aportaciones" & vbCrLf & vbCrLf
        b = ActualizarAportaciones(MensError, Tabla, vWhere)
    End If
    
eIntegracionAportacionesTesoreria:
    If Err.Number <> 0 Or Not b Then
        IntegracionAportacionesTesoreria = False
        conn.RollbackTrans
        ConnConta.RollbackTrans
        MsgBox "Se ha producido un error " & MensError, vbExclamation
        
    Else
        IntegracionAportacionesTesoreria = True
        conn.CommitTrans
        ConnConta.CommitTrans
    End If
    
    DesBloqueoManual ("INTAPO") 'Integracion de aportaciones en tesoreria
    
    Screen.MousePointer = vbDefault
    
End Function

'======================================================================
'GRABAR EN TESORERIA
'======================================================================
' ### [Monica] 17/01/2012
Private Function InsertarEnTesoreriaAPOQua(MenError As String, Socio As Long, Importe As Currency) As Boolean
'Guarda datos de Tesoreria en tablas: conta.scobros
' Tipo: 0 = almazara
'       1 = bodega

Dim b As Boolean
Dim Sql As String, Text33csb As String, Text41csb As String
Dim Sql4 As String
Dim Rs4 As ADODB.Recordset

Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim I As Byte
Dim DigConta As String
Dim CC As String
Dim vrefer As String
Dim letraser As String
Dim Rsx As ADODB.Recordset
Dim vSocio As cSocio
Dim Seccion As Integer
Dim FecVen As String
Dim ForpaNeg As String
Dim ForpaPos As String
Dim CtaBan As String
Dim fecfactu As String
Dim numfactu As String
Dim Text1csb As String
Dim Text2csb As String


    On Error GoTo EInsertarTesoreriaNew

    b = False
    InsertarEnTesoreriaAPOQua = False
    CadValues = ""
    CadValues2 = ""
    
    Seccion = vParamAplic.Seccionhorto
    
    
    Set vSocio = New cSocio
    If vSocio.LeerDatos(CStr(Socio)) Then
        If vSocio.LeerDatosSeccion(CStr(Socio), CStr(Seccion)) Then
            FecVen = txtcodigo(34).Text
            ForpaNeg = txtcodigo(40).Text
            ForpaPos = txtcodigo(42).Text
            CtaBan = txtcodigo(33).Text
            fecfactu = txtcodigo(49).Text
            numfactu = Format(vSocio.Codigo, "000000")
            
            
            If DBLet(Importe, "N") >= 0 Then
                ' si el importe de la regularizacion
                letraser = ""
                letraser = DevuelveDesdeBDNew(cAgro, "usuarios.stipom", "letraser", "codtipom", "APO", "T")
    
                Text33csb = "'Cargo Aportaciones Coop. de " & Format(DBLet(fecfactu, "F"), "dd/mm/yy") & "'"
                Text41csb = "de " & DBSet(Importe, "N")
    
                CC = DBLet(vSocio.Digcontrol, "T")
                If DBLet(vSocio.Digcontrol, "T") = "**" Then CC = "00"
    
    
                CadValuesAux2 = "(" & DBSet(letraser, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(fecfactu, "F") & ", 1," & DBSet(vSocio.CtaClien, "T") & ","
                CadValues2 = CadValuesAux2 & DBSet(ForpaPos, "N") & "," & DBSet(FecVen, "F") & "," & DBSet(Importe, "N") & ","
                CadValues2 = CadValues2 & DBSet(CtaBan, "T") & "," & DBSet(vSocio.Banco, "N", "S") & "," & DBSet(vSocio.Sucursal, "N", "S") & ","
                CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(vSocio.CuentaBan, "T", "S") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & ",1" ')"
                '[Monica]22/11/2013: Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    CadValues2 = CadValues2 & ", " & DBSet(vSocio.Iban, "T", "S") & ") "
                Else
                    CadValues2 = CadValues2 & ") "
                End If
    
                'Insertamos en la tabla scobro de la CONTA
                Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
                Sql = Sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
                Sql = Sql & " text33csb, text41csb, agente" ') "
                '[Monica]22/11/2013: Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    Sql = Sql & ", iban) "
                Else
                    Sql = Sql & ") "
                End If
                Sql = Sql & " VALUES " & CadValues2
                ConnConta.Execute Sql
            
            Else
                '[Monica]01/09/2014: añadido esto, si el importe es negativo lo tengo que cambiar a positivo
                Importe = Importe * (-1)
            
            
                Text1csb = "'Devolución Aportaciones de " & Format(DBLet(fecfactu, "F"), "dd/mm/yyyy") & "'"
                Text2csb = "de " & DBSet(Importe, "N")
    
                CC = DBLet(vSocio.Digcontrol, "T")
                If DBLet(vSocio.Digcontrol, "T") = "**" Then CC = "00"
    
                CadValuesAux2 = "(" & DBSet(vSocio.CtaClien, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(fecfactu, "F") & ", 9,"
                CadValues2 = CadValuesAux2 & DBSet(ForpaNeg, "N") & "," & DBSet(FecVen, "F") & "," & DBSet(Importe, "N") & ","
                CadValues2 = CadValues2 & DBSet(CtaBan, "T") & "," & DBSet(vSocio.Banco, "N", "S") & "," & DBSet(vSocio.Sucursal, "N", "S") & ","
                CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(vSocio.CuentaBan, "T", "S") & ","
                CadValues2 = CadValues2 & Text1csb & "," & DBSet(Text2csb, "T") '& ")"
                '[Monica]22/11/2013: Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    CadValues2 = CadValues2 & ", " & DBSet(vSocio.Iban, "T", "S") & ") "
                Else
                    CadValues2 = CadValues2 & ") "
                End If
    
                'Insertamos en la tabla scobro de la CONTA
                Sql = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,entidad,oficina,cc,cuentaba,text1csb,text2csb" ') "
                '[Monica]22/11/2013: Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    Sql = Sql & ", iban) "
                Else
                    Sql = Sql & ") "
                End If
                
                Sql = Sql & " VALUES " & CadValues2
                
                ConnConta.Execute Sql
            
            End If
            
        End If
    End If

    b = True
    
EInsertarTesoreriaNew:
    If Err.Number <> 0 Then
        b = False
        MenError = Err.Description
    End If
    InsertarEnTesoreriaAPOQua = b
End Function

Private Function CargarTemporalQua(nTabla1 As String, nSelect1 As String) As Boolean
Dim RS As ADODB.Recordset
Dim Sql As String
Dim Sql2 As String
Dim cValues As String
Dim AcumAnt As Long
Dim Kilos As Long
Dim KilosMed As Long
Dim Nregs As Integer
Dim SocioAnt As Long
Dim NombreAnt As String
Dim Diferencia As Long
Dim Entro As Boolean
Dim Importe As Currency


    On Error GoTo eCargarTablaTemporal

    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    Sql = "insert into tmpinformes (codusu, importe1, codigo1) "
    
    Sql2 = " select " & vUsu.Codigo & ", raporhco.numaport, raporhco.codsocio "
    Sql2 = Sql2 & " from " & nTabla1
    
    If nSelect1 <> "" Then Sql2 = Sql2 & " where " & nSelect1
    
    conn.Execute Sql & Sql2

    CargarTemporalQua = True
    
    Exit Function
    
eCargarTablaTemporal:
    MuestraError Err.Number, "Cargar Tabla Temporal", Err.Description
End Function


Private Function ActualizarAportaciones(MensError As String, cTabla As String, cWhere As String) As Boolean
Dim Sql As String

    On Error GoTo eActualizarAportaciones

    ActualizarAportaciones = False

    Sql = "update raporhco, tmpinformes set intconta = 1 where tmpinformes.codusu = " & vUsu.Codigo
    Sql = Sql & " and tmpinformes.importe1 = raporhco.numaport "
    
    conn.Execute Sql

    ActualizarAportaciones = True
    Exit Function

eActualizarAportaciones:
    MensError = MensError & vbCrLf & Err.Description
End Function

Private Function ActualizarAportacionesBol(MensError As String, cTabla As String, cWhere As String) As Boolean
Dim Sql As String

    On Error GoTo eActualizarAportacionesBol

    ActualizarAportacionesBol = False

    Sql = "update raportacion, tmpinformes set intconta = 1 where tmpinformes.codusu = " & vUsu.Codigo
    Sql = Sql & " and tmpinformes.importe1 = raportacion.numfactu "
    Sql = Sql & " and tmpinformes.fecha1 = raportacion.fecaport "
    Sql = Sql & " and (tmpinformes.nombre1 = raportacion.codtipom or tmpinformes.nombre1 is null) "
    Sql = Sql & " and tmpinformes.codigo1 = raportacion.codsocio "
    
    conn.Execute Sql

    ActualizarAportacionesBol = True
    Exit Function

eActualizarAportacionesBol:
    MensError = MensError & vbCrLf & Err.Description
End Function



Private Sub CargaCombo()
        
    Combo1(0).Clear
    'tipo de Aportacion
    Combo1(0).AddItem "No Contabilizada"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Contabilizada"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Ambas"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2


    'tipo de relacion con la cooperativa
    Combo1(1).AddItem "Todos"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Socio"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    Combo1(1).AddItem "Asociado"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 2
    Combo1(1).AddItem "Tercero"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 3




End Sub

Private Function BorradoMasivoAporQua(Tabla As String, vWhere As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim cadWhere As String
Dim Mens As String
Dim Nregs As Long

    On Error GoTo eBorradoMasivoAporQua

    BorradoMasivoAporQua = False

    Sql = "select raporhco.* from " & Tabla
    If vWhere <> "" Then Sql = Sql & " where " & vWhere
   
    Sql2 = Sql
    If vWhere <> "" Then
        Sql2 = Sql2 & " and intconta = 1"
    Else
        Sql2 = Sql2 & " where raporhco.intconta = 1"
    End If
    
    If TotalRegistrosConsulta(Sql2) > 0 Then
        Mens = "Hay aportaciones pasadas a Tesoreria. Revise."
        MsgBox Mens, vbExclamation
        Exit Function
    End If
   
    Sql2 = "delete from raporhco "
    If vWhere <> "" Then
        cadWhere = cadWhere & " where " & vWhere & " and intconta = 0  "
    Else
        cadWhere = cadWhere & " and intconta = 0 "
    End If
    Nregs = TotalRegistrosConsulta("select raporhco.* from " & Tabla & cadWhere)
    
    If MsgBox("Va a eliminar " & Nregs & " registros no contabilizados. ¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Function
    
    conn.Execute Sql2 & cadWhere
   
    BorradoMasivoAporQua = True
    Exit Function
   
eBorradoMasivoAporQua:
    
End Function


Private Function CargarTablaTemporal2(nTabla1 As String, nSelect1 As String, Precio1 As String, ByRef pb1 As ProgressBar) As Boolean
Dim RS As ADODB.Recordset
Dim Sql As String
Dim Sql2 As String
Dim cValues As String
Dim AcumAnt As Long
Dim Kilos As Long
Dim KilosMed As Long
Dim Nregs As Integer
Dim SocioAnt As Long
Dim NombreAnt As String
Dim Diferencia As Long
Dim Entro As Boolean
Dim Importe As Currency


    On Error GoTo eCargarTablaTemporal

    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    Sql = "insert into tmpinformes (codusu, codigo1, nombre1, importe1)  "
    
    Sql2 = " select " & vUsu.Codigo & ", rsocios.codsocio, rsocios.nomsocio, sum(if(kilosnet is null, 0,kilosnet)) "
    Sql2 = Sql2 & " from " & nTabla1 & " left join rhisfruta on rsocios.codsocio = rhisfruta.codsocio "
    
    If nSelect1 <> "" Then Sql2 = Sql2 & " where  " & nSelect1
    Sql2 = Sql2 & " group by 1,2,3"
    Sql2 = Sql2 & " having  sum(if(kilosnet is null, 0,kilosnet)) <> 0 "
    Sql2 = Sql2 & " order by 1,2,3"
    
    conn.Execute Sql & Sql2
    
    CargarTablaTemporal2 = True
    
    Exit Function
    
eCargarTablaTemporal:
    MuestraError Err.Number, "Cargar Tabla Temporal", Err.Description
End Function




Private Function ActualizarRegularizacionAltaSocio(Precio As Currency)
Dim Sql As String
Dim RS As ADODB.Recordset
Dim Sql2 As String
Dim SqlValues As String
Dim Descripc As String
Dim Campanya As String
Dim ImporIni As Currency
Dim Importe As Currency
Dim b As Boolean
Dim MensError As String
Dim SqlExiste As String
Dim Fecha As Date

    On Error GoTo eActualizarRegularizacion
        
        
    Sql = "ALTAPO" 'regularizacion de aportaciones alta socios
    
    'Bloquear para que nadie mas pueda realizarlo
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se puede realizar el proceso de Regularización de Aportaciones de Alta Socios. Hay otro usuario realizándolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
        
    conn.BeginTrans
    ConnConta.BeginTrans

    Sql = "select * from tmpinformes where codusu = " & vUsu.Codigo
    Sql = Sql & " order by codigo1 "
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Sql2 = "insert into raportacion (codsocio,fecaport,codaport,descripcion,campanya,kilos,importe) values "

    Campanya = Mid(Format(Year(CDate(vParam.FecIniCam)), "0000"), 3, 2) & "/" & Mid(Format(Year(CDate(vParam.FecFinCam)), "0000"), 3, 2)
    Descripc = "ACUMULADA " & Campanya

    b = True

    Pb6.visible = True
    Pb6.Max = TotalRegistrosConsulta(Sql)
    Pb6.Value = 0
    
    Fecha = vParam.FecIniCam 'DateAdd("d", (-1), vParam.FecIniCam)
    
    While Not RS.EOF And b
        IncrementarProgresNew Pb6, 1
    
        SqlValues = ""
        
        Importe = Round2(DBLet(RS!importe1, "N") * Precio, 2)
    
        SqlExiste = "select count(*) from raportacion where codsocio = " & DBSet(RS!Codigo1, "N") & " and codaport=0 and fecaport=" & DBSet(Fecha, "F")
        b = (TotalRegistros(SqlExiste) = 0)
        
        If Not b Then
            MsgBox "El registro para el socio " & Format(DBLet(RS!Codigo1, "N"), "000000") & " de fecha " & DBSet(Fecha, "F") & " y tipo 0 existe. Revise.", vbExclamation
        Else
            SqlValues = "(" & DBSet(RS!Codigo1, "N") & "," & DBSet(Fecha, "F") & ",0," & DBSet(Descripc, "T") & ","
            SqlValues = SqlValues & DBSet(Campanya, "T") & "," & DBSet(RS!importe1, "N") & "," & DBSet(Importe, "N") & ")"
            
            conn.Execute Sql2 & SqlValues
            
            MensError = "Insertando cobro en tesoreria"
            b = InsertarEnTesoreriaNewAPO(MensError, RS!Codigo1, DBLet(Importe, "N"), txtcodigo(51).Text, txtcodigo(52).Text, txtcodigo(53).Text, txtcodigo(50).Text, CStr(Fecha), 1)
            If Not b Then
                MsgBox "Error: " & MensError, vbExclamation
            End If
            
        End If
    
    
        RS.MoveNext
    Wend
    Set RS = Nothing
    
eActualizarRegularizacion:
    If Err.Number <> 0 Or Not b Then
        ActualizarRegularizacionAltaSocio = False
        conn.RollbackTrans
        ConnConta.RollbackTrans
    Else
        ActualizarRegularizacionAltaSocio = True
        conn.CommitTrans
        ConnConta.CommitTrans
    End If
    
    DesBloqueoManual ("ALTAPO") 'regularizacion de aportaciones
    
    Screen.MousePointer = vbDefault
    
End Function



Private Function CargarTablaTemporal3(nTabla1 As String, nSelect1 As String, Precio1 As String, ByRef pb1 As ProgressBar) As Boolean
Dim RS As ADODB.Recordset
Dim Sql As String
Dim Sql2 As String
Dim cValues As String
Dim AcumAnt As Long
Dim Kilos As Long
Dim KilosMed As Long
Dim Nregs As Integer
Dim SocioAnt As Long
Dim NombreAnt As String
Dim Diferencia As Long
Dim Entro As Boolean
Dim Importe As Currency
Dim SqlInsert As String

Dim rs3 As ADODB.Recordset
Dim Sql3 As String
Dim CadValues As String


    On Error GoTo eCargarTablaTemporal

    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    SqlInsert = "insert into tmpinformes (codusu, codigo1, nombre1, importe1) values "
    
    Sql2 = "select " & vUsu.Codigo & ", rsocios.codsocio, rsocios.nomsocio from rsocios "
    If nSelect1 <> "" Then Sql2 = Sql2 & " where  " & nSelect1
    Sql2 = Sql2 & " order by 1,2"
    
    CadValues = ""
    
    Set RS = New ADODB.Recordset
    RS.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        Sql3 = "select importe from raportacion where codaport = 0 and codsocio = " & DBSet(RS!Codsocio, "N")
        Sql3 = Sql3 & " and fecaport in (select max(fecaport) from raportacion where codaport = 0 and codsocio = " & DBSet(RS!Codsocio, "N") & ")"
        
        Set rs3 = New ADODB.Recordset
        rs3.Open Sql3, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(RS!Codsocio, "N") & "," & DBSet(RS!nomsocio, "T") & "," & DBSet(rs3!Importe * (-1), "N") & "),"
        End If
        Set rs3 = Nothing
        
        RS.MoveNext
    Wend
    
    Set RS = Nothing
    
    If CadValues <> "" Then
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        conn.Execute SqlInsert & CadValues
    End If
    
    CargarTablaTemporal3 = True
    
    Exit Function
    
eCargarTablaTemporal:
    MuestraError Err.Number, "Cargar Tabla Temporal", Err.Description
End Function



Private Function ActualizarRegularizacionBajaSocio()
Dim Sql As String
Dim RS As ADODB.Recordset
Dim Sql2 As String
Dim SqlValues As String
Dim Descripc As String
Dim Campanya As String
Dim ImporIni As Currency
Dim Importe As Currency
Dim b As Boolean
Dim MensError As String
Dim SqlExiste As String
Dim Fecha As Date

    On Error GoTo eActualizarRegularizacion
        
        
    Sql = "BAJAPO" 'regularizacion de aportaciones
    'Bloquear para que nadie mas pueda realizarlo
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se puede realizar el proceso de Regularización de Aportaciones de Baja Socios. Hay otro usuario realizándolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
        
    conn.BeginTrans
    ConnConta.BeginTrans

    Sql = "select * from tmpinformes where codusu = " & vUsu.Codigo
    Sql = Sql & " order by codigo1 "
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Sql2 = "insert into raportacion (codsocio,fecaport,codaport,descripcion,campanya,kilos,importe) values "

    Campanya = Mid(Format(Year(CDate(vParam.FecIniCam)), "0000"), 3, 2) & "/" & Mid(Format(Year(CDate(vParam.FecFinCam)), "0000"), 3, 2)
    Descripc = "BAJA SOCIO"

    b = True

    Pb7.visible = True
    Pb7.Max = TotalRegistrosConsulta(Sql)
    Pb7.Value = 0
    
    Fecha = txtcodigo(54).Text
    
    While Not RS.EOF And b
        IncrementarProgresNew Pb7, 1
    
        SqlValues = ""
        
        Importe = DBLet(RS!importe1, "N")
    
        SqlExiste = "select count(*) from raportacion where codsocio = " & DBSet(RS!Codigo1, "N") & " and codaport=9 and fecaport=" & DBSet(Fecha, "F")
        b = (TotalRegistros(SqlExiste) = 0)
        
        If Not b Then
            MsgBox "El registro para el socio " & Format(DBLet(RS!Codigo1, "N"), "000000") & " de fecha " & DBSet(Fecha, "F") & " y tipo 0 existe. Revise.", vbExclamation
        Else
            SqlValues = "(" & DBSet(RS!Codigo1, "N") & "," & DBSet(Fecha, "F") & ",9," & DBSet(Descripc, "T") & ","
            SqlValues = SqlValues & DBSet(Campanya, "T") & ",0," & DBSet(Importe, "N") & ")"
            
            conn.Execute Sql2 & SqlValues
            
            MensError = "Insertando pago en tesoreria"
            b = InsertarEnTesoreriaNewAPO(MensError, RS!Codigo1, DBLet(Importe, "N"), txtcodigo(57).Text, txtcodigo(55).Text, txtcodigo(56).Text, txtcodigo(58).Text, CStr(Fecha), 2)
            If Not b Then
                MsgBox "Error: " & MensError, vbExclamation
            End If
            
        End If
    
    
        RS.MoveNext
    Wend
    Set RS = Nothing
    
eActualizarRegularizacion:
    If Err.Number <> 0 Or Not b Then
        ActualizarRegularizacionBajaSocio = False
        conn.RollbackTrans
        ConnConta.RollbackTrans
    Else
        ActualizarRegularizacionBajaSocio = True
        conn.CommitTrans
        ConnConta.CommitTrans
    End If
    
    DesBloqueoManual ("BAJAPO") 'regularizacion de aportaciones
    
    Screen.MousePointer = vbDefault
    
End Function


Private Sub CargarListView(Index As Integer)
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarList

    'Los encabezados
    ListView1(Index).ColumnHeaders.Clear

 '   ListView1.ColumnHeaders.Add , , "Tipo", 650
    ListView1(Index).ColumnHeaders.Add , , "Tipo Movimiento", 2750
    
    Sql = "SELECT codtipom, nomtipom "
    Sql = Sql & " FROM usuarios.stipom "
    '[Monica]28/03/2014: en el caso de Bolbaite dejo seleccionar las facturas de almazara y de bodega si tuvieran
    If vParamAplic.Cooperativa = 14 Then
        Sql = Sql & " WHERE stipom.tipodocu in (1,2,3,4,7,8,9,10)"
    Else
        Sql = Sql & " WHERE stipom.tipodocu in (1,2,3,4)"
    End If
    Sql = Sql & " ORDER BY codtipom "
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        Set ItmX = ListView1(Index).ListItems.Add
        ItmX.Text = RS.Fields(1).Value ' Format(Rs.Fields(0).Value)
        ItmX.Key = RS.Fields(0).Value
'        ItmX.SubItems(1) = Rs.Fields(1).Value
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing

ECargarList:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar Tipo de Facturas.", Err.Description
    End If
End Sub


Private Function InsertarAportacionesBolbaite(Tabla As String, vWhere As String)
Dim Sql As String
Dim RS As ADODB.Recordset
Dim Sql2 As String
Dim SqlValues As String
Dim ImporIni As Currency
Dim Importe As Currency
Dim b As Boolean
Dim MensError As String
Dim SqlExiste As String

    On Error GoTo eInsertarAportacionesBolbaite
        
        
    Sql = "INSAPO" 'Integracion de aportaciones
    'Bloquear para que nadie mas pueda realizarlo
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se puede realizar el proceso de Inserción de Aportaciones. Hay otro usuario realizándolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
        
    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    LOG.Insertar 3, vUsu, "Insertar Aportaciones: " & vbCrLf & Tabla & vbCrLf & vWhere
    Set LOG = Nothing
    '-----------------------------------------------------------------------------
        
    conn.BeginTrans

    Sql = "select * from " & Tabla
    If vWhere <> "" Then Sql = Sql & " WHERE " & vWhere
    Sql = Sql & " order by codtipom, numfactu, fecfactu "
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    b = True


    SqlValues = ""

    Pb8.visible = True
    Pb8.Max = TotalRegistrosConsulta(Sql)
    Pb8.Value = 0
    
   
    Sql = "insert into raportacion (codsocio,fecaport,codaport,descripcion,campanya,kilos,importe,codtipom,numfactu,intconta,porcaport) values "
    
    While Not RS.EOF And b
        IncrementarProgresNew Pb8, 1
    
        Sql2 = "select * from raportacion where fecaport = " & DBSet(RS!fecfactu, "F")
        Sql2 = Sql2 & " and codtipom = " & DBSet(RS!CodTipom, "T")
        Sql2 = Sql2 & " and numfactu = " & DBSet(RS!numfactu, "N")
        
        If TotalRegistros(Sql2) = 0 Then
            Importe = Round2(DBLet(RS!BaseReten) * ImporteSinFormato(ComprobarCero(txtcodigo(69).Text)) / 100, 2)
        
            SqlValues = SqlValues & "(" & DBSet(RS!Codsocio, "N") & "," & DBSet(RS!fecfactu, "F") & "," & DBSet(txtcodigo(68).Text, "N") & ","
            SqlValues = SqlValues & DBSet(txtcodigo(63).Text, "T") & ",' ',0," & DBSet(Importe, "N") & "," & DBSet(RS!CodTipom, "T") & ","
            SqlValues = SqlValues & DBSet(RS!numfactu, "N") & ",0," & DBSet(txtcodigo(69).Text, "N") & "),"
        End If
    
        RS.MoveNext
    Wend
    Set RS = Nothing
    
    If SqlValues <> "" Then
        SqlValues = Mid(SqlValues, 1, Len(SqlValues) - 1)
        conn.Execute Sql & SqlValues
    End If
    
    
eInsertarAportacionesBolbaite:
    If Err.Number <> 0 Or Not b Then
        InsertarAportacionesBolbaite = False
        conn.RollbackTrans
        MsgBox "Se ha producido un error " & MensError, vbExclamation
    Else
        InsertarAportacionesBolbaite = True
        conn.CommitTrans
    End If
    
    DesBloqueoManual ("INSAPO") 'Insertar aportaciones
    
    Screen.MousePointer = vbDefault
    
End Function



Private Function InsertarAportacionesObligatoriasBolbaite(Tabla As String, vWhere As String)
Dim Sql As String
Dim RS As ADODB.Recordset
Dim Sql2 As String
Dim SqlValues As String
Dim ImporIni As Currency
Dim Importe As Currency
Dim b As Boolean
Dim MensError As String
Dim SqlExiste As String

    On Error GoTo eInsertarAportacionesObligatoriasBolbaite
        
        
    Sql = "INSAPO" 'Integracion de aportaciones
    'Bloquear para que nadie mas pueda realizarlo
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se puede realizar el proceso de Inserción de Aportaciones Obligatorias. Hay otro usuario realizándolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
        
    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    LOG.Insertar 3, vUsu, "Insertar Aport.Obligatorias: " & vbCrLf & Tabla & vbCrLf & vWhere
    Set LOG = Nothing
    '-----------------------------------------------------------------------------
        
    conn.BeginTrans

    Sql = "select * from " & Tabla
    If vWhere <> "" Then Sql = Sql & " WHERE " & vWhere
    Sql = Sql & " order by codsocio"
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    b = True


    SqlValues = ""

    Pb9.visible = True
    Pb9.Max = TotalRegistrosConsulta(Sql)
    Pb9.Value = 0
    
    Sql = "insert into raportacion (codsocio,fecaport,codaport,descripcion,campanya,kilos,importe,codtipom,numfactu,intconta,porcaport) values "
    
    While Not RS.EOF And b
        IncrementarProgresNew Pb9, 1
    
        Sql2 = "select * from raportacion where fecaport = " & DBSet(txtcodigo(74).Text, "F")
        Sql2 = Sql2 & " and codsocio = " & DBSet(RS!Codsocio, "N")
        Sql2 = Sql2 & " and codaport = " & DBSet(txtcodigo(71).Text, "N")
        
        If TotalRegistros(Sql2) = 0 Then
            Importe = ImporteSinFormato(txtcodigo(73).Text)
        
            SqlValues = SqlValues & "(" & DBSet(RS!Codsocio, "N") & "," & DBSet(txtcodigo(74).Text, "F") & "," & DBSet(txtcodigo(71).Text, "N") & ","
            SqlValues = SqlValues & DBSet(txtcodigo(72).Text, "T") & ",' ',0," & DBSet(Importe, "N") & "," & ValorNulo & ","
            SqlValues = SqlValues & "0,0,0),"
        End If
    
        RS.MoveNext
    Wend
    Set RS = Nothing
    
    If SqlValues <> "" Then
        SqlValues = Mid(SqlValues, 1, Len(SqlValues) - 1)
        conn.Execute Sql & SqlValues
    End If
    
    
eInsertarAportacionesObligatoriasBolbaite:
    If Err.Number <> 0 Or Not b Then
        InsertarAportacionesObligatoriasBolbaite = False
        conn.RollbackTrans
        MsgBox "Se ha producido un error " & MensError, vbExclamation
    Else
        InsertarAportacionesObligatoriasBolbaite = True
        conn.CommitTrans
    End If
    
    DesBloqueoManual ("INSAPO") 'Insertar aportaciones
    
    Screen.MousePointer = vbDefault
    
End Function


Private Function IntegracionAportacionesTesoreriaBolbaite(Tabla As String, vWhere As String)
Dim Sql As String
Dim RS As ADODB.Recordset
Dim Sql2 As String
Dim SqlValues As String
Dim Descripc As String
Dim Campanya As String
Dim ImporIni As Currency
Dim Importe As Currency
Dim b As Boolean
Dim MensError As String
Dim SqlExiste As String

    On Error GoTo eIntegracionAportacionesTesoreria
        
        
    Sql = "INTAPO" 'Integracion de aportaciones
    'Bloquear para que nadie mas pueda realizarlo
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se puede realizar el proceso de Integración de Aportaciones. Hay otro usuario realizándolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
        
    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    LOG.Insertar 3, vUsu, "Int.Tesoreria Aportaciones: " & vbCrLf & Tabla & vbCrLf & vWhere
    Set LOG = Nothing
    '-----------------------------------------------------------------------------
        
    conn.BeginTrans
    ConnConta.BeginTrans

    Sql = "select * from " & Tabla
    If vWhere <> "" Then Sql = Sql & " WHERE " & vWhere
    Sql = Sql & " order by 1 "
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    b = True

    Pb10.visible = True
    Pb10.Max = TotalRegistrosConsulta(Sql)
    Pb10.Value = 0
    
    While Not RS.EOF And b
        IncrementarProgresNew Pb10, 1
    
        MensError = "Insertando cobro en tesoreria" & vbCrLf & vbCrLf
        b = InsertarEnTesoreriaAPOBol(MensError, RS)  'Rs!Codsocio, DBLet(Rs!Importe, "N"))
    
        RS.MoveNext
    Wend
    Set RS = Nothing
    
    If b Then
        MensError = "Actualizando Aportaciones" & vbCrLf & vbCrLf
        b = ActualizarAportacionesBol(MensError, Tabla, vWhere)
    End If
    
eIntegracionAportacionesTesoreria:
    If Err.Number <> 0 Or Not b Then
        IntegracionAportacionesTesoreriaBolbaite = False
        conn.RollbackTrans
        ConnConta.RollbackTrans
        MsgBox "Se ha producido un error " & MensError, vbExclamation
        
    Else
        IntegracionAportacionesTesoreriaBolbaite = True
        conn.CommitTrans
        ConnConta.CommitTrans
    End If
    
    DesBloqueoManual ("INTAPO") 'Integracion de aportaciones en tesoreria
    
    Screen.MousePointer = vbDefault
    
End Function

Private Function CargarTemporalBol(nTabla1 As String, nSelect1 As String) As Boolean
Dim RS As ADODB.Recordset
Dim Sql As String
Dim Sql2 As String
Dim cValues As String
Dim AcumAnt As Long
Dim Kilos As Long
Dim KilosMed As Long
Dim Nregs As Integer
Dim SocioAnt As Long
Dim NombreAnt As String
Dim Diferencia As Long
Dim Entro As Boolean
Dim Importe As Currency


    On Error GoTo eCargarTablaTemporal

    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    Sql = "insert into tmpinformes (codusu, codigo1, fecha1, importe1, nombre1) "
    
    Sql2 = " select " & vUsu.Codigo & ", raportacion.codsocio, raportacion.fecaport, raportacion.numfactu, raportacion.codtipom "
    Sql2 = Sql2 & " from " & nTabla1
    
    If nSelect1 <> "" Then Sql2 = Sql2 & " where " & nSelect1
    
    conn.Execute Sql & Sql2

    CargarTemporalBol = True
    
    Exit Function
    
eCargarTablaTemporal:
    MuestraError Err.Number, "Cargar Tabla Temporal", Err.Description
End Function


Private Function InsertarEnTesoreriaAPOBol(MenError As String, ByRef RS As ADODB.Recordset) As Boolean
'Guarda datos de Tesoreria en tablas: conta.scobros
' Tipo: 0 = almazara
'       1 = bodega

Dim b As Boolean
Dim Sql As String, Text33csb As String, Text41csb As String
Dim Sql4 As String
Dim Rs4 As ADODB.Recordset

Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim I As Byte
Dim DigConta As String
Dim CC As String
Dim vrefer As String
Dim letraser As String
Dim Rsx As ADODB.Recordset
Dim vSocio As cSocio
Dim Seccion As Integer
Dim FecVen As String
Dim ForpaNeg As String
Dim ForpaPos As String
Dim CtaBan As String
Dim fecfactu As String
Dim numfactu As String
Dim Importe As Currency

Dim Text1csb As String
Dim Text2csb As String


    On Error GoTo EInsertarTesoreriaNew

    b = False
    InsertarEnTesoreriaAPOBol = False
    CadValues = ""
    CadValues2 = ""
    
    Seccion = vParamAplic.Seccionhorto
    
    Set vSocio = New cSocio
    If vSocio.LeerDatos(CStr(RS!Codsocio)) Then
        If vSocio.LeerDatosSeccion(CStr(RS!Codsocio), CStr(Seccion)) Then
            FecVen = txtcodigo(86).Text
            ForpaNeg = txtcodigo(84).Text
            ForpaPos = txtcodigo(85).Text
            CtaBan = txtcodigo(83).Text
            fecfactu = RS!fecaport
            If RS!numfactu = 0 Then
                letraser = ""
                letraser = DevuelveDesdeBDNew(cAgro, "usuarios.stipom", "letraser", "codtipom", "APO", "T")
                numfactu = Mid(Format(Year(CDate(fecfactu)), "0000"), 3, 2) & Format(vSocio.Codigo, "000000")
            Else
                letraser = ""
                letraser = DevuelveDesdeBDNew(cAgro, "usuarios.stipom", "letraser", "codtipom", DBLet(RS!CodTipom), "T")
                numfactu = DBLet(RS!numfactu, "N")
            End If
            
            Importe = DBLet(RS!Importe)
            
            If DBLet(Importe, "N") >= 0 Then
                Text33csb = "'Cargo Aportaciones de " & Format(DBLet(fecfactu, "F"), "dd/mm/yy") & "'"
                Text41csb = "de " & DBSet(Importe, "N")
    
                CC = DBLet(vSocio.Digcontrol, "T")
                If DBLet(vSocio.Digcontrol, "T") = "**" Then CC = "00"
    
                CadValuesAux2 = "(" & DBSet(letraser, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(fecfactu, "F") & ", 9," & DBSet(vSocio.CtaProv, "T") & ","
                CadValues2 = CadValuesAux2 & DBSet(ForpaPos, "N") & "," & DBSet(FecVen, "F") & "," & DBSet(Importe, "N") & ","
                CadValues2 = CadValues2 & DBSet(CtaBan, "T") & "," & DBSet(vSocio.Banco, "N", "S") & "," & DBSet(vSocio.Sucursal, "N", "S") & ","
                CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(vSocio.CuentaBan, "T", "S") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & ",1" ')"
                '[Monica]22/11/2013: Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    CadValues2 = CadValues2 & ", " & DBSet(vSocio.Iban, "T", "S") & ") "
                Else
                    CadValues2 = CadValues2 & ") "
                End If
                
    
                'Insertamos en la tabla scobro de la CONTA
                Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
                Sql = Sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
                Sql = Sql & " text33csb, text41csb, agente" ') "
                '[Monica]22/11/2013: Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    Sql = Sql & ", iban) "
                Else
                    Sql = Sql & ") "
                End If
                
                Sql = Sql & " VALUES " & CadValues2
                
                ConnConta.Execute Sql
            
            Else
                '[Monica]01/09/2014: añadido esto, si el importe es negativo lo tengo que cambiar a positivo
                Importe = Importe * (-1)
            
            
                Text1csb = "'Abono Aportaciones de " & Format(DBLet(fecfactu, "F"), "dd/mm/yy") & "'"
                Text2csb = "de " & DBSet(Importe, "N")
    
                CC = DBLet(vSocio.Digcontrol, "T")
                If DBLet(vSocio.Digcontrol, "T") = "**" Then CC = "00"
    
                CadValuesAux2 = "(" & DBSet(vSocio.CtaProv, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(fecfactu, "F") & ", 9,"
                CadValues2 = CadValuesAux2 & DBSet(ForpaNeg, "N") & "," & DBSet(FecVen, "F") & "," & DBSet(Importe, "N") & ","
                CadValues2 = CadValues2 & DBSet(CtaBan, "T") & "," & DBSet(vSocio.Banco, "N", "S") & "," & DBSet(vSocio.Sucursal, "N", "S") & ","
                CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(vSocio.CuentaBan, "T", "S") & ","
                CadValues2 = CadValues2 & Text1csb & "," & DBSet(Text2csb, "T") '& ")"
                '[Monica]22/11/2013: Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    CadValues2 = CadValues2 & ", " & DBSet(vSocio.Iban, "T", "S") & ") "
                Else
                    CadValues2 = CadValues2 & ") "
                End If
    
                'Insertamos en la tabla scobro de la CONTA
                Sql = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,entidad,oficina,cc,cuentaba,text1csb,text2csb" ') "
                '[Monica]22/11/2013: Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    Sql = Sql & ", iban) "
                Else
                    Sql = Sql & ") "
                End If
                
                Sql = Sql & " VALUES " & CadValues2
                
                ConnConta.Execute Sql
            
            End If
        End If
    End If

    b = True
    
EInsertarTesoreriaNew:
    If Err.Number <> 0 Then
        b = False
        MenError = Err.Description
    End If
    InsertarEnTesoreriaAPOBol = b
End Function



Private Function InsertarDevolucionesQua(vTabla As String, vSelect As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim RS As ADODB.Recordset
Dim b As Boolean
Dim CadValues As String
Dim SqlExiste As String
Dim SqlInsert As String
Dim Importe As Currency
Dim NumApor As Long
Dim CodTipoMov As String
Dim vTipoMov As CTiposMov
Dim devuelve As String
Dim Existe As Boolean
    
    On Error GoTo eActualizarDevoluciones
    
    InsertarDevolucionesQua = False
    
    Sql = "DEVAPO" 'devolucion aportaciones
    'Bloquear para que nadie mas pueda realizarlo
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se puede realizar el proceso de Devolución de Aportaciones. Hay otro usuario realizándolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
        
    conn.BeginTrans

    
'    NumApor = DevuelveValor("select contador from usuarios.stipom where codtipom = 'APO'")
'
'
'    Sql = "insert into raporhco (numaport,codsocio,codcampo,poligono,parcela,codparti,codvarie,impaport,fecaport,anoplant,observac,supcoope,ejercicio,intconta)"
'    Sql2 = " select @nroapor:=@nroapor + 1, codsocio,codcampo,poligono,parcela,codparti,raporhco.codvarie,impaport * (-1),"
'    Sql2 = Sql2 & DBSet(txtcodigo(112).Text, "F") & ","
'    Sql2 = Sql2 & "anoplant,observac,supcoope,ejercicio,0 "
'    Sql2 = Sql2 & " from " & vTabla & ", (select @nroapor:= " & NumApor & ") aaa " '(select contador from usuarios.stipom where codtipom = 'APO')) aaa"
'    If vSelect <> "" Then Sql2 = Sql2 & " where " & vSelect
'
'    NumApor = NumApor + TotalRegistrosConsulta(Sql2)
'
'
'    conn.Execute Sql & Sql2
'
'    Sql = "update usuarios.stipom set contador = " & DBSet(NumApor, "N") & " where codtipom = 'APO'" ' (select max(numaport) from ariagro.raporhco where fecaport = " & DBSet(txtcodigo(112).Text, "F") & ") and codtipom = 'APO'"
'    conn.Execute Sql
    
    
    b = True
    
    '[Monica]15/09/2014: las aportaciones de cualquier campaña se insertarán siempre en la campaña actual
    SqlInsert = "insert into ariagro.raporhco (numaport,codsocio,codcampo,poligono,parcela,codparti,codvarie,impaport," & _
                "fecaport,anoplant,observac,supcoope,ejercicio,intconta) values "
    
    Sql = "select raporhco.* from " & vTabla
    Sql = Sql & " where " & vSelect
    
    CargarProgres pb12, TotalRegistrosConsulta(Sql)
    pb12.visible = True
    
    
    CadValues = ""
    CodTipoMov = "APO"
    
    Set vTipoMov = New CTiposMov
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF And b
        IncrementarProgres pb12, 1
        DoEvents
        
        
        NumApor = vTipoMov.ConseguirContador(CodTipoMov)
    
        Do
            devuelve = DevuelveDesdeBDNew(cAgro, "raporhco", "numaport", "numaport", CStr(NumApor), "N")
            If devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (CodTipoMov)
                NumApor = vTipoMov.ConseguirContador(CodTipoMov)
            Else
                Existe = False
            End If
        Loop Until Not Existe
                
        CadValues = "(" & DBSet(NumApor, "N") & "," & DBSet(RS!Codsocio, "N") & "," & DBSet(RS!CodCampo, "N") & ","
        CadValues = CadValues & DBSet(RS!poligono, "N") & "," & DBSet(RS!Parcela, "N") & "," & DBSet(RS!codparti, "N") & ","
        CadValues = CadValues & DBSet(RS!codvarie, "N") & "," & DBSet(RS!ImpAport * (-1), "N") & "," & DBSet(txtcodigo(112).Text, "F") & ","
        CadValues = CadValues & DBSet(RS!anoplant, "N") & "," & ValorNulo & "," & DBSet(RS!supcoope, "N") & ","
        CadValues = CadValues & DBSet(txtcodigo(98).Text, "N") & ",0)"
                
        conn.Execute SqlInsert & CadValues
                
        b = vTipoMov.IncrementarContador(CodTipoMov)
        
        RS.MoveNext
    Wend
    
    Set RS = Nothing
    Set vTipoMov = Nothing
    
    If b Then
        InsertarDevolucionesQua = True
        Pb5.visible = False
        conn.CommitTrans
        Exit Function
    End If
    
    
eActualizarDevoluciones:
    pb12.visible = False
    If Err.Number <> 0 Then
        InsertarDevolucionesQua = False
        conn.RollbackTrans
    Else
        InsertarDevolucionesQua = True
        conn.CommitTrans
    End If
    
    DesBloqueoManual ("DEVAPO") 'devolucion de aportaciones
    
    Screen.MousePointer = vbDefault
    
End Function


Private Function InsertarTemporalDevolQua(vTabla As String, vSelect As String) As Boolean
Dim Sql As String

    On Error GoTo eInsertarTemporal
    
    InsertarTemporalDevolQua = False
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
                                            'socio,  campo,    variedad,  importe
    Sql = "insert into tmpinformes (codusu, codigo1, importe1, importe2, importe3)"
    Sql = Sql & " select " & vUsu.Codigo & ", rapohco.codsocio, raporhco.codcampo, raporhco.codvarie, sum(coalesce(raporhco.impaport,0)) importe "
    Sql = Sql & " from " & vTabla
    Sql = Sql & " where " & vSelect
    Sql = Sql & " group by 1,2,3,4 "
    Sql = Sql & " order by 1,2,3,4 "
    
    conn.Execute Sql

    InsertarTemporalDevolQua = True
    Exit Function

eInsertarTemporal:
    MuestraError Err.Number, "Insertar Temporal", Err.Description
End Function

