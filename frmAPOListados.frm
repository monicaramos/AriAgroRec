VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAPOListados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   8955
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   8145
   Icon            =   "frmAPOListados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameRegularizacion 
      Height          =   8880
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   8085
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
         Height          =   1920
         Left            =   120
         TabIndex        =   51
         Top             =   5790
         Width           =   7710
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
            Index           =   18
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   12
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1440
            Width           =   1365
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
            Index           =   18
            Left            =   3840
            Locked          =   -1  'True
            TabIndex        =   54
            Top             =   1440
            Width           =   3810
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
            Index           =   15
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   9
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   360
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
            Index           =   16
            Left            =   3525
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   720
            Width           =   4125
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
            Index           =   16
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   10
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   720
            Width           =   1095
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
            Index           =   17
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   11
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1080
            Width           =   1095
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
            Index           =   17
            Left            =   3525
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   1080
            Width           =   4125
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   2160
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   1440
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
            TabIndex        =   58
            Top             =   1485
            Width           =   1935
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   5
            Left            =   2160
            Picture         =   "frmAPOListados.frx":000C
            ToolTipText     =   "Buscar fecha"
            Top             =   360
            Width           =   240
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
            Index           =   17
            Left            =   180
            TabIndex        =   57
            Top             =   405
            Width           =   1920
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
            Height          =   240
            Index           =   4
            Left            =   180
            TabIndex        =   56
            Top             =   765
            Width           =   1920
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   2160
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   720
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   9
            Left            =   2160
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   1080
            Width           =   240
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
            TabIndex        =   55
            Top             =   1125
            Width           =   2160
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos de Selección"
         DragMode        =   1  'Automatic
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
         Height          =   4890
         Left            =   120
         TabIndex        =   40
         Top             =   840
         Width           =   7710
         Begin VB.Frame FrameDevolucion 
            BorderStyle     =   0  'None
            Height          =   780
            Left            =   90
            TabIndex        =   448
            Top             =   3960
            Width           =   7575
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
               Index           =   123
               Left            =   2340
               MaxLength       =   10
               TabIndex        =   8
               Tag             =   "Código Postal|T|S|||clientes|codposta|||"
               Top             =   90
               Width           =   1050
            End
            Begin VB.Label Label4 
               Caption         =   "Precio Devolución"
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
               Index           =   144
               Left            =   135
               TabIndex        =   449
               Top             =   135
               Width           =   2130
            End
         End
         Begin VB.CheckBox ChkDevolucion 
            Caption         =   "Devolucion de Capital Social"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   240
            Left            =   225
            TabIndex        =   7
            Top             =   3690
            Width           =   3075
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
            Index           =   14
            Left            =   2430
            MaxLength       =   10
            TabIndex        =   6
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   3105
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
            Left            =   3540
            Locked          =   -1  'True
            TabIndex        =   42
            Text            =   "Text5"
            Top             =   885
            Width           =   4035
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
            Left            =   3540
            Locked          =   -1  'True
            TabIndex        =   41
            Text            =   "Text5"
            Top             =   510
            Width           =   4035
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
            Index           =   11
            Left            =   2460
            MaxLength       =   16
            TabIndex        =   1
            Top             =   885
            Width           =   1050
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
            Left            =   2460
            MaxLength       =   16
            TabIndex        =   0
            Top             =   510
            Width           =   1050
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
            Left            =   2445
            MaxLength       =   10
            TabIndex        =   3
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   1890
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
            Left            =   2445
            MaxLength       =   10
            TabIndex        =   2
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   1530
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
            Left            =   5640
            MaxLength       =   10
            TabIndex        =   5
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   2445
            Width           =   1050
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
            Left            =   2445
            MaxLength       =   10
            TabIndex        =   4
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   2445
            Width           =   1050
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Regularización"
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
            Index           =   10
            Left            =   210
            TabIndex        =   60
            Top             =   2835
            Width           =   2130
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   4
            Left            =   2160
            Picture         =   "frmAPOListados.frx":0097
            ToolTipText     =   "Buscar fecha"
            Top             =   3150
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   2160
            MouseIcon       =   "frmAPOListados.frx":0122
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar socio"
            Top             =   885
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   2
            Left            =   2175
            MouseIcon       =   "frmAPOListados.frx":0274
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar socio"
            Top             =   510
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Socio"
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
            Index           =   9
            Left            =   255
            TabIndex        =   50
            Top             =   345
            Width           =   540
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
            Index           =   8
            Left            =   1395
            TabIndex        =   49
            Top             =   885
            Width           =   735
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
            Index           =   7
            Left            =   1410
            TabIndex        =   48
            Top             =   510
            Width           =   780
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   1
            Left            =   2145
            Picture         =   "frmAPOListados.frx":03C6
            ToolTipText     =   "Buscar fecha"
            Top             =   1890
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   0
            Left            =   2145
            Picture         =   "frmAPOListados.frx":0451
            ToolTipText     =   "Buscar fecha"
            Top             =   1530
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
            Index           =   6
            Left            =   1305
            TabIndex        =   47
            Top             =   1890
            Width           =   735
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
            Index           =   5
            Left            =   1305
            TabIndex        =   46
            Top             =   1530
            Width           =   780
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
            Index           =   4
            Left            =   210
            TabIndex        =   45
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Precio Aumento Kilos"
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
            Index           =   3
            Left            =   225
            TabIndex        =   44
            Top             =   2370
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Precio Disminución Kilos"
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
            Left            =   3600
            TabIndex        =   43
            Top             =   2370
            Width           =   1815
         End
      End
      Begin VB.CommandButton CmdAcepRegul 
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
         Left            =   5595
         TabIndex        =   13
         Top             =   8355
         Width           =   1065
      End
      Begin VB.CommandButton CmdCancel 
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
         Index           =   0
         Left            =   6765
         TabIndex        =   14
         Top             =   8355
         Width           =   1065
      End
      Begin MSComctlLib.ProgressBar pb2 
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   8070
         Visible         =   0   'False
         Width           =   7785
         _ExtentX        =   13732
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
         TabIndex        =   59
         Top             =   7740
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
         TabIndex        =   39
         Top             =   300
         Width           =   5160
      End
   End
   Begin VB.Frame FrameCobros 
      Height          =   6555
      Left            =   0
      TabIndex        =   24
      Top             =   30
      Width           =   8085
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
         Index           =   125
         Left            =   1665
         MaxLength       =   10
         TabIndex        =   21
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   4890
         Width           =   1050
      End
      Begin VB.CheckBox chkNegativas 
         Caption         =   "Sólo Negativas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5325
         TabIndex        =   414
         Tag             =   "Correo|N|N|||rsocios|correo||N|"
         Top             =   3450
         Width           =   1965
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Imprimir resumen"
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
         Height          =   240
         Left            =   5325
         TabIndex        =   412
         Top             =   3180
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   3960
         TabIndex        =   61
         Top             =   3645
         Width           =   2955
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
            Index           =   12
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   62
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   540
            Width           =   1350
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   6
            Left            =   1050
            Picture         =   "frmAPOListados.frx":04DC
            ToolTipText     =   "Buscar fecha"
            Top             =   540
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Certificado"
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
            Index           =   18
            Left            =   30
            TabIndex        =   63
            Top             =   240
            Width           =   1815
         End
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
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   20
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   4170
         Width           =   1050
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
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   19
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3420
         Width           =   1050
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
         Left            =   1725
         MaxLength       =   10
         TabIndex        =   18
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2670
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
         Index           =   2
         Left            =   1725
         MaxLength       =   10
         TabIndex        =   17
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2265
         Width           =   1350
      End
      Begin VB.CommandButton CmdCancel 
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
         Index           =   1
         Left            =   6690
         TabIndex        =   23
         Top             =   5730
         Width           =   1065
      End
      Begin VB.CommandButton CmdAceptar 
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
         Left            =   5520
         TabIndex        =   22
         Top             =   5730
         Width           =   1065
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
         Left            =   1725
         MaxLength       =   16
         TabIndex        =   15
         Top             =   1125
         Width           =   1035
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
         Left            =   1725
         MaxLength       =   16
         TabIndex        =   16
         Top             =   1545
         Width           =   1035
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
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "Text5"
         Top             =   1125
         Width           =   4950
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
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "Text5"
         Top             =   1545
         Width           =   4950
      End
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   255
         Left            =   465
         TabIndex        =   36
         Top             =   5355
         Visible         =   0   'False
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         Caption         =   "Precio Devolución"
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
         Index           =   146
         Left            =   495
         TabIndex        =   450
         Top             =   4590
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Precio Disminución Kilos"
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
         Left            =   495
         TabIndex        =   35
         Top             =   3870
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Precio Aumento Kilos"
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
         Left            =   495
         TabIndex        =   34
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
         TabIndex        =   33
         Top             =   315
         Width           =   5160
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
         Index           =   16
         Left            =   495
         TabIndex        =   32
         Top             =   1965
         Width           =   1815
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
         Left            =   705
         TabIndex        =   31
         Top             =   2265
         Width           =   645
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
         Left            =   705
         TabIndex        =   30
         Top             =   2670
         Width           =   600
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1455
         Picture         =   "frmAPOListados.frx":0567
         ToolTipText     =   "Buscar fecha"
         Top             =   2265
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1455
         Picture         =   "frmAPOListados.frx":05F2
         ToolTipText     =   "Buscar fecha"
         Top             =   2670
         Width           =   240
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
         Left            =   735
         TabIndex        =   29
         Top             =   1125
         Width           =   645
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
         Left            =   750
         TabIndex        =   28
         Top             =   1545
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
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
         Left            =   495
         TabIndex        =   27
         Top             =   840
         Width           =   540
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1440
         MouseIcon       =   "frmAPOListados.frx":067D
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1125
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1425
         MouseIcon       =   "frmAPOListados.frx":07CF
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1545
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7335
      Top             =   4185
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameIntTesorBol 
      Height          =   7530
      Left            =   0
      TabIndex        =   291
      Top             =   0
      Width           =   8085
      Begin VB.CommandButton CmdAcepIntTesBol 
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
         Left            =   5715
         TabIndex        =   305
         Top             =   6660
         Width           =   1065
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
         TabIndex        =   312
         Top             =   3810
         Width           =   7800
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
            Index           =   83
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   303
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1575
            Width           =   1050
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
            Index           =   83
            Left            =   3495
            Locked          =   -1  'True
            TabIndex        =   315
            Top             =   1575
            Width           =   4185
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
            Index           =   86
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   300
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   360
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
            Index           =   85
            Left            =   3495
            Locked          =   -1  'True
            TabIndex        =   314
            Top             =   765
            Width           =   4185
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
            Index           =   85
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   301
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   765
            Width           =   1050
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
            Index           =   84
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   302
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1170
            Width           =   1050
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
            Index           =   84
            Left            =   3495
            Locked          =   -1  'True
            TabIndex        =   313
            Top             =   1170
            Width           =   4185
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   41
            Left            =   2130
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   1575
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
            Index           =   23
            Left            =   180
            TabIndex        =   319
            Top             =   1620
            Width           =   1890
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   26
            Left            =   2130
            Picture         =   "frmAPOListados.frx":0921
            ToolTipText     =   "Buscar fecha"
            Top             =   360
            Width           =   240
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
            Index           =   95
            Left            =   180
            TabIndex        =   318
            Top             =   405
            Width           =   1890
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
            Height          =   240
            Index           =   22
            Left            =   180
            TabIndex        =   317
            Top             =   810
            Width           =   1770
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   45
            Left            =   2130
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   765
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   44
            Left            =   2130
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   1170
            Width           =   240
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
            Index           =   21
            Left            =   180
            TabIndex        =   316
            Top             =   1215
            Width           =   1860
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
         TabIndex        =   292
         Top             =   780
         Width           =   7800
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
            Index           =   75
            Left            =   3510
            Locked          =   -1  'True
            TabIndex        =   323
            Text            =   "Text5"
            Top             =   2355
            Width           =   4185
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
            Index           =   75
            Left            =   2400
            MaxLength       =   3
            TabIndex        =   299
            Top             =   2355
            Width           =   1035
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
            Index           =   82
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   297
            Text            =   "Text5"
            Top             =   885
            Width           =   4200
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
            Index           =   81
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   295
            Text            =   "Text5"
            Top             =   510
            Width           =   4200
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
            Index           =   82
            Left            =   2400
            MaxLength       =   16
            TabIndex        =   294
            Top             =   870
            Width           =   1050
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
            Index           =   81
            Left            =   2400
            MaxLength       =   16
            TabIndex        =   293
            Top             =   510
            Width           =   1050
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
            Index           =   80
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   298
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   1860
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
            Index           =   79
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   296
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   1485
            Width           =   1350
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Aportación"
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
            Index           =   82
            Left            =   210
            TabIndex        =   324
            Top             =   2265
            Width           =   1560
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   40
            Left            =   2100
            MouseIcon       =   "frmAPOListados.frx":09AC
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar aportación"
            Top             =   2385
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   43
            Left            =   2100
            MouseIcon       =   "frmAPOListados.frx":0AFE
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar socio"
            Top             =   870
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   42
            Left            =   2100
            MouseIcon       =   "frmAPOListados.frx":0C50
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar socio"
            Top             =   510
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Socio"
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
            Index           =   93
            Left            =   225
            TabIndex        =   311
            Top             =   405
            Width           =   540
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
            Index           =   92
            Left            =   1365
            TabIndex        =   310
            Top             =   885
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
            Index           =   91
            Left            =   1380
            TabIndex        =   309
            Top             =   510
            Width           =   690
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   24
            Left            =   2100
            Picture         =   "frmAPOListados.frx":0DA2
            ToolTipText     =   "Buscar fecha"
            Top             =   1860
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   23
            Left            =   2100
            Picture         =   "frmAPOListados.frx":0E2D
            ToolTipText     =   "Buscar fecha"
            Top             =   1485
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
            Index           =   90
            Left            =   1350
            TabIndex        =   308
            Top             =   1890
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
            Index           =   89
            Left            =   1350
            TabIndex        =   306
            Top             =   1515
            Width           =   690
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Aportación"
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
            Index           =   85
            Left            =   210
            TabIndex        =   304
            Top             =   1215
            Width           =   1815
         End
      End
      Begin VB.CommandButton CmdCancel 
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
         Index           =   10
         Left            =   6855
         TabIndex        =   307
         Top             =   6660
         Width           =   1065
      End
      Begin MSComctlLib.ProgressBar Pb10 
         Height          =   255
         Left            =   210
         TabIndex        =   320
         Top             =   6270
         Visible         =   0   'False
         Width           =   7740
         _ExtentX        =   13653
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
         TabIndex        =   322
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
         TabIndex        =   321
         Top             =   270
         Width           =   5160
      End
   End
   Begin VB.Frame FrameInsertarApoBol 
      Height          =   7470
      Left            =   0
      TabIndex        =   234
      Top             =   60
      Width           =   8085
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   2505
         Left            =   135
         TabIndex        =   259
         Top             =   4095
         Width           =   7740
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
            Index           =   68
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   267
            Text            =   "Text5"
            Top             =   300
            Width           =   5040
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
            Index           =   68
            Left            =   1560
            MaxLength       =   3
            TabIndex        =   266
            Top             =   285
            Width           =   1035
         End
         Begin VB.TextBox txtcodigo 
            Height          =   360
            Index           =   63
            Left            =   1560
            MaxLength       =   30
            TabIndex        =   252
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   840
            Width           =   6105
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
            Index           =   69
            Left            =   1560
            MaxLength       =   30
            TabIndex        =   253
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   1470
            Width           =   1020
         End
         Begin MSComctlLib.ProgressBar Pb8 
            Height          =   255
            Left            =   210
            TabIndex        =   260
            Top             =   1890
            Visible         =   0   'False
            Width           =   7530
            _ExtentX        =   13282
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Aportación"
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
            Index           =   73
            Left            =   240
            TabIndex        =   268
            Top             =   0
            Width           =   1560
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   36
            Left            =   1230
            MouseIcon       =   "frmAPOListados.frx":0EB8
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar aportación"
            Top             =   300
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Descripción"
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
            Index           =   65
            Left            =   270
            TabIndex        =   262
            Top             =   630
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Porcentaje de Aportación"
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
            Index           =   76
            Left            =   270
            TabIndex        =   261
            Top             =   1200
            Width           =   2820
         End
      End
      Begin VB.Frame Frame12 
         BorderStyle     =   0  'None
         Height          =   765
         Left            =   390
         TabIndex        =   263
         Top             =   4080
         Width           =   3135
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
            Index           =   70
            Left            =   1290
            MaxLength       =   10
            TabIndex        =   264
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   315
            Width           =   1450
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Recibo"
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
            Index           =   77
            Left            =   0
            TabIndex        =   265
            Top             =   60
            Width           =   1815
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   19
            Left            =   975
            Picture         =   "frmAPOListados.frx":100A
            ToolTipText     =   "Buscar fecha"
            Top             =   300
            Width           =   240
         End
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
         Index           =   62
         Left            =   1680
         MaxLength       =   7
         TabIndex        =   247
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   1545
         Width           =   1140
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
         Index           =   61
         Left            =   1680
         MaxLength       =   7
         TabIndex        =   246
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   1155
         Width           =   1140
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
         Index           =   67
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   236
         Text            =   "Text5"
         Top             =   3645
         Width           =   4995
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
         Index           =   66
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   235
         Text            =   "Text5"
         Top             =   3270
         Width           =   4995
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
         Index           =   67
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   251
         Top             =   3645
         Width           =   1035
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
         Index           =   66
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   250
         Top             =   3270
         Width           =   1035
      End
      Begin VB.CommandButton CmdAcepInsApoBol 
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
         Left            =   5595
         TabIndex        =   254
         Top             =   6615
         Width           =   1065
      End
      Begin VB.CommandButton CmdCancel 
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
         Index           =   8
         Left            =   6735
         TabIndex        =   255
         Top             =   6600
         Width           =   1065
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
         Index           =   65
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   249
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2685
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
         Index           =   64
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   248
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2280
         Width           =   1350
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1065
         Index           =   0
         Left            =   2940
         TabIndex        =   244
         Top             =   1110
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   1879
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
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
         Index           =   75
         Left            =   720
         TabIndex        =   258
         Top             =   1545
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
         Index           =   64
         Left            =   720
         TabIndex        =   257
         Top             =   1185
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nro.Factura"
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
         Index           =   63
         Left            =   390
         TabIndex        =   256
         Top             =   825
         Width           =   1170
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
         Index           =   74
         Left            =   2970
         TabIndex        =   245
         Top             =   825
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   3
         Left            =   7440
         Picture         =   "frmAPOListados.frx":1095
         ToolTipText     =   "Desmarcar todos"
         Top             =   840
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   2
         Left            =   7200
         Picture         =   "frmAPOListados.frx":1A97
         ToolTipText     =   "Marcar todos"
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   35
         Left            =   1365
         MouseIcon       =   "frmAPOListados.frx":82E9
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   3645
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   34
         Left            =   1365
         MouseIcon       =   "frmAPOListados.frx":843B
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   3270
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
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
         Index           =   72
         Left            =   390
         TabIndex        =   243
         Top             =   2985
         Width           =   540
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
         Index           =   71
         Left            =   720
         TabIndex        =   242
         Top             =   3645
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
         Index           =   70
         Left            =   720
         TabIndex        =   241
         Top             =   3270
         Width           =   690
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   21
         Left            =   1365
         Picture         =   "frmAPOListados.frx":858D
         ToolTipText     =   "Buscar fecha"
         Top             =   2670
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   20
         Left            =   1365
         Picture         =   "frmAPOListados.frx":8618
         ToolTipText     =   "Buscar fecha"
         Top             =   2310
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
         Index           =   69
         Left            =   705
         TabIndex        =   240
         Top             =   2670
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
         Index           =   67
         Left            =   720
         TabIndex        =   239
         Top             =   2310
         Width           =   690
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
         Index           =   66
         Left            =   390
         TabIndex        =   238
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
         TabIndex        =   237
         Top             =   315
         Width           =   5160
      End
   End
   Begin VB.Frame FrameAporObligatoria 
      Height          =   6330
      Left            =   -30
      TabIndex        =   269
      Top             =   0
      Width           =   8085
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
         Index           =   74
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   282
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1245
         Width           =   1350
      End
      Begin VB.CommandButton CmdCancel 
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
         Index           =   9
         Left            =   6825
         TabIndex        =   290
         Top             =   5460
         Width           =   1065
      End
      Begin VB.CommandButton CmdAcepApoObli 
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
         Left            =   5685
         TabIndex        =   289
         Top             =   5475
         Width           =   1065
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
         Index           =   78
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   285
         Top             =   2250
         Width           =   1035
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
         Index           =   77
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   284
         Top             =   1860
         Width           =   1035
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
         Index           =   77
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   277
         Text            =   "Text5"
         Top             =   1875
         Width           =   5085
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
         Index           =   78
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   276
         Text            =   "Text5"
         Top             =   2250
         Width           =   5085
      End
      Begin VB.Frame Frame14 
         BorderStyle     =   0  'None
         Height          =   2565
         Left            =   150
         TabIndex        =   270
         Top             =   2820
         Width           =   7785
         Begin MSComctlLib.ProgressBar Pb9 
            Height          =   255
            Left            =   150
            TabIndex        =   272
            Top             =   1980
            Visible         =   0   'False
            Width           =   7620
            _ExtentX        =   13441
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
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
            Index           =   73
            Left            =   1560
            MaxLength       =   12
            TabIndex        =   288
            Top             =   1500
            Width           =   1020
         End
         Begin VB.TextBox txtcodigo 
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
            Index           =   72
            Left            =   1560
            MaxLength       =   30
            TabIndex        =   287
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   840
            Width           =   6180
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
            Index           =   71
            Left            =   1560
            MaxLength       =   3
            TabIndex        =   286
            Top             =   270
            Width           =   1035
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
            Index           =   71
            Left            =   2670
            Locked          =   -1  'True
            TabIndex        =   271
            Text            =   "Text5"
            Top             =   270
            Width           =   5085
         End
         Begin VB.Label Label4 
            Caption         =   "Importe Aportación"
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
            Index           =   80
            Left            =   300
            TabIndex        =   275
            Top             =   1200
            Width           =   1875
         End
         Begin VB.Label Label4 
            Caption         =   "Descripción"
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
            Index           =   79
            Left            =   300
            TabIndex        =   274
            Top             =   630
            Width           =   1815
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   37
            Left            =   1230
            MouseIcon       =   "frmAPOListados.frx":86A3
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar aportación"
            Top             =   300
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Aportación"
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
            Index           =   78
            Left            =   300
            TabIndex        =   273
            Top             =   0
            Width           =   1560
         End
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   22
         Left            =   1365
         Picture         =   "frmAPOListados.frx":87F5
         ToolTipText     =   "Buscar fecha"
         Top             =   1230
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Aportación"
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
         Index           =   81
         Left            =   450
         TabIndex        =   283
         Top             =   945
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
         Left            =   420
         TabIndex        =   281
         Top             =   315
         Width           =   5160
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
         Index           =   88
         Left            =   705
         TabIndex        =   280
         Top             =   1875
         Width           =   645
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
         Index           =   87
         Left            =   705
         TabIndex        =   279
         Top             =   2250
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
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
         Index           =   86
         Left            =   420
         TabIndex        =   278
         Top             =   1590
         Width           =   540
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   39
         Left            =   1380
         MouseIcon       =   "frmAPOListados.frx":8880
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   2220
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   38
         Left            =   1380
         MouseIcon       =   "frmAPOListados.frx":89D2
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1860
         Width           =   240
      End
   End
   Begin VB.Frame FrameIntTesorQua 
      Height          =   7530
      Left            =   0
      TabIndex        =   147
      Top             =   0
      Width           =   8085
      Begin VB.CommandButton CmdCancel 
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
         Index           =   5
         Left            =   6855
         TabIndex        =   171
         Top             =   7005
         Width           =   1065
      End
      Begin VB.CommandButton CmdAcepIntTesQua 
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
         Left            =   5625
         TabIndex        =   170
         Top             =   7005
         Width           =   1065
      End
      Begin VB.Frame Frame7 
         Caption         =   "Datos de Selección"
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
         Height          =   3225
         Left            =   120
         TabIndex        =   156
         Top             =   780
         Width           =   7845
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
            Index           =   48
            Left            =   3555
            Locked          =   -1  'True
            TabIndex        =   182
            Text            =   "Text5"
            Top             =   1860
            Width           =   4155
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
            Index           =   43
            Left            =   3555
            Locked          =   -1  'True
            TabIndex        =   181
            Text            =   "Text5"
            Top             =   1485
            Width           =   4155
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
            Index           =   48
            Left            =   2445
            MaxLength       =   10
            TabIndex        =   162
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   1845
            Width           =   1095
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
            Index           =   43
            Left            =   2445
            MaxLength       =   10
            TabIndex        =   161
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   1485
            Width           =   1095
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
            Index           =   47
            Left            =   2445
            MaxLength       =   10
            TabIndex        =   164
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   2790
            Width           =   1320
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
            Index           =   46
            Left            =   2445
            MaxLength       =   10
            TabIndex        =   163
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   2430
            Width           =   1320
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
            Index           =   45
            Left            =   2445
            MaxLength       =   16
            TabIndex        =   160
            Top             =   930
            Width           =   1095
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
            Index           =   44
            Left            =   2445
            MaxLength       =   16
            TabIndex        =   159
            Top             =   570
            Width           =   1095
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
            Index           =   44
            Left            =   3570
            Locked          =   -1  'True
            TabIndex        =   158
            Text            =   "Text5"
            Top             =   570
            Width           =   4155
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
            Index           =   45
            Left            =   3570
            Locked          =   -1  'True
            TabIndex        =   157
            Text            =   "Text5"
            Top             =   945
            Width           =   4155
         End
         Begin VB.Label Label4 
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
            Index           =   40
            Left            =   210
            TabIndex        =   185
            Top             =   1305
            Width           =   525
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   26
            Left            =   2115
            MouseIcon       =   "frmAPOListados.frx":8B24
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar clase"
            Top             =   1845
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   25
            Left            =   2115
            MouseIcon       =   "frmAPOListados.frx":8C76
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar clase"
            Top             =   1485
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
            Index           =   57
            Left            =   1350
            TabIndex        =   184
            Top             =   1830
            Width           =   690
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
            Index           =   56
            Left            =   1350
            TabIndex        =   183
            Top             =   1470
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Aportacion"
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
            Index           =   55
            Left            =   210
            TabIndex        =   177
            Top             =   2160
            Width           =   1815
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
            Index           =   54
            Left            =   1350
            TabIndex        =   176
            Top             =   2460
            Width           =   735
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
            Index           =   53
            Left            =   1350
            TabIndex        =   175
            Top             =   2790
            Width           =   690
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   15
            Left            =   2100
            Picture         =   "frmAPOListados.frx":8DC8
            ToolTipText     =   "Buscar fecha"
            Top             =   2790
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   14
            Left            =   2100
            Picture         =   "frmAPOListados.frx":8E53
            ToolTipText     =   "Buscar fecha"
            Top             =   2430
            Width           =   240
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
            Index           =   52
            Left            =   1380
            TabIndex        =   174
            Top             =   570
            Width           =   735
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
            Index           =   51
            Left            =   1365
            TabIndex        =   173
            Top             =   945
            Width           =   690
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Socio"
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
            Index           =   41
            Left            =   210
            TabIndex        =   172
            Top             =   420
            Width           =   540
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   24
            Left            =   2115
            MouseIcon       =   "frmAPOListados.frx":8EDE
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar socio"
            Top             =   960
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   23
            Left            =   2130
            MouseIcon       =   "frmAPOListados.frx":9030
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar socio"
            Top             =   570
            Width           =   240
         End
      End
      Begin VB.Frame Frame6 
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
         Height          =   2370
         Left            =   120
         TabIndex        =   148
         Top             =   4020
         Width           =   7845
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
            Index           =   49
            Left            =   2445
            MaxLength       =   10
            TabIndex        =   165
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   300
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
            Index           =   42
            Left            =   3495
            Locked          =   -1  'True
            TabIndex        =   151
            Top             =   1515
            Width           =   4275
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
            Index           =   42
            Left            =   2445
            MaxLength       =   10
            TabIndex        =   168
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1515
            Width           =   1050
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
            Index           =   40
            Left            =   2445
            MaxLength       =   10
            TabIndex        =   167
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1110
            Width           =   1050
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
            Index           =   40
            Left            =   3495
            Locked          =   -1  'True
            TabIndex        =   150
            Top             =   1110
            Width           =   4275
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
            Index           =   34
            Left            =   2445
            MaxLength       =   10
            TabIndex        =   166
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   705
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
            Index           =   33
            Left            =   3810
            Locked          =   -1  'True
            TabIndex        =   149
            Top             =   1920
            Width           =   3960
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
            Index           =   33
            Left            =   2445
            MaxLength       =   10
            TabIndex        =   169
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1920
            Width           =   1350
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   13
            Left            =   2175
            Picture         =   "frmAPOListados.frx":9182
            ToolTipText     =   "Buscar fecha"
            Top             =   300
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Aportación"
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
            Index           =   58
            Left            =   180
            TabIndex        =   186
            Top             =   345
            Width           =   1845
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
            Height          =   195
            Index           =   9
            Left            =   180
            TabIndex        =   155
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   22
            Left            =   2175
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   1515
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   21
            Left            =   2175
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   1110
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
            Height          =   195
            Index           =   8
            Left            =   180
            TabIndex        =   154
            Top             =   1155
            Width           =   1725
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
            Index           =   39
            Left            =   180
            TabIndex        =   153
            Top             =   750
            Width           =   1980
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   12
            Left            =   2175
            Picture         =   "frmAPOListados.frx":920D
            ToolTipText     =   "Buscar fecha"
            Top             =   705
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
            Index           =   7
            Left            =   180
            TabIndex        =   152
            Top             =   1965
            Width           =   1980
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   20
            Left            =   2175
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   1920
            Width           =   240
         End
      End
      Begin MSComctlLib.ProgressBar Pb4 
         Height          =   255
         Left            =   120
         TabIndex        =   178
         Top             =   6720
         Visible         =   0   'False
         Width           =   7830
         _ExtentX        =   13811
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label3 
         Caption         =   "Integración Aportaciones Tesorería"
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
         TabIndex        =   180
         Top             =   270
         Width           =   5160
      End
      Begin VB.Label Label1 
         Caption         =   "lb1"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   180
         TabIndex        =   179
         Top             =   6390
         Visible         =   0   'False
         Width           =   6105
      End
   End
   Begin VB.Frame FrameRegBajaSocios 
      Height          =   5400
      Left            =   0
      TabIndex        =   211
      Top             =   0
      Width           =   8085
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
         Height          =   2070
         Left            =   120
         TabIndex        =   213
         Top             =   2130
         Width           =   7845
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
            Index           =   58
            Left            =   2385
            MaxLength       =   10
            TabIndex        =   227
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1575
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
            Index           =   58
            Left            =   3750
            Locked          =   -1  'True
            TabIndex        =   216
            Top             =   1575
            Width           =   3990
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
            Index           =   57
            Left            =   2385
            MaxLength       =   10
            TabIndex        =   224
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   360
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
            Index           =   56
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   215
            Top             =   765
            Width           =   4260
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
            Index           =   56
            Left            =   2385
            MaxLength       =   10
            TabIndex        =   225
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   765
            Width           =   1050
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
            Index           =   55
            Left            =   2385
            MaxLength       =   10
            TabIndex        =   226
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1170
            Width           =   1050
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
            Index           =   55
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   214
            Top             =   1170
            Width           =   4260
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   32
            Left            =   2115
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   1575
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
            Index           =   17
            Left            =   180
            TabIndex        =   220
            Top             =   1620
            Width           =   1935
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   17
            Left            =   2115
            Picture         =   "frmAPOListados.frx":9298
            ToolTipText     =   "Buscar fecha"
            Top             =   360
            Width           =   240
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
            Index           =   62
            Left            =   180
            TabIndex        =   219
            Top             =   405
            Width           =   1965
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
            Height          =   195
            Index           =   16
            Left            =   180
            TabIndex        =   218
            Top             =   810
            Width           =   1830
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   31
            Left            =   2115
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   765
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   30
            Left            =   2115
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   1170
            Width           =   240
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
            Height          =   195
            Index           =   15
            Left            =   180
            TabIndex        =   217
            Top             =   1215
            Width           =   2025
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
         TabIndex        =   212
         Top             =   780
         Width           =   7845
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
            Index           =   59
            Left            =   2385
            MaxLength       =   10
            TabIndex        =   222
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   360
            Width           =   1050
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
            Index           =   59
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   232
            Top             =   360
            Width           =   4260
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
            Index           =   54
            Left            =   2370
            MaxLength       =   10
            TabIndex        =   223
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   765
            Width           =   1350
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   33
            Left            =   2115
            ToolTipText     =   "Buscar socio"
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Socio"
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
            Index           =   18
            Left            =   180
            TabIndex        =   233
            Top             =   405
            Width           =   1515
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   18
            Left            =   2115
            Picture         =   "frmAPOListados.frx":9323
            ToolTipText     =   "Buscar fecha"
            Top             =   780
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Devolución"
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
            Index           =   61
            Left            =   180
            TabIndex        =   230
            Top             =   765
            Width           =   1830
         End
      End
      Begin VB.CommandButton CmdAcepRegBajaSocios 
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
         Left            =   5730
         TabIndex        =   228
         Top             =   4755
         Width           =   1065
      End
      Begin VB.CommandButton CmdCancel 
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
         Index           =   7
         Left            =   6900
         TabIndex        =   229
         Top             =   4755
         Width           =   1065
      End
      Begin MSComctlLib.ProgressBar Pb7 
         Height          =   255
         Left            =   210
         TabIndex        =   231
         Top             =   4320
         Visible         =   0   'False
         Width           =   7785
         _ExtentX        =   13732
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
         TabIndex        =   221
         Top             =   270
         Width           =   5160
      End
   End
   Begin VB.Frame FrameDevolAporQua 
      Height          =   7140
      Left            =   0
      TabIndex        =   384
      Top             =   0
      Width           =   8085
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
         Index           =   112
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   396
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   5280
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
         Index           =   111
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   394
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3960
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
         Index           =   110
         Left            =   1785
         MaxLength       =   3
         TabIndex        =   392
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2820
         Width           =   1050
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
         Index           =   109
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   391
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2460
         Width           =   1050
      End
      Begin VB.CommandButton CmdCancel 
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
         Index           =   13
         Left            =   6690
         TabIndex        =   400
         Top             =   6435
         Width           =   1065
      End
      Begin VB.CommandButton CmdAcepDevApoQua 
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
         Left            =   5520
         TabIndex        =   398
         Top             =   6450
         Width           =   1065
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
         Index           =   108
         Left            =   1785
         MaxLength       =   16
         TabIndex        =   390
         Top             =   1650
         Width           =   1035
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
         Index           =   107
         Left            =   1785
         MaxLength       =   16
         TabIndex        =   389
         Top             =   1260
         Width           =   1035
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
         Index           =   107
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   388
         Text            =   "Text5"
         Top             =   1275
         Width           =   4860
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
         Index           =   108
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   387
         Text            =   "Text5"
         Top             =   1650
         Width           =   4860
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
         Index           =   109
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   386
         Text            =   "Text5"
         Top             =   2460
         Width           =   4860
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
         Index           =   110
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   385
         Text            =   "Text5"
         Top             =   2835
         Width           =   4860
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
         Index           =   106
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   393
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3570
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
         Index           =   98
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   395
         Tag             =   "Código Postal|T|S|||clientes|codposta|0000||"
         Top             =   4560
         Width           =   1050
      End
      Begin MSComctlLib.ProgressBar pb12 
         Height          =   255
         Left            =   420
         TabIndex        =   397
         Top             =   6030
         Visible         =   0   'False
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
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
         Index           =   129
         Left            =   750
         TabIndex        =   411
         Top             =   3600
         Width           =   690
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
         Index           =   127
         Left            =   750
         TabIndex        =   410
         Top             =   3960
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Devolución"
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
         Index           =   120
         Left            =   495
         TabIndex        =   409
         Top             =   4980
         Width           =   1815
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   34
         Left            =   1530
         Picture         =   "frmAPOListados.frx":93AE
         ToolTipText     =   "Buscar fecha"
         Top             =   5280
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   33
         Left            =   1530
         Picture         =   "frmAPOListados.frx":9439
         ToolTipText     =   "Buscar fecha"
         Top             =   3960
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Aportación"
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
         Index           =   128
         Left            =   495
         TabIndex        =   408
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
         TabIndex        =   407
         Top             =   315
         Width           =   5160
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Index           =   126
         Left            =   495
         TabIndex        =   406
         Top             =   2145
         Width           =   1815
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
         Index           =   125
         Left            =   765
         TabIndex        =   405
         Top             =   2445
         Width           =   690
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
         Index           =   124
         Left            =   765
         TabIndex        =   404
         Top             =   2805
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
         Index           =   123
         Left            =   795
         TabIndex        =   403
         Top             =   1275
         Width           =   690
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
         Index           =   122
         Left            =   810
         TabIndex        =   402
         Top             =   1650
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
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
         Index           =   121
         Left            =   495
         TabIndex        =   401
         Top             =   1035
         Width           =   540
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   57
         Left            =   1500
         MouseIcon       =   "frmAPOListados.frx":94C4
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   56
         Left            =   1500
         MouseIcon       =   "frmAPOListados.frx":9616
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1260
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   55
         Left            =   1500
         MouseIcon       =   "frmAPOListados.frx":9768
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2820
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   54
         Left            =   1500
         MouseIcon       =   "frmAPOListados.frx":98BA
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2475
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   32
         Left            =   1530
         Picture         =   "frmAPOListados.frx":9A0C
         ToolTipText     =   "Buscar fecha"
         Top             =   3570
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Ejercicio"
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
         Index           =   119
         Left            =   495
         TabIndex        =   399
         Top             =   4560
         Width           =   1815
      End
   End
   Begin VB.Frame FrameListAporQua 
      Height          =   5850
      Left            =   30
      TabIndex        =   120
      Top             =   30
      Width           =   8085
      Begin VB.CheckBox Check2 
         Caption         =   "Salta página por socio"
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
         Height          =   240
         Left            =   5025
         TabIndex        =   189
         Top             =   4740
         Width           =   2670
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
         Left            =   750
         Style           =   2  'Dropdown List
         TabIndex        =   188
         Tag             =   "Recolectado|N|N|0|1|rcampos|recolect||N|"
         Top             =   4890
         Width           =   2370
      End
      Begin VB.Frame Frame4 
         Caption         =   "Tipo"
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
         Height          =   780
         Left            =   5055
         TabIndex        =   144
         Top             =   3360
         Width           =   2775
         Begin VB.OptionButton Opcion1 
            Caption         =   "Fecha"
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
            Index           =   1
            Left            =   1290
            TabIndex        =   146
            Top             =   300
            Width           =   930
         End
         Begin VB.OptionButton Opcion1 
            Caption         =   "Año"
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
            Index           =   0
            Left            =   300
            TabIndex        =   145
            Top             =   300
            Width           =   1290
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Imprimir Resumen"
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
         Height          =   240
         Left            =   5025
         TabIndex        =   143
         Top             =   4350
         Width           =   2490
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
         Index           =   41
         Left            =   1770
         MaxLength       =   10
         TabIndex        =   126
         Tag             =   "Código Postal|T|S|||clientes|codposta|0000||"
         Top             =   4080
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
         Index           =   39
         Left            =   1785
         MaxLength       =   10
         TabIndex        =   124
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2850
         Width           =   1050
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
         Index           =   38
         Left            =   1785
         MaxLength       =   10
         TabIndex        =   123
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2460
         Width           =   1050
      End
      Begin VB.CommandButton CmdCancel 
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
         Index           =   4
         Left            =   6705
         TabIndex        =   132
         Top             =   5265
         Width           =   1065
      End
      Begin VB.CommandButton CmdAcepListAporQua 
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
         Left            =   5535
         TabIndex        =   127
         Top             =   5265
         Width           =   1065
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
         Index           =   37
         Left            =   1785
         MaxLength       =   16
         TabIndex        =   122
         Top             =   1725
         Width           =   1035
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
         Index           =   36
         Left            =   1785
         MaxLength       =   16
         TabIndex        =   121
         Top             =   1320
         Width           =   1035
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
         Index           =   36
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   131
         Text            =   "Text5"
         Top             =   1335
         Width           =   4995
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
         Index           =   37
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   130
         Text            =   "Text5"
         Top             =   1710
         Width           =   4995
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
         Index           =   38
         Left            =   2850
         Locked          =   -1  'True
         TabIndex        =   129
         Text            =   "Text5"
         Top             =   2460
         Width           =   4995
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
         Index           =   39
         Left            =   2850
         Locked          =   -1  'True
         TabIndex        =   128
         Text            =   "Text5"
         Top             =   2835
         Width           =   4995
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
         Index           =   35
         Left            =   1770
         MaxLength       =   10
         TabIndex        =   125
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3690
         Width           =   1350
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   0
         Left            =   7740
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   4710
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Situación"
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
         Index           =   59
         Left            =   480
         TabIndex        =   187
         Top             =   4560
         Width           =   1185
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
         Index           =   50
         Left            =   750
         TabIndex        =   142
         Top             =   4080
         Width           =   690
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   11
         Left            =   1470
         Picture         =   "frmAPOListados.frx":9A97
         ToolTipText     =   "Buscar fecha"
         Top             =   4080
         Width           =   240
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
         Index           =   48
         Left            =   750
         TabIndex        =   141
         Top             =   3690
         Width           =   690
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Aportación"
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
         Index           =   49
         Left            =   480
         TabIndex        =   140
         Top             =   3345
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
         TabIndex        =   139
         Top             =   315
         Width           =   5160
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Index           =   47
         Left            =   465
         TabIndex        =   138
         Top             =   2100
         Width           =   1815
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
         Index           =   46
         Left            =   735
         TabIndex        =   137
         Top             =   2445
         Width           =   690
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
         Index           =   45
         Left            =   735
         TabIndex        =   136
         Top             =   2805
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
         Index           =   44
         Left            =   765
         TabIndex        =   135
         Top             =   1335
         Width           =   690
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
         Index           =   43
         Left            =   780
         TabIndex        =   134
         Top             =   1710
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
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
         Index           =   42
         Left            =   510
         TabIndex        =   133
         Top             =   1050
         Width           =   540
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   19
         Left            =   1470
         MouseIcon       =   "frmAPOListados.frx":9B22
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1725
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   16
         Left            =   1470
         MouseIcon       =   "frmAPOListados.frx":9C74
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   18
         Left            =   1470
         MouseIcon       =   "frmAPOListados.frx":9DC6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2850
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   17
         Left            =   1470
         MouseIcon       =   "frmAPOListados.frx":9F18
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2490
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   10
         Left            =   1470
         Picture         =   "frmAPOListados.frx":A06A
         ToolTipText     =   "Buscar fecha"
         Top             =   3690
         Width           =   240
      End
   End
   Begin VB.Frame FrameRegAltaSocios 
      Height          =   5400
      Left            =   0
      TabIndex        =   190
      Top             =   0
      Width           =   8130
      Begin VB.CommandButton CmdCancel 
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
         Index           =   6
         Left            =   6900
         TabIndex        =   206
         Top             =   4755
         Width           =   1065
      End
      Begin VB.CommandButton CmdAcepRegAltaSocios 
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
         Left            =   5730
         TabIndex        =   205
         Top             =   4755
         Width           =   1065
      End
      Begin VB.Frame Frame9 
         Caption         =   "Datos de Selección"
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
         Height          =   1005
         Left            =   120
         TabIndex        =   199
         Top             =   840
         Width           =   7845
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
            Index           =   60
            Left            =   1695
            MaxLength       =   10
            TabIndex        =   200
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   450
            Width           =   1275
         End
         Begin VB.Label Label4 
            Caption         =   "Precio Kilo"
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
            Index           =   68
            Left            =   195
            TabIndex        =   207
            Top             =   465
            Width           =   1815
         End
      End
      Begin VB.Frame Frame8 
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
         Height          =   2025
         Left            =   120
         TabIndex        =   191
         Top             =   1890
         Width           =   7845
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
            Index           =   53
            Left            =   3435
            Locked          =   -1  'True
            TabIndex        =   194
            Top             =   1170
            Width           =   4215
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
            Index           =   53
            Left            =   2475
            MaxLength       =   10
            TabIndex        =   203
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1170
            Width           =   945
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
            Index           =   52
            Left            =   2475
            MaxLength       =   10
            TabIndex        =   202
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   765
            Width           =   945
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
            Index           =   52
            Left            =   3435
            Locked          =   -1  'True
            TabIndex        =   193
            Top             =   765
            Width           =   4215
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
            Index           =   51
            Left            =   2475
            MaxLength       =   10
            TabIndex        =   201
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   360
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
            Index           =   50
            Left            =   3840
            Locked          =   -1  'True
            TabIndex        =   192
            Top             =   1575
            Width           =   3810
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
            Index           =   50
            Left            =   2475
            MaxLength       =   10
            TabIndex        =   204
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1575
            Width           =   1350
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
            Index           =   13
            Left            =   180
            TabIndex        =   198
            Top             =   1215
            Width           =   1755
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   29
            Left            =   2205
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   1170
            Width           =   420
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   28
            Left            =   2205
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   765
            Width           =   420
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
            Height          =   240
            Index           =   12
            Left            =   180
            TabIndex        =   197
            Top             =   810
            Width           =   1740
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
            Index           =   60
            Left            =   180
            TabIndex        =   196
            Top             =   405
            Width           =   1965
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   16
            Left            =   2205
            Picture         =   "frmAPOListados.frx":A0F5
            ToolTipText     =   "Buscar fecha"
            Top             =   360
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
            Index           =   11
            Left            =   180
            TabIndex        =   195
            Top             =   1620
            Width           =   2025
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   27
            Left            =   2205
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   1575
            Width           =   420
         End
      End
      Begin MSComctlLib.ProgressBar Pb6 
         Height          =   255
         Left            =   210
         TabIndex        =   208
         Top             =   4320
         Visible         =   0   'False
         Width           =   7785
         _ExtentX        =   13732
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
         TabIndex        =   210
         Top             =   270
         Width           =   5160
      End
      Begin VB.Label Label1 
         Caption         =   "lb1"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   14
         Left            =   270
         TabIndex        =   209
         Top             =   3990
         Visible         =   0   'False
         Width           =   7680
      End
   End
   Begin VB.Frame FrameCertificadoCPi 
      Height          =   7440
      Left            =   0
      TabIndex        =   415
      Top             =   0
      Width           =   8085
      Begin VB.TextBox txtcodigo 
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
         Index           =   115
         Left            =   1965
         MaxLength       =   40
         TabIndex        =   426
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   4770
         Width           =   5730
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   114
         Left            =   1965
         MaxLength       =   40
         TabIndex        =   428
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   5670
         Width           =   5730
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   113
         Left            =   1965
         MaxLength       =   40
         TabIndex        =   427
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   5220
         Width           =   5730
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
         Index           =   122
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   443
         Text            =   "Text5"
         Top             =   3645
         Width           =   4680
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
         Index           =   122
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   423
         Top             =   3645
         Width           =   1035
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
         Index           =   121
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   424
         Text            =   "Text5"
         Top             =   3270
         Width           =   4680
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
         Index           =   121
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   422
         Top             =   3270
         Width           =   1035
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
         Index           =   120
         Left            =   3030
         Locked          =   -1  'True
         TabIndex        =   420
         Text            =   "Text5"
         Top             =   1590
         Width           =   4680
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
         Index           =   119
         Left            =   3030
         Locked          =   -1  'True
         TabIndex        =   418
         Text            =   "Text5"
         Top             =   1215
         Width           =   4680
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
         Index           =   120
         Left            =   1935
         MaxLength       =   16
         TabIndex        =   417
         Top             =   1590
         Width           =   1035
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
         Index           =   119
         Left            =   1935
         MaxLength       =   16
         TabIndex        =   416
         Top             =   1215
         Width           =   1035
      End
      Begin VB.CommandButton CmdAcepCertCPi 
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
         Left            =   5430
         TabIndex        =   430
         Top             =   6720
         Width           =   1065
      End
      Begin VB.CommandButton CmdCancel 
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
         Index           =   14
         Left            =   6600
         TabIndex        =   431
         Top             =   6720
         Width           =   1065
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
         Index           =   118
         Left            =   1950
         MaxLength       =   10
         TabIndex        =   421
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2565
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
         Index           =   117
         Left            =   1935
         MaxLength       =   10
         TabIndex        =   419
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2160
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
         Index           =   116
         Left            =   1965
         MaxLength       =   10
         TabIndex        =   425
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   4335
         Width           =   1360
      End
      Begin VB.TextBox txtcodigo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   97
         Left            =   1965
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   429
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   6135
         Width           =   5715
      End
      Begin VB.Label Label4 
         Caption         =   "Presidente"
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
         Index           =   143
         Left            =   480
         TabIndex        =   447
         Top             =   4815
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Firmado ante "
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
         Index           =   133
         Left            =   480
         TabIndex        =   446
         Top             =   5715
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Secretario"
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
         Index           =   132
         Left            =   480
         TabIndex        =   445
         Top             =   5265
         Width           =   1815
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   60
         Left            =   1635
         MouseIcon       =   "frmAPOListados.frx":A180
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar aportacion"
         Top             =   3645
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
         Index           =   131
         Left            =   930
         TabIndex        =   444
         Top             =   3645
         Width           =   645
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   59
         Left            =   1635
         MouseIcon       =   "frmAPOListados.frx":A2D2
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar aportacion"
         Top             =   3270
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Aportación"
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
         Index           =   142
         Left            =   480
         TabIndex        =   442
         Top             =   2970
         Width           =   1560
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
         Index           =   141
         Left            =   930
         TabIndex        =   441
         Top             =   3270
         Width           =   645
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   58
         Left            =   1665
         MouseIcon       =   "frmAPOListados.frx":A424
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1230
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   53
         Left            =   1665
         MouseIcon       =   "frmAPOListados.frx":A576
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1590
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
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
         Index           =   140
         Left            =   480
         TabIndex        =   440
         Top             =   975
         Width           =   540
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
         Index           =   139
         Left            =   975
         TabIndex        =   439
         Top             =   1590
         Width           =   600
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
         Index           =   138
         Left            =   960
         TabIndex        =   438
         Top             =   1215
         Width           =   645
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   37
         Left            =   1665
         Picture         =   "frmAPOListados.frx":A6C8
         ToolTipText     =   "Buscar fecha"
         Top             =   2565
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   36
         Left            =   1665
         Picture         =   "frmAPOListados.frx":A753
         ToolTipText     =   "Buscar fecha"
         Top             =   2160
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
         Index           =   137
         Left            =   930
         TabIndex        =   437
         Top             =   2625
         Width           =   600
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
         Index           =   136
         Left            =   930
         TabIndex        =   436
         Top             =   2220
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
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
         Index           =   135
         Left            =   480
         TabIndex        =   435
         Top             =   1965
         Width           =   1815
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
         Index           =   29
         Left            =   480
         TabIndex        =   434
         Top             =   315
         Width           =   5160
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   35
         Left            =   1695
         Picture         =   "frmAPOListados.frx":A7DE
         ToolTipText     =   "Buscar fecha"
         Top             =   4335
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Certificado"
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
         Index           =   134
         Left            =   480
         TabIndex        =   433
         Top             =   4035
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Observaciones"
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
         Index           =   130
         Left            =   480
         TabIndex        =   432
         Top             =   6135
         Width           =   1815
      End
   End
   Begin VB.Frame FrameCertificadoBol 
      Height          =   7530
      Left            =   0
      TabIndex        =   325
      Top             =   0
      Width           =   8085
      Begin VB.TextBox txtcodigo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   95
         Left            =   1800
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   338
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   5640
         Width           =   5850
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   94
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   337
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   5220
         Width           =   5850
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   93
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   336
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   4830
         Width           =   5850
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   92
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   335
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   4440
         Width           =   5850
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
         Index           =   76
         Left            =   1830
         MaxLength       =   10
         TabIndex        =   334
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3930
         Width           =   1360
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
         Index           =   91
         Left            =   1815
         MaxLength       =   10
         TabIndex        =   332
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2625
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
         Index           =   90
         Left            =   1815
         MaxLength       =   10
         TabIndex        =   331
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2265
         Width           =   1350
      End
      Begin VB.CommandButton CmdCancel 
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
         Index           =   11
         Left            =   6555
         TabIndex        =   340
         Top             =   6855
         Width           =   1065
      End
      Begin VB.CommandButton CmdAcepCertBol 
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
         Left            =   5385
         TabIndex        =   339
         Top             =   6855
         Width           =   1065
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
         Index           =   89
         Left            =   1800
         MaxLength       =   16
         TabIndex        =   330
         Top             =   1590
         Width           =   1035
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
         Index           =   88
         Left            =   1800
         MaxLength       =   16
         TabIndex        =   329
         Top             =   1200
         Width           =   1035
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
         Index           =   88
         Left            =   2895
         Locked          =   -1  'True
         TabIndex        =   328
         Text            =   "Text5"
         Top             =   1215
         Width           =   4770
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
         Index           =   89
         Left            =   2895
         Locked          =   -1  'True
         TabIndex        =   327
         Text            =   "Text5"
         Top             =   1590
         Width           =   4770
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
         Index           =   87
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   333
         Top             =   3270
         Width           =   1035
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
         Index           =   87
         Left            =   2865
         Locked          =   -1  'True
         TabIndex        =   326
         Text            =   "Text5"
         Top             =   3270
         Width           =   4725
      End
      Begin VB.Label Label4 
         Caption         =   "Observaciones"
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
         Index           =   105
         Left            =   345
         TabIndex        =   354
         Top             =   5640
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Tesorero"
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
         Index           =   104
         Left            =   345
         TabIndex        =   353
         Top             =   5220
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Secretario"
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
         Index           =   103
         Left            =   345
         TabIndex        =   352
         Top             =   4830
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Presidente"
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
         Index           =   102
         Left            =   345
         TabIndex        =   351
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Certificado"
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
         Index           =   84
         Left            =   345
         TabIndex        =   350
         Top             =   3630
         Width           =   1815
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   28
         Left            =   1560
         Picture         =   "frmAPOListados.frx":A869
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
         Left            =   345
         TabIndex        =   349
         Top             =   315
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
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
         Index           =   101
         Left            =   345
         TabIndex        =   348
         Top             =   1965
         Width           =   1815
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
         Index           =   100
         Left            =   795
         TabIndex        =   347
         Top             =   2265
         Width           =   645
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
         Index           =   99
         Left            =   795
         TabIndex        =   346
         Top             =   2625
         Width           =   600
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   27
         Left            =   1530
         Picture         =   "frmAPOListados.frx":A8F4
         ToolTipText     =   "Buscar fecha"
         Top             =   2610
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   25
         Left            =   1530
         Picture         =   "frmAPOListados.frx":A97F
         ToolTipText     =   "Buscar fecha"
         Top             =   2250
         Width           =   240
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
         Index           =   98
         Left            =   825
         TabIndex        =   345
         Top             =   1215
         Width           =   645
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
         Index           =   97
         Left            =   840
         TabIndex        =   344
         Top             =   1590
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
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
         Index           =   96
         Left            =   345
         TabIndex        =   343
         Top             =   975
         Width           =   540
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   49
         Left            =   1530
         MouseIcon       =   "frmAPOListados.frx":AA0A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1590
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   48
         Left            =   1530
         MouseIcon       =   "frmAPOListados.frx":AB5C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1230
         Width           =   240
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
         Index           =   94
         Left            =   795
         TabIndex        =   342
         Top             =   3270
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Aportación"
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
         Index           =   83
         Left            =   345
         TabIndex        =   341
         Top             =   2970
         Width           =   1560
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   47
         Left            =   1500
         MouseIcon       =   "frmAPOListados.frx":ACAE
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar aportacion"
         Top             =   3270
         Width           =   240
      End
   End
   Begin VB.Frame FrameDevolAporBol 
      Height          =   6870
      Left            =   0
      TabIndex        =   355
      Top             =   0
      Width           =   8085
      Begin MSComctlLib.ProgressBar Pb11 
         Height          =   255
         Left            =   360
         TabIndex        =   383
         Top             =   5955
         Visible         =   0   'False
         Width           =   7290
         _ExtentX        =   12859
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
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
         Index           =   96
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   380
         Text            =   "Text5"
         Top             =   4215
         Width           =   4905
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
         Index           =   96
         Left            =   1710
         MaxLength       =   3
         TabIndex        =   361
         Top             =   4185
         Width           =   1035
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
         Index           =   105
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   368
         Text            =   "Text5"
         Top             =   3405
         Width           =   4905
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
         Index           =   105
         Left            =   1710
         MaxLength       =   3
         TabIndex        =   360
         Top             =   3405
         Width           =   1035
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
         Index           =   104
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   366
         Text            =   "Text5"
         Top             =   1635
         Width           =   4905
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
         Index           =   103
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   364
         Text            =   "Text5"
         Top             =   1215
         Width           =   4905
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
         Index           =   104
         Left            =   1710
         MaxLength       =   16
         TabIndex        =   357
         Top             =   1635
         Width           =   1035
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
         Index           =   103
         Left            =   1710
         MaxLength       =   16
         TabIndex        =   356
         Top             =   1200
         Width           =   1035
      End
      Begin VB.CommandButton CmdAcepDevApor 
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
         Left            =   5400
         TabIndex        =   365
         Top             =   6270
         Width           =   1065
      End
      Begin VB.CommandButton CmdCancel 
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
         Index           =   12
         Left            =   6570
         TabIndex        =   367
         Top             =   6270
         Width           =   1065
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
         Index           =   102
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   359
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2670
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
         Index           =   101
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   358
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2265
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
         Index           =   100
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   362
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   4935
         Width           =   1350
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   99
         Left            =   1710
         MaxLength       =   30
         TabIndex        =   363
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   5460
         Width           =   5940
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   52
         Left            =   1410
         MouseIcon       =   "frmAPOListados.frx":AE00
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar aportacion"
         Top             =   4185
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Tipo Aportación"
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
         Index           =   107
         Left            =   360
         TabIndex        =   382
         Top             =   3885
         Width           =   2250
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
         Index           =   106
         Left            =   660
         TabIndex        =   381
         Top             =   4185
         Width           =   735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   51
         Left            =   1410
         MouseIcon       =   "frmAPOListados.frx":AF52
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar aportacion"
         Top             =   3405
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Aportación"
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
         Index           =   118
         Left            =   360
         TabIndex        =   379
         Top             =   3105
         Width           =   1560
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
         Index           =   117
         Left            =   660
         TabIndex        =   378
         Top             =   3405
         Width           =   735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   50
         Left            =   1440
         MouseIcon       =   "frmAPOListados.frx":B0A4
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1605
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   46
         Left            =   1440
         MouseIcon       =   "frmAPOListados.frx":B1F6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1230
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
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
         Index           =   116
         Left            =   360
         TabIndex        =   377
         Top             =   975
         Width           =   540
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
         Index           =   115
         Left            =   705
         TabIndex        =   376
         Top             =   1635
         Width           =   690
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
         Index           =   114
         Left            =   690
         TabIndex        =   375
         Top             =   1215
         Width           =   735
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   31
         Left            =   1440
         Picture         =   "frmAPOListados.frx":B348
         ToolTipText     =   "Buscar fecha"
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   30
         Left            =   1440
         Picture         =   "frmAPOListados.frx":B3D3
         ToolTipText     =   "Buscar fecha"
         Top             =   2295
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
         Index           =   113
         Left            =   660
         TabIndex        =   374
         Top             =   2715
         Width           =   690
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
         Index           =   112
         Left            =   660
         TabIndex        =   373
         Top             =   2310
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
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
         Index           =   111
         Left            =   360
         TabIndex        =   372
         Top             =   2010
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
         TabIndex        =   371
         Top             =   315
         Width           =   5160
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   29
         Left            =   1410
         Picture         =   "frmAPOListados.frx":B45E
         ToolTipText     =   "Buscar fecha"
         Top             =   4935
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Devolución"
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
         Index           =   110
         Left            =   360
         TabIndex        =   370
         Top             =   4695
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Descripción"
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
         Index           =   109
         Left            =   360
         TabIndex        =   369
         Top             =   5340
         Width           =   1815
      End
   End
   Begin VB.Frame FrameCalculoAporQua 
      Height          =   7140
      Left            =   30
      TabIndex        =   90
      Top             =   -30
      Width           =   8085
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
         Index           =   32
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   118
         Text            =   "Text5"
         Top             =   1200
         Width           =   4905
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
         Index           =   32
         Left            =   1725
         MaxLength       =   16
         TabIndex        =   91
         Top             =   1200
         Width           =   1035
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
         Index           =   31
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   103
         Tag             =   "Código Postal|T|S|||clientes|codposta|0000||"
         Top             =   5400
         Width           =   1050
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
         Index           =   20
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   100
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   4470
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
         Index           =   28
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   115
         Text            =   "Text5"
         Top             =   3285
         Width           =   4905
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
         Index           =   27
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   114
         Text            =   "Text5"
         Top             =   2910
         Width           =   4905
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
         Index           =   30
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   96
         Text            =   "Text5"
         Top             =   2190
         Width           =   4905
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
         Index           =   29
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   94
         Text            =   "Text5"
         Top             =   1815
         Width           =   4905
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
         Index           =   30
         Left            =   1725
         MaxLength       =   16
         TabIndex        =   93
         Top             =   2190
         Width           =   1035
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
         Index           =   29
         Left            =   1725
         MaxLength       =   16
         TabIndex        =   92
         Top             =   1815
         Width           =   1035
      End
      Begin VB.CommandButton CmdAcepCalApoQua 
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
         Left            =   5385
         TabIndex        =   105
         Top             =   6450
         Width           =   1065
      End
      Begin VB.CommandButton CmdCancel 
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
         Index           =   3
         Left            =   6600
         TabIndex        =   107
         Top             =   6435
         Width           =   1065
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
         Index           =   28
         Left            =   1740
         MaxLength       =   3
         TabIndex        =   97
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3270
         Width           =   1050
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
         Index           =   27
         Left            =   1755
         MaxLength       =   3
         TabIndex        =   95
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2910
         Width           =   1050
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
         Index           =   26
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   99
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   3750
         Width           =   1050
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
         Index           =   25
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   102
         Tag             =   "Código Postal|T|S|||clientes|codposta|0000||"
         Top             =   4980
         Width           =   1050
      End
      Begin MSComctlLib.ProgressBar Pb5 
         Height          =   255
         Left            =   420
         TabIndex        =   98
         Top             =   6030
         Visible         =   0   'False
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   15
         Left            =   1440
         MouseIcon       =   "frmAPOListados.frx":B4E9
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar seccion"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   38
         Left            =   450
         TabIndex        =   119
         Top             =   960
         Width           =   780
      End
      Begin VB.Label Label4 
         Caption         =   "Ejercicio"
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
         Index           =   37
         Left            =   450
         TabIndex        =   117
         Top             =   5400
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Año"
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
         Index           =   28
         Left            =   450
         TabIndex        =   116
         Top             =   4980
         Width           =   1815
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   7
         Left            =   1470
         Picture         =   "frmAPOListados.frx":B63B
         ToolTipText     =   "Buscar fecha"
         Top             =   4470
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   14
         Left            =   1440
         MouseIcon       =   "frmAPOListados.frx":B6C6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   3285
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   1440
         MouseIcon       =   "frmAPOListados.frx":B818
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2910
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   1440
         MouseIcon       =   "frmAPOListados.frx":B96A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   2190
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   11
         Left            =   1440
         MouseIcon       =   "frmAPOListados.frx":BABC
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1815
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
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
         Index           =   36
         Left            =   450
         TabIndex        =   113
         Top             =   1575
         Width           =   540
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
         Index           =   35
         Left            =   750
         TabIndex        =   112
         Top             =   2190
         Width           =   600
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
         Index           =   34
         Left            =   735
         TabIndex        =   111
         Top             =   1815
         Width           =   645
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
         Index           =   33
         Left            =   705
         TabIndex        =   110
         Top             =   3255
         Width           =   600
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
         Index           =   32
         Left            =   705
         TabIndex        =   109
         Top             =   2895
         Width           =   645
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Index           =   31
         Left            =   450
         TabIndex        =   108
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
         Left            =   450
         TabIndex        =   106
         Top             =   315
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "Euros/Hda"
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
         Height          =   345
         Index           =   30
         Left            =   450
         TabIndex        =   104
         Top             =   3750
         Width           =   1305
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Aportación"
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
         Index           =   29
         Left            =   450
         TabIndex        =   101
         Top             =   4170
         Width           =   1815
      End
   End
   Begin VB.Frame FrameInforme 
      Height          =   6300
      Left            =   0
      TabIndex        =   64
      Top             =   45
      Width           =   8085
      Begin VB.CheckBox chkResumen 
         Caption         =   "Resumen"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4890
         TabIndex        =   72
         Tag             =   "Correo|N|N|||rsocios|correo||N|"
         Top             =   4380
         Width           =   1425
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
         Index           =   1
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Tag             =   "Tipo Relacion|N|N|0|2|rsocios|tiporelacion||N|"
         Top             =   4380
         Width           =   1590
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
         Index           =   19
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   86
         Text            =   "Text5"
         Top             =   3645
         Width           =   4950
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
         Index           =   13
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   85
         Text            =   "Text5"
         Top             =   3270
         Width           =   4950
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
         Index           =   19
         Left            =   1710
         MaxLength       =   16
         TabIndex        =   70
         Top             =   3645
         Width           =   1035
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
         Index           =   13
         Left            =   1710
         MaxLength       =   16
         TabIndex        =   69
         Top             =   3270
         Width           =   1035
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
         Index           =   24
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   76
         Text            =   "Text5"
         Top             =   1590
         Width           =   4950
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
         Index           =   23
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   75
         Text            =   "Text5"
         Top             =   1215
         Width           =   4950
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
         Index           =   24
         Left            =   1725
         MaxLength       =   16
         TabIndex        =   66
         Top             =   1590
         Width           =   1035
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
         Index           =   23
         Left            =   1725
         MaxLength       =   16
         TabIndex        =   65
         Top             =   1215
         Width           =   1035
      End
      Begin VB.CommandButton CmdAcepListado 
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
         Left            =   5430
         TabIndex        =   73
         Top             =   5535
         Width           =   1065
      End
      Begin VB.CommandButton CmdCancel 
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
         Index           =   2
         Left            =   6600
         TabIndex        =   74
         Top             =   5535
         Width           =   1065
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
         Index           =   21
         Left            =   1725
         MaxLength       =   10
         TabIndex        =   67
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2265
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
         Index           =   22
         Left            =   1725
         MaxLength       =   10
         TabIndex        =   68
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   2625
         Width           =   1350
      End
      Begin MSComctlLib.ProgressBar Pb3 
         Height          =   255
         Left            =   420
         TabIndex        =   77
         Top             =   5100
         Visible         =   0   'False
         Width           =   7290
         _ExtentX        =   12859
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Relación Cooperativa"
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
         Index           =   108
         Left            =   420
         TabIndex        =   413
         Top             =   4080
         Width           =   2070
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   1395
         MouseIcon       =   "frmAPOListados.frx":BC0E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar aportacion"
         Top             =   3645
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   1410
         MouseIcon       =   "frmAPOListados.frx":BD60
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar aportacion"
         Top             =   3270
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Aportación"
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
         Index           =   21
         Left            =   420
         TabIndex        =   89
         Top             =   2970
         Width           =   1560
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
         Index           =   20
         Left            =   720
         TabIndex        =   88
         Top             =   3645
         Width           =   690
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
         Index           =   19
         Left            =   705
         TabIndex        =   87
         Top             =   3270
         Width           =   735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   1425
         MouseIcon       =   "frmAPOListados.frx":BEB2
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1590
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1440
         MouseIcon       =   "frmAPOListados.frx":C004
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1215
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
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
         Index           =   27
         Left            =   420
         TabIndex        =   84
         Top             =   930
         Width           =   540
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
         Index           =   26
         Left            =   750
         TabIndex        =   83
         Top             =   1590
         Width           =   690
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
         Index           =   25
         Left            =   735
         TabIndex        =   82
         Top             =   1215
         Width           =   735
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   9
         Left            =   1455
         Picture         =   "frmAPOListados.frx":C156
         ToolTipText     =   "Buscar fecha"
         Top             =   2625
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   8
         Left            =   1455
         Picture         =   "frmAPOListados.frx":C1E1
         ToolTipText     =   "Buscar fecha"
         Top             =   2265
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
         Index           =   24
         Left            =   705
         TabIndex        =   81
         Top             =   2625
         Width           =   690
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
         Index           =   23
         Left            =   705
         TabIndex        =   80
         Top             =   2265
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
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
         Index           =   22
         Left            =   420
         TabIndex        =   79
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
         Left            =   450
         TabIndex        =   78
         Top             =   315
         Width           =   5160
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

Public Opcionlistado As Byte

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

'18= Certificado de aportaciones Coopic

Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmSoc As frmManSocios 'Socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCtas As frmCtasConta 'cuentas de contabilidad
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmFPa As frmForpaConta 'formas de pago de la contabilidad
Attribute frmFPa.VB_VarHelpID = -1
Private WithEvents frmApo As frmAPOTipos 'Tipo de Aportaciones
Attribute frmApo.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'para marcar que aportaciones queremos
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmSec As frmManSeccion 'para seleccionar
Attribute frmSec.VB_VarHelpID = -1
Private WithEvents frmVar As frmComVar 'Variedad
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmCla As frmBasico2 'Clase
Attribute frmCla.VB_VarHelpID = -1
Private WithEvents frmMens1 As frmMensajes 'para marcar que variedades queremos
Attribute frmMens1.VB_VarHelpID = -1
Private WithEvents frmMens2 As frmMensajes 'para marcar que variedades queremos en informe de aportaciones de quatretonda
Attribute frmMens2.VB_VarHelpID = -1


 
'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
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



Private Sub ChkDevolucion_Click()
    FrameDevolucion.Enabled = (ChkDevolucion.Value = 1)
    If Not FrameDevolucion.Enabled Then
        txtCodigo(123).Text = ""
'        txtCodigo(124).Text = ""
    End If
End Sub

Private Sub CmdAcepApoObli_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim i As Byte
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
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H socio
    cDesde = Trim(txtCodigo(77).Text)
    cHasta = Trim(txtCodigo(78).Text)
    nDesde = txtNombre(78).Text
    nHasta = txtNombre(78).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    Sql = "rsocios.fechabaja is null"
    If Not AnyadirAFormula(cadselect, Sql) Then Exit Sub
    
    
    If HayRegistros(tabla, cadselect) Then
        Sql2 = "select * from raportacion where (fecaport, codaport, codsocio) in (select " & DBSet(txtCodigo(74).Text, "F") & "," & DBSet(txtCodigo(71).Text, "N") & ", codsocio from "
        Sql2 = Sql2 & tabla
        If cadselect <> "" Then Sql2 = Sql2 & " where " & cadselect & ")"
        
        If TotalRegistros(Sql2) <> 0 Then
            If MsgBox("Existen aportaciones para algún socio/s de este tipo para esta fecha. " & vbCrLf & vbCrLf & " ¿ Desea continuar ? " & vbCrLf, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                Exit Sub
            End If
        End If
        If InsertarTemporal2(tabla, cadselect) Then
            indRPT = 83 ' "rManAportacion.rpt"
            
            cadTitulo = "Aportación Obligatoria"
        
            If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
            
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            
            cadParam = cadParam & "pTitulo=""Aportación Obligatoria""|"
            numParam = numParam + 1
            
            cadNombreRPT = nomDocu
            LlamarImprimir
            If MsgBox(" ¿ Desea continuar con el proceso ? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
    
                If InsertarAportacionesObligatoriasBolbaite(tabla, cadselect) Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                    cmdCancel_Click (0)
                End If
                
            End If
        End If
    End If
        

End Sub

Private Sub CmdAcepCertCPi_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim vSQL As String

    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H socio
    cDesde = Trim(txtCodigo(119).Text)
    cHasta = Trim(txtCodigo(120).Text)
    nDesde = txtNombre(119).Text
    nHasta = txtNombre(120).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H Fecha
    cDesde = Trim(txtCodigo(117).Text)
    cHasta = Trim(txtCodigo(118).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.fecaport}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    'Tipo de Aportacion
    cDesde = Trim(txtCodigo(121).Text)
    cHasta = Trim(txtCodigo(122).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.codaport}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHTipo= """) Then Exit Sub
    End If
    
    
    tabla = "raportacion INNER JOIN rsocios ON raportacion.codsocio = rsocios.codsocio"
    
    If HayRegistros(tabla, cadselect) Then
        cadParam = cadParam & "pObser=""" & txtCodigo(97).Text & """|"
        cadParam = cadParam & "pFecha=""" & txtCodigo(116).Text & """|"
        cadParam = cadParam & "pDesdeFecha=""" & txtCodigo(117).Text & """|"
        cadParam = cadParam & "pHastaFecha=""" & txtCodigo(118).Text & """|"
        cadParam = cadParam & "pSecretario=""" & txtCodigo(113).Text & """|"
        cadParam = cadParam & "pFirmado=""" & txtCodigo(114).Text & """|"
        cadParam = cadParam & "pPresidente=""" & txtCodigo(115).Text & """|"
        
        numParam = numParam + 6
        
        indRPT = 74 ' "rManAportacion.rpt"
        
        cadTitulo = "Certificado de Aportaciones"
    
        If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
        
        cadNombreRPT = nomDocu
        LlamarImprimir
    
    End If
End Sub


Private Sub CmdAcepCertBol_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim vSQL As String

    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H socio
    cDesde = Trim(txtCodigo(88).Text)
    cHasta = Trim(txtCodigo(89).Text)
    nDesde = txtNombre(88).Text
    nHasta = txtNombre(89).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H Fecha
    cDesde = Trim(txtCodigo(90).Text)
    cHasta = Trim(txtCodigo(91).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.fecaport}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    'Tipo de Aportacion
    If Not AnyadirAFormula(cadFormula, "{raportacion.codaport} = " & DBSet(txtCodigo(87).Text, "N")) Then Exit Sub
    If Not AnyadirAFormula(cadselect, "{raportacion.codaport} = " & DBSet(txtCodigo(87).Text, "N")) Then Exit Sub
    
    tabla = "raportacion INNER JOIN rsocios ON raportacion.codsocio = rsocios.codsocio"
    
    If HayRegistros(tabla, cadselect) Then
        cadParam = cadParam & "pPresi=""" & txtCodigo(92).Text & """|"
        cadParam = cadParam & "pSecre=""" & txtCodigo(93).Text & """|"
        cadParam = cadParam & "pTesor=""" & txtCodigo(94).Text & """|"
        cadParam = cadParam & "pObser=""" & txtCodigo(95).Text & """|"
        cadParam = cadParam & "pFecha=""" & txtCodigo(76).Text & """|"
        cadParam = cadParam & "pHastaFecha=""" & txtCodigo(91).Text & """|"
        numParam = numParam + 6
        
        indRPT = 74 ' "rManAportacion.rpt"
        
        cadTitulo = "Certificado de Aportaciones"
    
        If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
        
        cadNombreRPT = nomDocu
        LlamarImprimir
        If MsgBox(" ¿ Impresión correcta para actualizar ? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            If ActualizarTipo(tabla, cadselect) Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
            End If
        End If
    End If
End Sub

Private Function ActualizarTipo(tabla As String, cadselect As String) As Boolean
Dim Sql As String
Dim Nregs As Long

    On Error GoTo eActualizarTipo

    ActualizarTipo = False

    Sql = "select distinct rsocios.codsocio from " & tabla
    Sql = Sql & " where " & cadselect
    
    Nregs = TotalRegistrosConsulta(Sql)
    
    Sql = "update rtipoapor set numero = numero + " & DBSet(Nregs, "N")
    Sql = Sql & " where codaport = " & DBSet(txtCodigo(87).Text, "N")
    
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
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim vSQL As String

    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'Seleccionar registros de la empresa conectada
    'D/H socio
    cDesde = Trim(txtCodigo(107).Text)
    cHasta = Trim(txtCodigo(108).Text)
    nDesde = txtNombre(107).Text
    nHasta = txtNombre(108).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raporhco.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H clase
    cDesde = Trim(txtCodigo(109).Text)
    cHasta = Trim(txtCodigo(110).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{variedades.codclase}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHClase= """) Then Exit Sub
    End If
    
    vSQL = ""
    If txtCodigo(109).Text <> "" Then vSQL = vSQL & " and clases.codclase >= " & DBSet(txtCodigo(109).Text, "N")
    If txtCodigo(110).Text <> "" Then vSQL = vSQL & " and clases.codclase <= " & DBSet(txtCodigo(110).Text, "N")
    
                
    Set frmMens2 = New frmMensajes
    
    frmMens2.OpcionMensaje = 16
    frmMens2.cadWhere = vSQL
    frmMens2.Show vbModal
    
    Set frmMens2 = Nothing
    
    
    'D/H Fecha
    cDesde = Trim(txtCodigo(106).Text)
    cHasta = Trim(txtCodigo(111).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raporhco.fecaport}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    'Ejercicio
    If Not AnyadirAFormula(cadFormula, "{raporhco.ejercicio} = " & DBSet(txtCodigo(98).Text, "N")) Then Exit Sub
    If Not AnyadirAFormula(cadselect, "{raporhco.ejercicio} = " & DBSet(txtCodigo(98).Text, "N")) Then Exit Sub
    
    
    tabla = "raporhco INNER JOIN variedades ON raporhco.codvarie = variedades.codvarie "
    
    
    If HayRegistros(tabla, cadselect) Then
            indRPT = 83 ' "rManAportacion.rpt"
            
            cadTitulo = "Resumen Devolución de Aportaciones"
            
            If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
            
            cadNombreRPT = Replace(nomDocu, "APOInf.rpt", "APOInfAnyo.rpt")
            
            cadParam = cadParam & "pResumen=1|"
            numParam = numParam + 1

'            cadNombreRPT = nomDocu
            LlamarImprimir
            If MsgBox(" ¿ Desea continuar con el proceso ? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                If InsertarDevolucionesQua(tabla, cadselect) Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                End If
            End If
    End If


End Sub

Private Sub CmdAcepDevApor_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim vSQL As String

    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H socio
    cDesde = Trim(txtCodigo(103).Text)
    cHasta = Trim(txtCodigo(104).Text)
    nDesde = txtNombre(103).Text
    nHasta = txtNombre(104).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H Fecha
    cDesde = Trim(txtCodigo(101).Text)
    cHasta = Trim(txtCodigo(102).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.fecaport}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    'Tipo de Aportacion
    If Not AnyadirAFormula(cadFormula, "{raportacion.codaport} = " & DBSet(txtCodigo(105).Text, "N")) Then Exit Sub
    If Not AnyadirAFormula(cadselect, "{raportacion.codaport} = " & DBSet(txtCodigo(105).Text, "N")) Then Exit Sub
    
    
    
    'DAVID Agosto 2014
    'QUITAMOS del join   and rsocios.fechabaja is null
    tabla = "raportacion INNER JOIN rsocios ON raportacion.codsocio = rsocios.codsocio " 'and rsocios.fechabaja is null "
    
    If HayRegistros(tabla, cadselect) Then
        If InsertarTemporal(tabla, cadselect) Then
            indRPT = 83 ' "rManAportacion.rpt"
            
            cadTitulo = "Devolución de Aportaciones"
            
            cadParam = cadParam & "pTitulo=""Devolución de Aportaciones""|"
            numParam = numParam + 1
        
            If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
            
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            
            cadNombreRPT = nomDocu
            LlamarImprimir
            If MsgBox(" ¿ Desea continuar con el proceso ? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                If ActualizarDevoluciones(tabla, cadselect) Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                End If
            End If
        End If
    End If

End Sub

Private Function ActualizarDevoluciones(vtabla As String, vSelect As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
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
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Sql2 = "insert into raportacion (codsocio,fecaport,codaport,descripcion,campanya,kilos,importe,codtipom,numfactu,intconta,porcaport) values "


    b = True

    Pb11.visible = True
    Pb11.Max = TotalRegistrosConsulta(Sql)
    Pb11.Value = 0
    
    SqlValues = ""
    
    While Not Rs.EOF And b
        IncrementarProgresNew Pb11, 1
    
        SqlExiste = "select count(*) from raportacion where codsocio = " & DBSet(Rs!Codigo1, "N") & " and codaport=" & DBSet(txtCodigo(96).Text, "N") & " and fecaport=" & DBSet(txtCodigo(100).Text, "F") & " and numfactu = 0"
        b = (TotalRegistros(SqlExiste) = 0)
        
        If Not b Then
            MsgBox "El registro para el socio " & Format(DBLet(Rs!Codigo1, "N"), "000000") & " de fecha " & txtCodigo(100).Text & " y tipo " & DBSet(txtCodigo(96).Text, "N") & " existe. Revise.", vbExclamation
        Else
            Importe = DBLet(Rs!Importe, "N") * (-1)
        
            SqlValues = "(" & DBSet(Rs!Codigo1, "N") & "," & DBSet(txtCodigo(100).Text, "F") & "," & DBSet(txtCodigo(96).Text, "N") & "," & DBSet(txtCodigo(99).Text, "T") & ",'',0,"
            SqlValues = SqlValues & DBSet(Importe, "N") & "," & ValorNulo & ",0,0,0)"
            
            conn.Execute Sql2 & SqlValues
        End If
    
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
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

Private Function InsertarTemporal(vtabla As String, vSelect As String) As Boolean
Dim Sql As String

    On Error GoTo eInsertarTemporal
    
    InsertarTemporal = False
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
                                            'socio, fecaport,codaport,numfactu, codtipom, importe
    Sql = "insert into tmpinformes (codusu, codigo1, fecha1, campo1, importe1, nombre1, importe2)"
    Sql = Sql & " select " & vUsu.Codigo & ", raportacion.codsocio, fecaport, codaport, numfactu, codtipom, importe "
    Sql = Sql & " from " & vtabla
    Sql = Sql & " where " & vSelect
    
    conn.Execute Sql

    InsertarTemporal = True
    Exit Function

eInsertarTemporal:
    MuestraError Err.Number, "Insertar Temporal", Err.Description
End Function



Private Function InsertarTemporal2(vtabla As String, vSelect As String) As Boolean
Dim Sql As String

    On Error GoTo eInsertarTemporal
    
    InsertarTemporal2 = False
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
                                            'socio, fecaport,codaport,numfactu, codtipom, importe
    Sql = "insert into tmpinformes (codusu, codigo1, fecha1, campo1, importe1, nombre1, importe2)"
    Sql = Sql & " select " & vUsu.Codigo & ", codsocio, " & DBSet(txtCodigo(74).Text, "F") & "," & DBSet(txtCodigo(71).Text, "N") & ", 0, null," & DBSet(txtCodigo(73).Text, "N")
    Sql = Sql & " from " & vtabla
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
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim Tipos As String
Dim Sql As String

    
    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'Tipo de movimiento:
    Tipos = ""
    For i = 1 To ListView1(0).ListItems.Count
        If ListView1(0).ListItems(i).Checked Then
            Tipos = Tipos & DBSet(ListView1(0).ListItems(i).Key, "T") & ","
        End If
    Next i
    
    If Tipos = "" Then
        MsgBox "Debe seleccionar al menos un tipo de factura.", vbExclamation
        Exit Sub
    Else
        ' quitamos la ultima coma
        Tipos = "{" & tabla & ".codtipom} in (" & Mid(Tipos, 1, Len(Tipos) - 1) & ")"
        If Not AnyadirAFormula(cadselect, Tipos) Then Exit Sub
        Tipos = Replace(Replace(Tipos, "(", "["), ")", "]")
        If Not AnyadirAFormula(cadFormula, Tipos) Then Exit Sub
    End If
    
    'D/H socio
    cDesde = Trim(txtCodigo(66).Text)
    cHasta = Trim(txtCodigo(67).Text)
    nDesde = txtNombre(66).Text
    nHasta = txtNombre(67).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H Nro de Factura
    cDesde = Trim(txtCodigo(61).Text)
    cHasta = Trim(txtCodigo(62).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFactura= """) Then Exit Sub
    End If
    
    
    Select Case Opcionlistado
    Case 11 'Insercion de aportaciones
        
        'D/H Fecha factura
        cDesde = Trim(txtCodigo(64).Text)
        cHasta = Trim(txtCodigo(65).Text)
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & tabla & ".fecfactu}"
            TipCod = "F"
            If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
        End If
        
        
        Sql = " not (rfactsoc.codtipom, rfactsoc.fecfactu, rfactsoc.numfactu) in (select codtipom, fecaport, numfactu from raportacion) "
        If Not AnyadirAFormula(cadselect, Sql) Then Exit Sub
        
        If HayRegistros(tabla, cadselect) Then
            If InsertarAportacionesBolbaite(tabla, cadselect) Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                cmdCancel_Click (0)
            End If
        End If
        
    Case 12 'Impresion de recibos
        'D/H Fecha factura
        cDesde = Trim(txtCodigo(64).Text)
        cHasta = Trim(txtCodigo(65).Text)
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & tabla & ".fecaport}"
            TipCod = "F"
            If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
        End If
        
        cadParam = cadParam & "pFecha=""" & txtCodigo(70).Text & """|"
        numParam = numParam + 1
        
        If HayRegistros(tabla, cadselect) Then
            indRPT = 100 'Impresion de Recibos de aportaciones
            ConSubInforme = True
            
            cadTitulo = "Impresión de Recibos Aportaciones"
            If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
            
            cadNombreRPT = nomDocu
              
              
            LlamarImprimir
        End If
    End Select

End Sub

Private Sub CmdAcepCalApoQua_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim vSQL As String

InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
     '======== FORMULA  ====================================
    'Seleccionar registros de la empresa conectada
    'SECCION
    Codigo = "{rsocios_seccion.codsecci}=" & txtCodigo(32).Text
    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
    If Not AnyadirAFormula(cadselect, Codigo) Then Exit Sub
    
    'D/H socio
    cDesde = Trim(txtCodigo(29).Text)
    cHasta = Trim(txtCodigo(30).Text)
    nDesde = txtNombre(29).Text
    nHasta = txtNombre(30).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rcampos.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'hasta el año de plantacion
    Codigo = "{rcampos.anoplant}<=" & txtCodigo(25).Text
    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
    If Not AnyadirAFormula(cadselect, Codigo) Then Exit Sub
    
    
    'D/H clase
    cDesde = Trim(txtCodigo(27).Text)
    cHasta = Trim(txtCodigo(28).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{variedades.codclase}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHClase= """) Then Exit Sub
    End If
    
    vSQL = ""
    If txtCodigo(27).Text <> "" Then vSQL = vSQL & " and clases.codclase >= " & DBSet(txtCodigo(27).Text, "N")
    If txtCodigo(28).Text <> "" Then vSQL = vSQL & " and clases.codclase <= " & DBSet(txtCodigo(28).Text, "N")
    
                
    Set frmMens1 = New frmMensajes
    
    frmMens1.OpcionMensaje = 16
    frmMens1.cadWhere = vSQL
    frmMens1.Show vbModal
    
    Set frmMens1 = Nothing
    
    
    tabla = "((rsocios INNER JOIN rcampos ON rsocios.codsocio = rcampos.codsocio and rcampos.fecbajas is null and rsocios.fechabaja is null) "
    tabla = tabla & " INNER JOIN rsocios_seccion ON rcampos.codsocio = rsocios_seccion.codsocio and rsocios_seccion.fecbaja is null) "
    tabla = tabla & " INNER JOIN variedades ON rcampos.codvarie = variedades.codvarie "
    
    If HayRegistros(tabla, cadselect) Then
        If CalculoAportacionQuatretonda(tabla, cadselect) Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
            cmdCancel_Click (3)
        End If
    End If

End Sub

Private Function CalculoAportacionQuatretonda(vtabla As String, vWhere As String) As Boolean
Dim Sql As String
Dim Importe As Currency
Dim Rs As ADODB.Recordset
Dim cadErr As String
Dim NumApor As Long
Dim vTipoMov As CTiposMov
Dim b As Boolean
Dim SQLinsert As String
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
    SQLinsert = "insert into ariagro.raporhco (numaport,codsocio,codcampo,poligono,parcela,codparti,codvarie,impaport," & _
                "fecaport,anoplant,observac,supcoope,ejercicio,intconta) values "
    
    Sql = "select rcampos.* from " & vtabla
    Sql = Sql & " where " & vWhere
    
    CargarProgres Pb5, TotalRegistrosConsulta(Sql)
    Pb5.visible = True
    
    
    CadValues = ""
    CodTipoMov = "APO"
    
    Set vTipoMov = New CTiposMov
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF And b
        Sql2 = "select count(*) from ariagro.raporhco where codsocio = " & DBSet(Rs!Codsocio, "N")
        Sql2 = Sql2 & " and codcampo = " & DBSet(Rs!codCampo, "N") & " and codvarie = " & DBSet(Rs!codvarie, "N")
        Sql2 = Sql2 & " and fecaport = " & DBSet(txtCodigo(20).Text, "F")
        
        IncrementarProgres Pb5, 1
        DoEvents
        
        
        If TotalRegistros(Sql2) > 0 Then
            b = False
            cadErr = "Ya existe la aportación para el socio " & DBLet(Rs!Codsocio, "N") & ", campo " & _
                    DBLet(Rs!codCampo, "N") & ", variedad " & DBLet(Rs!codvarie, "N") & _
                    " y fecha de aportación " & txtCodigo(20).Text & ". Revise."
        Else
            Importe = Round2(Round2(DBLet(Rs!supcoope, "N") / vParamAplic.Faneca, 2) * CCur(ImporteSinFormato(txtCodigo(26).Text)), 2)
        
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
                
                CadValues = "(" & DBSet(NumApor, "N") & "," & DBSet(Rs!Codsocio, "N") & "," & DBSet(Rs!codCampo, "N") & ","
                CadValues = CadValues & DBSet(Rs!Poligono, "N") & "," & DBSet(Rs!Parcela, "N") & "," & DBSet(Rs!codparti, "N") & ","
                CadValues = CadValues & DBSet(Rs!codvarie, "N") & "," & DBSet(Importe, "N") & "," & DBSet(txtCodigo(20).Text, "F") & ","
                CadValues = CadValues & DBSet(Rs!anoplant, "N") & "," & ValorNulo & "," & DBSet(Rs!supcoope, "N") & ","
                CadValues = CadValues & DBSet(txtCodigo(31).Text, "N") & ",0)"
                
                conn.Execute SQLinsert & CadValues
                
                b = vTipoMov.IncrementarContador(CodTipoMov)
           End If
        End If
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
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
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim b As Boolean
Dim vSQL As String

InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    'D/H socio
    cDesde = Trim(txtCodigo(81).Text)
    cHasta = Trim(txtCodigo(82).Text)
    nDesde = txtNombre(81).Text
    nHasta = txtNombre(82).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H Fecha aportacion
    cDesde = Trim(txtCodigo(79).Text)
    cHasta = Trim(txtCodigo(80).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.fecaport}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    ' del tipo de aportacion
    If Not AnyadirAFormula(cadFormula, "{raportacion.codaport} = " & DBSet(txtCodigo(75).Text, "N")) Then Exit Sub
    If Not AnyadirAFormula(cadselect, "{raportacion.codaport} = " & DBSet(txtCodigo(75).Text, "N")) Then Exit Sub
    
    ' Condicion de que no esten contabilizados
    If Not AnyadirAFormula(cadFormula, "{raportacion.intconta} = 0") Then Exit Sub
    If Not AnyadirAFormula(cadselect, "{raportacion.intconta} = 0") Then Exit Sub
    
    tabla = "raportacion"
    
    If Not HayRegParaInforme(tabla, cadselect) Then Exit Sub
        
    If CargarTemporalBol(tabla, cadselect) Then
    
        TerminaBloquear
        
        tabla = tabla & " INNER JOIN tmpinformes ON raportacion.codsocio = tmpinformes.codigo1 and tmpinformes.codusu = " & vUsu.Codigo
        tabla = tabla & " and raportacion.fecaport = tmpinformes.fecha1 and raportacion.numfactu = tmpinformes.importe1 and (raportacion.codtipom = tmpinformes.nombre1 or raportacion.codtipom is null) "
        
        If Not BloqueaRegistro(tabla, cadselect) Then
            MsgBox "No se pueden Integrar en Tesoreria Aportaciones. Hay registros bloqueados.", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        ' comprobacion de que todos los socios tienen que estar en la seccion de almazara
        b = SociosEnSeccion("tmpinformes", "codusu = " & vUsu.Codigo, vParamAplic.Seccionhorto)
        If b Then b = ComprobarCtaContable_new("tmpinformes", 2, vParamAplic.Seccionhorto)
    
        If b Then
            If IntegracionAportacionesTesoreriaBolbaite(tabla, cadselect) Then
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
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim b As Boolean
Dim vSQL As String

InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    'D/H socio
    cDesde = Trim(txtCodigo(44).Text)
    cHasta = Trim(txtCodigo(45).Text)
    nDesde = txtNombre(44).Text
    nHasta = txtNombre(45).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raporhco.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H Fecha aportacion
    cDesde = Trim(txtCodigo(46).Text)
    cHasta = Trim(txtCodigo(47).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raporhco.fecaport}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    'D/H Clase
    cDesde = Trim(txtCodigo(43).Text)
    cHasta = Trim(txtCodigo(48).Text)
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
    If Not AnyadirAFormula(cadselect, "{raporhco.intconta} = 0") Then Exit Sub
    
    vSQL = ""
    If txtCodigo(43).Text <> "" Then vSQL = vSQL & " and clases.codclase >= " & DBSet(txtCodigo(43).Text, "N")
    If txtCodigo(48).Text <> "" Then vSQL = vSQL & " and clases.codclase <= " & DBSet(txtCodigo(48).Text, "N")
    
                
    Set frmMens2 = New frmMensajes
    
    frmMens2.OpcionMensaje = 16
    frmMens2.cadWhere = vSQL
    frmMens2.Show vbModal
    
    Set frmMens2 = Nothing
    
    
    tabla = "raporhco INNER JOIN variedades ON raporhco.codvarie = variedades.codvarie "

    If Not HayRegParaInforme(tabla, cadselect) Then Exit Sub
        
    If CargarTemporalQua(tabla, cadselect) Then
    
        'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
    '    TerminaBloquear
        tabla = "(" & tabla & ") INNER JOIN tmpinformes ON raporhco.numaport = tmpinformes.importe1 and tmpinformes.codusu = " & vUsu.Codigo
        If Not BloqueaRegistro(tabla, cadselect) Then
            MsgBox "No se pueden Integrar en Tesoreria Aportaciones. Hay registros bloqueados.", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        ' Comprobaciones
        ' comprobacion de que todos los socios tienen que estar en la seccion de almazara
        b = SociosEnSeccion("tmpinformes", "codusu = " & vUsu.Codigo, vParamAplic.Seccionhorto)
        If b Then b = ComprobarCtaContable_new("tmpinformes", 1, vParamAplic.Seccionhorto)
    
        If b Then
            If IntegracionAportacionesTesoreria(tabla, cadselect) Then
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
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim vSQL As String

InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H socio
    cDesde = Trim(txtCodigo(23).Text)
    cHasta = Trim(txtCodigo(24).Text)
    nDesde = txtNombre(23).Text
    nHasta = txtNombre(24).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H Fecha
    cDesde = Trim(txtCodigo(21).Text)
    cHasta = Trim(txtCodigo(22).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.fecaport}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    'D/H Tipo de Aportacion
    cDesde = Trim(txtCodigo(13).Text)
    cHasta = Trim(txtCodigo(19).Text)
    nDesde = txtNombre(13).Text
    nHasta = txtNombre(19).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.codaport}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHAportacion= """) Then Exit Sub
    End If
    
    vSQL = ""
    If txtCodigo(13).Text <> "" Then vSQL = vSQL & " and rtipoapor.codaport >= " & DBSet(txtCodigo(13).Text, "N")
    If txtCodigo(19).Text <> "" Then vSQL = vSQL & " and rtipoapor.codaport <= " & DBSet(txtCodigo(19).Text, "N")
    
                
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
            If Not AnyadirAFormula(cadselect, "rsocios.tiporelacion=" & Combo1(1).ListIndex - 1) Then Exit Sub
            cadParam = cadParam & "pRelacion=" & Combo1(1).ListIndex & "|"
            numParam = numParam + 1
    End Select
    
    cadParam = cadParam & "pResumen=" & chkResumen(0).Value & "|"
    numParam = numParam + 1
    
    tabla = "raportacion INNER JOIN rsocios ON raportacion.codsocio = rsocios.codsocio and rsocios.fechabaja is null "
    
    If HayRegistros(tabla, cadselect) Then
        indRPT = 101 ' "rManAportacion.rpt"
        
        cadTitulo = "Informe Aportaciones"
    
        If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
        
        cadNombreRPT = nomDocu
        LlamarImprimir
    
    End If

End Sub

Private Sub CmdAcepListAporQua_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim vSQL As String

InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H socio
    cDesde = Trim(txtCodigo(36).Text)
    cHasta = Trim(txtCodigo(37).Text)
    nDesde = txtNombre(36).Text
    nHasta = txtNombre(37).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raporhco.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H Fecha
    cDesde = Trim(txtCodigo(35).Text)
    cHasta = Trim(txtCodigo(41).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raporhco.fecaport}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    If Opcionlistado = 6 Then
        'D/H Clase
        cDesde = Trim(txtCodigo(38).Text)
        cHasta = Trim(txtCodigo(39).Text)
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
    If txtCodigo(38).Text <> "" Then vSQL = vSQL & " and variedades.codclase >= " & DBSet(txtCodigo(38).Text, "N")
    If txtCodigo(39).Text <> "" Then vSQL = vSQL & " and variedades.codclase <= " & DBSet(txtCodigo(39).Text, "N")
    
                
    Set frmMens2 = New frmMensajes
    
    frmMens2.OpcionMensaje = 16
    frmMens2.cadWhere = vSQL
    frmMens2.Show vbModal
    
    Set frmMens2 = Nothing
    
    
    If Opcionlistado = 6 Then ' solo en el caso del listado
        Select Case Combo1(0).ListIndex
            Case 0
                ' Condicion de que no esten contabilizados
                If Not AnyadirAFormula(cadFormula, "{raporhco.intconta} = 0") Then Exit Sub
                If Not AnyadirAFormula(cadselect, "{raporhco.intconta} = 0") Then Exit Sub
            Case 1
                ' Condicion de que esten contabilizados
                If Not AnyadirAFormula(cadFormula, "{raporhco.intconta} = 1") Then Exit Sub
                If Not AnyadirAFormula(cadselect, "{raporhco.intconta} = 1") Then Exit Sub
            Case 2
            
        End Select
    End If
    
    tabla = "(raporhco INNER JOIN variedades ON raporhco.codvarie = variedades.codvarie) "
    
    If HayRegistros(tabla, cadselect) Then
        Select Case Opcionlistado
            Case 6
                indRPT = 83 'informe de APORTACIONES para Quatretonda
            
                If Not PonerParamRPT(indRPT, cadParam, 1, nomDocu) Then Exit Sub
                                   
                cadNombreRPT = nomDocu '"rAPOInf.rpt"
                
                cadTitulo = "Informe Aportaciones"
                
                
                '[Monica]24/01/2012: salta página por socio, nuevo report
                If Check2.Value Then
                    cadNombreRPT = Replace(cadNombreRPT, "APOInf.rpt", "APOInfSocio.rpt")
                    cadTitulo = cadTitulo & " por Socio "
                    '[Monica]18/09/2014: añado lo del resumen cuando es por socio por las devoluciones
                    cadParam = cadParam & "pResumen=" & Me.Check1.Value & "|"
                    numParam = numParam + 1
                Else
                    If Me.Opcion1(0).Value Then
                        cadNombreRPT = Replace(cadNombreRPT, "APOInf.rpt", "APOInfAnyo.rpt")
                        cadTitulo = cadTitulo & " por Año "
                        
                        cadParam = cadParam & "pResumen=" & Me.Check1.Value & "|"
                        numParam = numParam + 1
                    End If
                End If
                
                
                frmImprimir.NombreRPT = cadNombreRPT
                cadTitulo = "Informe Aportaciones"
                LlamarImprimir
            
            Case 7 ' borrado masivo de aportaciones
                If BorradoMasivoAporQua(tabla, cadselect) Then
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
Dim i As Byte
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
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    ' socios dados de alta durante la campaña anterior
    Sql = "rsocios.codsocio = " & DBSet(txtCodigo(59).Text, "N")
    If Not AnyadirAFormula(cadselect, Sql) Then Exit Sub
    
    tabla = "rsocios"
    
    If HayRegistros(tabla, cadselect) Then
        Me.Label1(1).Caption = "Cargando tabla temporal"
        Me.Label1(1).visible = True
        Me.Refresh
        DoEvents
        If CargarTablaTemporal3(tabla, cadselect, "0", Me.Pb7) Then
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
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim b As Boolean

InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    'D/H socio
    cDesde = Trim(txtCodigo(10).Text)
    cHasta = Trim(txtCodigo(11).Text)
    nDesde = txtNombre(10).Text
    nHasta = txtNombre(11).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H Fecha factura
    cDesde = Trim(txtCodigo(8).Text)
    cHasta = Trim(txtCodigo(9).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.fecaport}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    tabla = "raportacion INNER JOIN rsocios ON raportacion.codsocio = rsocios.codsocio and rsocios.fechabaja is null "

    If HayRegistros(tabla, cadselect) Then
        Me.Label1(1).Caption = "Cargando tabla temporal"
        Me.Label1(1).visible = True
        Me.Refresh
        DoEvents
        If CargarTablaTemporal(tabla, cadselect, txtCodigo(4).Text, txtCodigo(5).Text, Me.Pb2, txtCodigo(123)) Then
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

Private Function SociosEnSeccion(vtabla As String, vWhere As String, Seccion As Integer) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim b As Boolean

    On Error GoTo ESocSec

    SociosEnSeccion = False

    'Seleccionamos los distintos socios, cuentas que vamos a facturar
    Sql = "SELECT DISTINCT " & vtabla & ".codigo1 codsocio"
    Sql = Sql & " from " & vtabla
    If vWhere <> "" Then Sql = Sql & " where " & vWhere
    Sql = Sql & " order by 1 "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    b = True

    While Not Rs.EOF And b
        Sql = "select * from rsocios_seccion where codsocio = " & DBSet(Rs!Codsocio, "N") & " and codsecci = " & DBSet(Seccion, "N")

        If Not (RegistrosAListar(Sql, cAgro) > 0) Then
        'si no lo encuentra
            b = False 'no encontrado
            Sql = "El Socio " & Format(Rs!Codsocio, "000000") & " no tiene registro en la seccion " & Seccion
        End If

        Rs.MoveNext
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
Dim Rs As ADODB.Recordset
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
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Sql2 = "insert into raportacion (codsocio,fecaport,codaport,descripcion,campanya,kilos,importe) values "

    Campanya = Mid(Format(Year(CDate(txtCodigo(8).Text)) + 1, "0000"), 3, 2) & "/" & Mid(Format(Year(CDate(txtCodigo(9).Text)), "0000"), 3, 2)
    Descripc = "ACUMULADA " & Campanya

    b = True

    Pb2.visible = True
    Pb2.Max = TotalRegistrosConsulta(Sql)
    Pb2.Value = 0
    
    While Not Rs.EOF And b
        IncrementarProgresNew Pb2, 1
    
        SqlValues = ""
        
        Sql = "select importe from raportacion where codsocio=" & DBSet(Rs!Codigo1, "N") & " and codaport=0 and fecaport=" & DBSet(txtCodigo(8).Text, "F")
    
        ImporIni = DevuelveValor(Sql)
        Importe = ImporIni + DBLet(Rs!importe4, "N")
    
        SqlExiste = "select count(*) from raportacion where codsocio = " & DBSet(Rs!Codigo1, "N") & " and codaport=0 and fecaport=" & DBSet(txtCodigo(14).Text, "F")
        b = (TotalRegistros(SqlExiste) = 0)
        
        If Not b Then
            MsgBox "El registro para el socio " & Format(DBLet(Rs!Codigo1, "N"), "000000") & " de fecha " & DBSet(txtCodigo(9).Text, "F") & " y tipo 0 existe. Revise.", vbExclamation
        Else
            
            '[Monica]27/10/2015: en el caso de que el socio no quiera devolucion grabamos el registro de acumulado anterior
            If NoDevolverAporSocio(CStr(Rs!Codigo1)) Then
                SqlValues = "(" & DBSet(Rs!Codigo1, "N") & "," & DBSet(txtCodigo(14).Text, "F") & ",0," & DBSet(Descripc, "T") & ","
                SqlValues = SqlValues & DBSet(Campanya, "T") & "," & DBSet(Rs!importe1, "N") & "," & DBSet(ImporIni, "N") & ")"
            Else
                ' como estaba
                SqlValues = "(" & DBSet(Rs!Codigo1, "N") & "," & DBSet(txtCodigo(14).Text, "F") & ",0," & DBSet(Descripc, "T") & ","
                SqlValues = SqlValues & DBSet(Campanya, "T") & "," & DBSet(Rs!importe2, "N") & "," & DBSet(Importe, "N") & ")"
            End If
            
            conn.Execute Sql2 & SqlValues
            
            '[Monica]09/04/2019: en caso de que hay devolucion de capital social grabo registro
            If DBLet(Rs!Precio2, "N") <> 0 Then
'                SqlValues = "(" & DBSet(Rs!Codigo1, "N") & "," & DBSet(txtCodigo(14).Text, "F") & ",3," & DBSet(txtCodigo(124), "T") & ","
'                SqlValues = SqlValues & DBSet(Campanya, "T") & "," & DBSet(Rs!importe2, "N") & "," & DBSet(Rs!Importe5, "N") & ")"
                If DBLet(Rs!importe3, "N") < 0 Then
                    SqlValues = "update raportacion set importe = (" & DBSet(ImporIni + Rs!importe5, "N") & "), kilos = " & DBSet(Rs!importe1, "N")
                Else
                    SqlValues = "update raportacion set importe = (" & DBSet(Importe + Rs!importe5, "N") & "), kilos = " & DBSet(Rs!importe2, "N")
                End If
                SqlValues = SqlValues & " where codsocio=" & DBSet(Rs!Codigo1, "N") & " and codaport=0 and fecaport=" & DBSet(txtCodigo(14).Text, "F")
                
                conn.Execute SqlValues
            End If
            
            MensError = "Insertando cobro en tesoreria"
            If NoDevolverAporSocio(CStr(Rs!Codigo1)) Then
                '[Monica]09/04/2019: insertamos en tesoreria la devolucion de capital unicamente
                b = True
                If DBLet(Rs!Precio2, "N") <> 0 Then
                    If DBLet(Rs!importe3) > 0 Then
                        b = InsertarEnTesoreriaNewAPO(MensError, Rs!Codigo1, DBLet(Rs!importe5, "N") + DBLet(Rs!importe4, "N"), txtCodigo(15).Text, txtCodigo(17).Text, txtCodigo(16).Text, txtCodigo(18).Text, txtCodigo(14).Text, 0)
                    Else
                        b = InsertarEnTesoreriaNewAPO(MensError, Rs!Codigo1, DBLet(Rs!importe5, "N"), txtCodigo(15).Text, txtCodigo(17).Text, txtCodigo(16).Text, txtCodigo(18).Text, txtCodigo(14).Text, 0)
                    End If
                End If
            Else
                '[Monica]10/04/2019: le añadimos la devolucion de capital social si la hay
                b = InsertarEnTesoreriaNewAPO(MensError, Rs!Codigo1, DBLet(Rs!importe4, "N"), txtCodigo(15).Text, txtCodigo(17).Text, txtCodigo(16).Text, txtCodigo(18).Text, txtCodigo(14).Text, 0)
            End If
        End If
    
    
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
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
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim Cad1 As String

InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H socio
    cDesde = Trim(txtCodigo(0).Text)
    cHasta = Trim(txtCodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H Fecha factura
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{raportacion.fecaport}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    
    tabla = "raportacion INNER JOIN rsocios ON raportacion.codsocio = rsocios.codsocio and rsocios.fechabaja is null "
    
    
    If HayRegistros(tabla, cadselect) Then
        If CargarTablaTemporal(tabla, cadselect, txtCodigo(6).Text, txtCodigo(7).Text, Me.Pb1, txtCodigo(125)) Then
            '[Monica]20/01/2016: si es mogente solo las de regularizacion negativa
            If chkNegativas.Value = 1 Then BorrarPositivas
            
            cadFormula = ""
            If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
            
            '[Monica]20/04/2016: para el caso de Mogente quitamos los que tienen importe 0 ( no fecha de baja )
            If vParamAplic.Cooperativa = 3 Then
                If Not AnyadirAFormula(cadFormula, "{tmpinformes.importe3} <> 0 ") Then Exit Sub
            End If
            
            Cad1 = "tmpinformes.codusu = " & vUsu.Codigo
            If vParamAplic.Cooperativa = 3 Then
                Cad1 = Cad1 & " and tmpinformes.importe3 <> 0" ' antes estaba importe2
            End If
            
            
            If Not HayRegistros("tmpinformes", Cad1) Then Exit Sub
            
            Select Case Opcionlistado
                Case 1 'Informe de aportaciones
                    'Nombre fichero .rpt a Imprimir
                    indRPT = 70 'informe de APORTACIONES
                
                    If Not PonerParamRPT(indRPT, "", 1, nomDocu) Then Exit Sub
                                       
                    cadNombreRPT = nomDocu '"rAPOInforme.rpt"
                    
                    frmImprimir.NombreRPT = cadNombreRPT
                    
                    cadTitulo = "Informe Aportaciones"
                    LlamarImprimir
                Case 3 ' Certificado de aportaciones
                    cadParam = cadParam & "pDesdeFecha=""" & txtCodigo(2).Text & """|"
                    cadParam = cadParam & "pHastaFecha=""" & txtCodigo(3).Text & """|"
                    cadParam = cadParam & "pFecha=""" & txtCodigo(12).Text & """|"
                    numParam = numParam + 3
                    '[Monica]11/03/2015:imprimimos el resumen
                    If vParamAplic.Cooperativa = 3 Then
                        cadParam = cadParam & "pResumen=" & Check3.Value & "|"
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


Private Function CargarTablaTemporal(nTabla1 As String, nSelect1 As String, Precio1 As String, Precio2 As String, ByRef Pb1 As ProgressBar, PrecioDev As String) As Boolean
Dim Rs As ADODB.Recordset
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
Dim ImporteDev As Currency

    On Error GoTo eCargarTablaTemporal

    If ExistenRegistrosAcumulados(nTabla1, nSelect1) Then
        CargarTablaTemporal = False
        Exit Function
    End If

    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql


    '[Monica]09/04/2019: añadimos las columnas de devolucion de capital para el certificado precio2 e importe5
    Sql = "insert into tmpinformes (codusu, codigo1, nombre1, importe1, importe2, importe3, precio1, importe4, precio2, importe5) values "
    
    Sql2 = " select raportacion.codsocio, nomsocio, fecaport, codaport, kilos "
    Sql2 = Sql2 & " from " & nTabla1
    
    If nSelect1 <> "" Then Sql2 = Sql2 & " where  " & nSelect1
    Sql2 = Sql2 & " order by 1, 3, 4"
    
    
    Pb1.visible = True
    Pb1.Max = TotalRegistrosConsulta(Sql2)
    Pb1.Value = 0
    
    
    cValues = ""
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        SocioAnt = Rs.Fields(0).Value
        NombreAnt = Rs.Fields(1).Value
        
        Kilos = 0
        Nregs = 0
        AcumAnt = 0
    End If
    
    Entro = False
    
    While Not Rs.EOF
        Entro = True
        
        Pb1.Value = Pb1.Value + 1
        DoEvents
        
        If SocioAnt <> Rs.Fields(0).Value Then
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
            cValues = cValues & DBSet(Importe, "N") & ","
            
            '[Monica]09/04/2019: añadimos las columnas correspondientes a la devolucion de capital
            cValues = cValues & DBSet(ComprobarCero(PrecioDev) * (-1), "N") & ","
            If Diferencia > 0 Then
                ImporteDev = Round2(KilosMed * ComprobarCero(PrecioDev) * (-1), 2)
            Else
                ImporteDev = Round2(AcumAnt * ComprobarCero(PrecioDev) * (-1), 2)
            End If
            
            cValues = cValues & DBSet(ImporteDev, "N") & "),"
            
            Kilos = 0
            Nregs = 0
            AcumAnt = 0
            
            SocioAnt = Rs.Fields(0).Value
            NombreAnt = Rs.Fields(1).Value
        
        End If
    
        If Rs!Codaport = 0 Then
            AcumAnt = Rs!Kilos
            Nregs = 0
        Else
            Kilos = Kilos + Rs!Kilos
            Nregs = Nregs + 1
        End If
        
        Rs.MoveNext
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
        cValues = cValues & DBSet(Importe, "N") & ","
    
        '[Monica]09/04/2019: añadimos las columnas correspondientes a la devolucion de capital
        cValues = cValues & DBSet(ComprobarCero(PrecioDev) * (-1), "N") & ","
        If Diferencia > 0 Then
            ImporteDev = Round2(KilosMed * ComprobarCero(PrecioDev) * (-1), 2)
        Else
            ImporteDev = Round2(AcumAnt * ComprobarCero(PrecioDev) * (-1), 2)
        End If
        cValues = cValues & DBSet(ImporteDev, "N") & "),"
    
        Kilos = 0
        Nregs = 0
        AcumAnt = 0
    End If

    If cValues <> "" Then
        cValues = Mid(cValues, 1, Len(cValues) - 1)
        conn.Execute Sql & cValues
    End If

    Set Rs = Nothing

    CargarTablaTemporal = True
    Pb1.visible = False
    
    Exit Function
    
eCargarTablaTemporal:
    MuestraError Err.Number, "Cargar Tabla Temporal", Err.Description
End Function

Private Function ExistenRegistrosAcumulados(nTabla As String, nWhere As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim i As Long
Dim cadMen As String
Dim Cad As String


    On Error GoTo eExistenRegistrosAcumulados
    
    ExistenRegistrosAcumulados = False
    
    Sql = "select raportacion.codsocio, count(*) from " & nTabla
    Sql = Sql & " where codaport = 0 "
    If nWhere <> "" Then
        Sql = Sql & " and " & nWhere
    End If
    Sql = Sql & " group by 1 "
    Sql = Sql & " having count(*) > 1"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        cadMen = "Los siguientes socios tienen más de un registro de acumulado anterior entre las fechas: "
        i = 0
        While Not Rs.EOF
            i = i + 1
            Cad = Cad & Format(Rs.Fields(0), "000000") & ","
            
            If i = 10 Then
                Cad = Cad & vbCrLf
                i = 0
            End If
            
            Rs.MoveNext
        Wend
        
    End If
    Set Rs = Nothing
    
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
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim b As Boolean
Dim Sql As String


InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    ' socios dados de alta durante la campaña
    Sql = "((rsocios.fechaalta between " & DBSet(vParam.FecIniCam, "F") & " and " & DBSet(vParam.FecFinCam, "F") & ") or "
    Sql = Sql & " rsocios.codsocio in (select codsocio from rsocios_seccion where codsecci = " & DBSet(vParamAplic.SeccionAlmaz, "N")
    Sql = Sql & " and fecalta between " & DBSet(vParam.FecIniCam, "F") & " and " & DBSet(vParam.FecFinCam, "F")
    Sql = Sql & " and fecbaja is null)) "
    
    If Not AnyadirAFormula(cadselect, Sql) Then Exit Sub
    
    
    Sql = "rsocios.fechabaja is null"
    If Not AnyadirAFormula(cadselect, Sql) Then Exit Sub
    
    Sql = "rsocios.codsocio in (select codsocio from (rcampos inner join variedades on rcampos.codvarie = variedades.codvarie) "
    Sql = Sql & " inner join productos on variedades.codprodu = productos.codprodu "
    Sql = Sql & " where productos.codgrupo = 5) "
    If Not AnyadirAFormula(cadselect, Sql) Then Exit Sub
    
    
    tabla = "rsocios"
    
    
    If HayRegistros(tabla, cadselect) Then
        Me.Label1(1).Caption = "Cargando tabla temporal"
        Me.Label1(1).visible = True
        Me.Refresh
        DoEvents
        If CargarTablaTemporal2(tabla, cadselect, txtCodigo(60).Text, Me.Pb6) Then
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
                If ActualizarRegularizacionAltaSocio(txtCodigo(60).Text) Then
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
    
        Select Case Opcionlistado
            Case 1 ' informe de aportaciones
                PonerFoco txtCodigo(0)
                txtCodigo(3).Text = Format(vParam.FecFinCam, "dd/mm/yyyy")
            Case 2 ' regularizacion
                txtCodigo(9).Text = Format(vParam.FecFinCam, "dd/mm/yyyy")
                txtCodigo(14).Text = Format(DateAdd("d", 1, vParam.FecFinCam), "dd/mm/yyyy")
            
                '[Monica]30/01/2014: valores por defecto de las formas de pago
                txtCodigo(16).Text = Format(vParamAplic.ForpaNega, "000")
                txtCodigo_LostFocus (16)
                txtCodigo(17).Text = Format(vParamAplic.ForpaPosi, "000")
                txtCodigo_LostFocus (17)
            
                PonerFoco txtCodigo(10)
            Case 3 ' Certificado de Aportaciones
                PonerFoco txtCodigo(0)
                txtCodigo(3).Text = Format(vParam.FecFinCam, "dd/mm/yyyy")
                txtCodigo(12).Text = Format(Now, "dd/mm/yyyy")
            Case 4 ' Informe de Aportaciones en el mantenimiento
                PonerFoco txtCodigo(23)
                '[Monica]18/01/2016: Añadimos la relacion con la cooperativa
                Combo1(1).ListIndex = 0
            Case 5 ' calculo de aportaciones de quatretonda
                PonerFoco txtCodigo(32)
            Case 6 ' listado de aportaciones para quatretonda
                Opcion1(0).Value = True
                PonerFoco txtCodigo(33)
                Combo1(0).ListIndex = 0
            Case 7 ' borrrado masivo de aportaciones de quatretonda
                PonerFoco txtCodigo(44)
            Case 8 ' integracion en tesoreria de quatretonda
                PonerFoco txtCodigo(44)
                
                '[Monica]30/01/2014: valores por defecto de las formas de pago
                txtCodigo(40).Text = Format(vParamAplic.ForpaNega, "000")
                txtCodigo_LostFocus (40)
                txtCodigo(42).Text = Format(vParamAplic.ForpaPosi, "000")
                txtCodigo_LostFocus (42)

            Case 9 ' integracion en tesoreria alta de socios moixent
                PonerFoco txtCodigo(60)
            
                '[Monica]30/01/2014: valores por defecto de las formas de pago
                txtCodigo(52).Text = Format(vParamAplic.ForpaNega, "000")
                txtCodigo_LostFocus (52)
                txtCodigo(53).Text = Format(vParamAplic.ForpaPosi, "000")
                txtCodigo_LostFocus (53)
            
            Case 10 ' integracion en tesoreria baja de socios moixent
                PonerFoco txtCodigo(59)
                
                '[Monica]30/01/2014: valores por defecto de las formas de pago
                txtCodigo(56).Text = Format(vParamAplic.ForpaNega, "000")
                txtCodigo_LostFocus (56)
                txtCodigo(55).Text = Format(vParamAplic.ForpaPosi, "000")
                txtCodigo_LostFocus (55)
                
            Case 11 ' Insercion de aportaciones de Bolbaite
                PonerFoco txtCodigo(61)
                txtCodigo(69).Text = vParamAplic.PorcenAFO ' por defecto
                If txtCodigo(69).Text <> "" Then txtCodigo(69).Text = Format(txtCodigo(69).Text, "##0.00")
            
            Case 12 ' Impresion de Recibos de Bolbaite
                PonerFoco txtCodigo(61)
                txtCodigo(70).Text = Format(Now, "dd/mm/yyyy")
                
            Case 13 ' Aportacion obligatoria
                PonerFoco txtCodigo(74)
                txtCodigo(74).Text = Format(Now, "dd/mm/yyyy")
                
                
            Case 14 ' integracion a contabilidad de aportaciones bolbaite
                PonerFoco txtCodigo(81)
                txtCodigo(86).Text = Format(Now, "dd/mm/yyyy")
                
                '[Monica]30/01/2014:
                txtCodigo(85).Text = Format(vParamAplic.ForpaNega, "000")
                txtCodigo_LostFocus (85)
                txtCodigo(84).Text = Format(vParamAplic.ForpaPosi, "000")
                txtCodigo_LostFocus (84)
                
            Case 15 ' certificado de retenciones
                PonerFoco txtCodigo(88)
                
            Case 16 ' devolucion de aportaciones
                PonerFoco txtCodigo(103)
                
            Case 17 ' devolucion de aportaciones de quatretonda
                PonerFoco txtCodigo(107)
                
            Case 18 ' certificado de aportaciones de coopic
                txtCodigo(115).Text = vParam.PerContacto
                PonerFoco txtCodigo(119)
                
        End Select
        Screen.MousePointer = vbDefault
    
    End If
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
    For H = 0 To 29
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    For H = 33 To 53
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    For H = 54 To 60
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    
    
    For H = 0 To imgAyuda.Count - 1
        imgAyuda(H).Picture = frmPpal.ImageListB.ListImages(10).Picture
    Next H


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
    Me.FrameCertificadoCPi.visible = False
    
    Select Case Opcionlistado
        Case 1 ' rendimiento por articulo
            FrameCobrosVisible True, H, W
            tabla = "raportacion"
            Me.Pb1.visible = False
            Frame1.visible = False
            Frame1.Enabled = False
            Label1(0).Caption = "Informe de Aportaciones"
        
        Case 2 ' regularizacion
            ConexionConta vParamAplic.SeccionAlmaz
        
            FrameRegularizacionVisible True, H, W
            tabla = "raportacion"
            Me.Pb1.visible = False
            
            '[Monica]09/04/2019: por si quieren aplicar devolucion
            Me.FrameDevolucion.Enabled = False
            Me.ChkDevolucion.Value = 0
            
            
        Case 3 ' Certificado de aportaciones
            FrameCobrosVisible True, H, W
            tabla = "raportacion"
            Me.Pb1.visible = False
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
            
            '[Monica]09/04/2019: devolucion de capital
            Me.Label4(146).visible = (vParamAplic.Cooperativa = 3)
            Me.txtCodigo(125).Enabled = (vParamAplic.Cooperativa = 3)
            Me.txtCodigo(125).visible = (vParamAplic.Cooperativa = 3)
            
            
        Case 4 ' Informe de aportaciones
            FrameInformesVisible True, H, W
            tabla = "raportacion"
            Me.Pb1.visible = False
            Label1(0).Caption = "Certificado de Aportaciones"
            
            CargaCombo
                
        Case 5 ' Cálculo de Aportaciones de Quatretonda
            FrameCalculoAporQuaVisible True, H, W
            tabla = "rcampos"
            Me.Pb1.visible = False
            Label1(0).Caption = "Cálculo de Aportaciones"
    
        Case 6 ' Listado de aportaciones para quatretonda
            FrameListAporQuaVisible True, H, W
            tabla = "raporhco"
            Me.Pb1.visible = False
            CargaCombo
                    
        Case 7 ' borrado masivo
            FrameListAporQuaVisible True, H, W
            tabla = "raporhco"
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
            FrameDevolAporQuaVisible True, H, W
            tabla = "raporhco"
            Me.Pb1.visible = False
            CargaCombo
            txtCodigo(112).Text = Format(Now, "dd/mm/yyyy")
            
            
        Case 8 ' integracion en tesoreria
            ConexionConta vParamAplic.Seccionhorto
            FrameIntTesorQuaVisible True, H, W
            tabla = "raporhco"
            Me.Pb4.visible = False
            
        Case 9 ' integracion en tesoresia del alta de socios de mogente
            ConexionConta vParamAplic.SeccionAlmaz
        
            FrameRegAltaSociosVisible True, H, W
            tabla = "rsocios"
            Me.Pb6.visible = False
            
        Case 10 ' integracion en tesoresia del alta de socios de mogente
            ConexionConta vParamAplic.SeccionAlmaz
        
            FrameRegBajaSociosVisible True, H, W
            tabla = "rsocios"
            Me.Pb7.visible = False
            
        Case 11 ' insercion de aportaciones para bolbaite
            FrameInsertarApoBolVisible True, H, W
            tabla = "rfactsoc"
            Me.Pb8.visible = False
            Frame12.visible = False
            Frame12.Enabled = False
            
            CargarListView 0
            
        Case 12 ' Impresion de recibos de bolbaite
            FrameInsertarApoBolVisible True, H, W
            
            Label1(19).Caption = "Impresión de Recibos"
            tabla = "raportacion"
            Me.Pb8.visible = False
            Frame5.visible = False
            Frame5.Enabled = False
            Me.CmdAcepInsApoBol.Top = 5100
            Me.CmdCancel(8).Top = 5100
            
            CargarListView 0
            
        Case 13 ' aportacion obligatoria de bolbaite
            FrameAportacionObligatoriaVisible True, H, W
            
            tabla = "rsocios"
            Me.Pb9.visible = False
            
        Case 14
            FrameIntTesorBolVisible True, H, W
            
            ConexionConta vParamAplic.Seccionhorto
            tabla = "raportacion"
            Me.Pb10.visible = False
            
        Case 15 ' certificado de aportacion bolbaite
            FrameCertificadoBolVisible True, H, W
            
            tabla = "raportacion"
        
        Case 16 ' devolucion de aportaciones de bolbaite
            FrameDevolAporBolVisible True, H, W
            tabla = "raportacion"
            
        Case 18 ' certificado de aportaciones de coopic
            FrameCertificadoCPiVisible True, H, W
            tabla = "raportacion"
        
    End Select
    
    Me.Width = W + 70
    Me.Height = H + 350
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not vSeccion Is Nothing Then
        vSeccion.CerrarConta
        Set vSeccion = Nothing
    End If
End Sub

Private Sub frmApo_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de tipo de aportaciones
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtCodigo(indCodigo).Text = Format(txtCodigo(indCodigo).Text, "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmCla_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de clases
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
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

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " {raportacion.codaport} in (" & CadenaSeleccion & ")"
        Sql2 = " {raportacion.codaport} in [" & CadenaSeleccion & "]"
    Else
        Sql = " {raportacion.codaport} = -1 "
    End If
    If Not AnyadirAFormula(cadselect, Sql) Then Exit Sub
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
    If Not AnyadirAFormula(cadselect, Sql) Then Exit Sub
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
    If Not AnyadirAFormula(cadselect, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub

Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub Image1_Click(Index As Integer)
Dim i As Integer
Dim TotalCant As Currency
Dim TotalImporte As Currency

    Screen.MousePointer = vbHourglass
    
    Select Case Index
        ' informe de resultados y listado de retenciones
        Case 2
            For i = 1 To ListView1(0).ListItems.Count
                ListView1(0).ListItems(i).Checked = True
            Next i
        Case 3
            For i = 1 To ListView1(0).ListItems.Count
                ListView1(0).ListItems(i).Checked = False
            Next i
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
            
        ' certificado de aportaciones de coopic
        Case 35, 36, 37
            indice = Index + 81
            
        Case Else
            indice = Index
    End Select
    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
    imgFec(0).Tag = indice 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(indice).Text <> "" Then frmC.NovaData = txtCodigo(indice).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(0).Tag))
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
                
        'certificado de aportaciones de coopic
        Case 58
            AbrirFrmSocios (119)
        Case 53
            AbrirFrmSocios (120)
        Case 59, 60 ' tipo de aportacion
            AbrirFrmTipoAportacion (Index + 62)
    End Select
    
    PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub AbrirFrmCuentas(indice As Integer)
    indCodigo = indice
    Set frmCtas = New frmCtasConta
    frmCtas.DatosADevolverBusqueda = "0|1|"
    frmCtas.CodigoActual = txtCodigo(indCodigo)
    frmCtas.Show vbModal
    Set frmCtas = Nothing
End Sub

Private Sub AbrirFrmForpaConta(indice As Integer)
    indCodigo = indice
    Set frmFPa = New frmForpaConta
    frmFPa.DatosADevolverBusqueda = "0|1|"
    frmFPa.CodigoActual = txtCodigo(indCodigo)
    frmFPa.Show vbModal
    Set frmFPa = Nothing
End Sub





Private Sub Opcion1_Click(Index As Integer)
    Check1.Enabled = Opcion1(0).Value
    If Not Check1.Enabled Then Check1.Value = 0
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
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
            
            'certificado de aportaciones coopic 11/06/2018
            Case 119: KEYBusqueda KeyAscii, 58 'socio desde
            Case 120: KEYBusqueda KeyAscii, 53 'socio hasta
            Case 117: KEYFecha KeyAscii, 59 'fecha desde
            Case 118: KEYFecha KeyAscii, 60 'fecha hasta
            Case 121: KEYBusqueda KeyAscii, 59 'tipo de aportacion
            Case 122: KEYBusqueda KeyAscii, 60 'tipo de aportacion
            Case 116: KEYFecha KeyAscii, 35 'fecha certificado
            
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
Dim Cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 0, 1, 10, 11, 23, 24, 29, 30, 36, 37, 44, 45, 59, 66, 67, 77, 78, 81, 82, 88, 89, 103, 104, 107, 108, 119, 120 'socios
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
        Case 8, 9, 12, 14, 15, 20, 35, 41, 46, 47, 34, 49, 51, 54, 57, 64, 65, 74, 86, 79, 80, 100 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index), True
            
        Case 2, 3, 21, 22, 76, 90, 91, 101, 102, 106, 111, 112, 117, 118, 116
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index), False
            
        Case 6, 7, 60, 123, 125 'precios
            PonerFormatoDecimal txtCodigo(Index), 7
            
        Case 16, 17, 40, 42, 52, 53, 55, 56, 84, 85 ' forma de pago
            If vSeccion Is Nothing Then Exit Sub
            
            If vParamAplic.ContabilidadNueva Then
                If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtCodigo(Index).Text, "N")
            Else
                If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtCodigo(Index).Text, "N")
            End If
            If txtNombre(Index).Text = "" Then
                MsgBox "Forma de Pago  no existe en la contabilidad. Reintroduzca.", vbExclamation
            End If
        
        Case 18, 33, 50, 58, 83 ' cta de banco
            If vSeccion Is Nothing Then Exit Sub
        
            If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = PonerNombreCuenta(txtCodigo(Index), 2)
            
        Case 4, 5 ' importes
            PonerFormatoDecimal txtCodigo(Index), 7
            
        Case 13, 19, 68, 71, 75, 87, 96, 105, 121, 122 ' codigo de aportaciones
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rtipoapor", "nomaport", "codaport", "N")
        
        Case 27, 28, 38, 39, 43, 48, 109, 110 'CLASES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "clases", "nomclase", "codclase", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
        
        Case 32, 33 'SECCIONES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rseccion", "nomsecci", "codsecci", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
    
        Case 25 'Año
            PonerFormatoEntero txtCodigo(Index)
        
        Case 26 ' Euros/hanegada para el calculo de aportaciones quatetonda
            PonerFormatoDecimal txtCodigo(Index), 3
        
        Case 31 'Ejercicio
            PonerFormatoEntero txtCodigo(Index)
        
        Case 69 'porcentaje de aportacion
            PonerFormatoDecimal txtCodigo(Index), 4
            
        Case 61, 62 'numero de factura
            PonerFormatoEntero txtCodigo(Index)
            
        Case 73 ' importe de la aportacion obligatoria
            PonerFormatoDecimal txtCodigo(Index), 3
            
        Case 92, 93, 94
            txtCodigo(Index).Text = UCase(txtCodigo(Index))
        
        Case 123 ' precion de devolucion de capital social
            PonerFormatoDecimal txtCodigo(123), 7
            
        
    End Select
End Sub


Private Sub FrameCalculoAporQuaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCalculoAporQua.visible = visible
    If visible = True Then
        Me.FrameCalculoAporQua.Top = -90
        Me.FrameCalculoAporQua.Left = 0
        Me.FrameCalculoAporQua.Height = 7140
        Me.FrameCalculoAporQua.Width = 8085
        W = Me.FrameCalculoAporQua.Width
        H = Me.FrameCalculoAporQua.Height
    End If
End Sub


Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 6555 '5790
        Me.FrameCobros.Width = 8085
        W = Me.FrameCobros.Width
        H = Me.FrameCobros.Height
    End If
End Sub


Private Sub FrameInformesVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameInforme.visible = visible
    If visible = True Then
        Me.FrameInforme.Top = -90
        Me.FrameInforme.Left = 0
        Me.FrameInforme.Height = 6300
        Me.FrameInforme.Width = 8085
        W = Me.FrameInforme.Width
        H = Me.FrameInforme.Height
    End If
End Sub

Private Sub FrameListAporQuaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameListAporQua.visible = visible
    If visible = True Then
        Me.FrameListAporQua.Top = -90
        Me.FrameListAporQua.Left = 0
        Me.FrameListAporQua.Height = 6660
        Me.FrameListAporQua.Width = 8085
        W = Me.FrameListAporQua.Width
        H = Me.FrameListAporQua.Height
    End If
End Sub

Private Sub FrameDevolAporQuaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameDevolAporQua.visible = visible
    If visible = True Then
        Me.FrameDevolAporQua.Top = -90
        Me.FrameDevolAporQua.Left = 0
        Me.FrameDevolAporQua.Height = 7140
        Me.FrameDevolAporQua.Width = 8085
        W = Me.FrameDevolAporQua.Width
        H = Me.FrameDevolAporQua.Height
    End If
End Sub


Private Sub FrameIntTesorQuaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameIntTesorQua.visible = visible
    If visible = True Then
        Me.FrameIntTesorQua.Top = -90
        Me.FrameIntTesorQua.Left = 0
        Me.FrameIntTesorQua.Height = 7530
        Me.FrameIntTesorQua.Width = 8085
        W = Me.FrameIntTesorQua.Width
        H = Me.FrameIntTesorQua.Height
    End If
End Sub

Private Sub FrameRegularizacionVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameRegularizacion.visible = visible
    If visible = True Then
        Me.FrameRegularizacion.Top = -90
        Me.FrameRegularizacion.Left = 0
        Me.FrameRegularizacion.Height = 8880 '7530
        Me.FrameRegularizacion.Width = 8085
        W = Me.FrameRegularizacion.Width
        H = Me.FrameRegularizacion.Height
    End If
End Sub

Private Sub FrameInsertarApoBolVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameInsertarApoBol.visible = visible
    If visible = True Then
        Me.FrameInsertarApoBol.Top = -90
        Me.FrameInsertarApoBol.Left = 0
        Me.FrameInsertarApoBol.Height = 7530
        
        If Opcionlistado = 12 Then Me.FrameInsertarApoBol.Height = 6460
        
        Me.FrameInsertarApoBol.Width = 8085 '6735 '6555
        W = Me.FrameInsertarApoBol.Width
        H = Me.FrameInsertarApoBol.Height
    End If
End Sub


Private Sub FrameAportacionObligatoriaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameAporObligatoria.visible = visible
    If visible = True Then
        Me.FrameAporObligatoria.Top = -90
        Me.FrameAporObligatoria.Left = 0
        Me.FrameAporObligatoria.Height = 6330
        Me.FrameAporObligatoria.Width = 8085
        W = Me.FrameAporObligatoria.Width
        H = Me.FrameAporObligatoria.Height
    End If
End Sub

Private Sub FrameIntTesorBolVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameIntTesorBol.visible = visible
    If visible = True Then
        Me.FrameIntTesorBol.Top = -90
        Me.FrameIntTesorBol.Left = 0
        Me.FrameIntTesorBol.Height = 7530
        Me.FrameIntTesorBol.Width = 8085
        W = Me.FrameIntTesorBol.Width
        H = Me.FrameIntTesorBol.Height
    End If
End Sub

Private Sub FrameCertificadoBolVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCertificadoBol.visible = visible
    If visible = True Then
        Me.FrameCertificadoBol.Top = -90
        Me.FrameCertificadoBol.Left = 0
        Me.FrameCertificadoBol.Height = 7530
        Me.FrameCertificadoBol.Width = 8085
        W = Me.FrameCertificadoBol.Width
        H = Me.FrameCertificadoBol.Height
    End If
End Sub

Private Sub FrameDevolAporBolVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameDevolAporBol.visible = visible
    If visible = True Then
        Me.FrameDevolAporBol.Top = -90
        Me.FrameDevolAporBol.Left = 0
        Me.FrameDevolAporBol.Height = 6900
        Me.FrameDevolAporBol.Width = 8085
        W = Me.FrameDevolAporBol.Width
        H = Me.FrameDevolAporBol.Height
    End If
End Sub


Private Sub FrameCertificadoCPiVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCertificadoCPi.visible = visible
    If visible = True Then
        Me.FrameCertificadoCPi.Top = -90
        Me.FrameCertificadoCPi.Left = 0
        Me.FrameCertificadoCPi.Height = 7740
        Me.FrameCertificadoCPi.Width = 8085
        W = Me.FrameCertificadoCPi.Width
        H = Me.FrameCertificadoCPi.Height
    End If
End Sub





Private Sub FrameRegAltaSociosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameRegAltaSocios.visible = visible
    If visible = True Then
        Me.FrameRegAltaSocios.Top = -90
        Me.FrameRegAltaSocios.Left = 0
        Me.FrameRegAltaSocios.Height = 5400
        Me.FrameRegAltaSocios.Width = 8130
        W = Me.FrameRegAltaSocios.Width
        H = Me.FrameRegAltaSocios.Height
    End If
End Sub


Private Sub FrameRegBajaSociosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameRegBajaSocios.visible = visible
    If visible = True Then
        Me.FrameRegBajaSocios.Top = -90
        Me.FrameRegBajaSocios.Left = 0
        Me.FrameRegBajaSocios.Height = 5400
        Me.FrameRegBajaSocios.Width = 8085
        W = Me.FrameRegBajaSocios.Width
        H = Me.FrameRegBajaSocios.Height
    End If
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

Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
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
    Set frmCla = New frmBasico2
    AyudaClasesCom frmCla, txtCodigo(indice).Text
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
    frmVar.CodigoActual = txtCodigo(indCodigo)
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub

Private Function HayRegistros(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
Dim Rs As ADODB.Recordset

    Sql = "Select * FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Rs.EOF Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        HayRegistros = False
    Else
        HayRegistros = True
    End If

End Function

Private Function ProcesarCambios(cadWhere As String) As Boolean
Dim Sql As String
Dim SQL1 As String
Dim i As Integer
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

    Select Case Opcionlistado
        Case 1
            If txtCodigo(6).Text = "" Then
                MsgBox "Debe introducir un valor en Precio Aumento de Kilos. Revise.", vbExclamation
                b = False
            End If
            If b Then
                If txtCodigo(7).Text = "" Then
                    MsgBox "Debe introducir un valor en Precio Disminución de Kilos. Revise.", vbExclamation
                    b = False
                End If
            End If
        Case 2
            If txtCodigo(4).Text = "" Then
                MsgBox "Debe introducir un valor en Precio Aumento de Kilos. Revise.", vbExclamation
                b = False
            End If
            If b Then
                If txtCodigo(5).Text = "" Then
                    MsgBox "Debe introducir un valor en Precio Disminución de Kilos. Revise.", vbExclamation
                    b = False
                End If
            End If
            If b Then
                If txtCodigo(8).Text = "" Or txtCodigo(9).Text = "" Then
                    MsgBox "Debe introducir valor en desde/hasta Fecha Factura. Revise.", vbExclamation
                    b = False
                End If
            End If
            
            If b Then
                If Me.ChkDevolucion.Value Then
                    If txtCodigo(123).Text = "" Then
                        MsgBox "Debe introducir un valor en Precio de Devolucion de capital. Revise.", vbExclamation
                        PonerFoco txtCodigo(123)
                        b = False
                    End If
'                    If b Then
'                        If txtCodigo(124) = "" Then
'                            MsgBox "Debe introducir un valor en Concepto de Devolucion de Capital. Revise.", vbExclamation
'                            PonerFoco txtCodigo(124)
'                            b = False
'                        End If
'                    End If
                End If
            End If
            
        Case 5 ' calculo de aportaciones de quatretonda
            If txtCodigo(32).Text = "" Then
                MsgBox "Debe introducir una sección. Reintroduzca.", vbExclamation
                PonerFoco txtCodigo(32)
                b = False
            End If
            ' debe introducir todos los datos para el calculo de aportaciones
            ' importe por hda
            If b Then
                If CDbl(ComprobarCero(txtCodigo(26).Text)) = "0" Then
                    MsgBox "Debe introducir el importe por hanegada. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(26)
                    b = False
                End If
            End If
            ' fecha de aportacion
            If b Then
                If txtCodigo(20).Text = "" Then
                    MsgBox "Debe introducir la Fecha de Aportación. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(20)
                    b = False
                End If
            End If
            ' año
            If b Then
                If txtCodigo(25).Text = "" Then
                    MsgBox "Debe introducir el Año. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(25)
                    b = False
                End If
            End If
            ' Ejercicio
            If b Then
                If txtCodigo(31).Text = "" Then
                    MsgBox "Debe introducir el Ejercicio. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(31)
                    b = False
                End If
            End If
            
        '[Monica]15/09/2014
        Case 17 ' Devoluciones de aportaciones de quatretonda
            ' debe introducir todos los datos para el calculo de aportaciones
            ' fecha de aportacion
            If b Then
                If txtCodigo(112).Text = "" Then
                    MsgBox "Debe introducir la Fecha de Devolución de Aportación. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(112)
                    b = False
                End If
            End If
            ' Ejercicio
            If b Then
                If txtCodigo(98).Text = "" Then
                    MsgBox "Debe introducir el Ejercicio de devolución. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(98)
                    b = False
                End If
            End If
            
            
        Case 8 ' Integracion de aportaciones en tesoreria
            If txtCodigo(34).Text = "" Then
                MsgBox "Debe introducir la Fecha de Vencimiento. Reintroduzca.", vbExclamation
                PonerFoco txtCodigo(34)
                b = False
            End If
            
            If b Then
                If txtCodigo(33).Text = "" Then
                    MsgBox "Debe introducir la Cuenta de Banco Prevista. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(33)
                    b = False
                Else
                    If PonerNombreCuenta(txtCodigo(33), 2) = "" Then
'                        MsgBox "La Cuenta de Banco Prevista no existe en Contabilidad. Reintroduzca.", vbExclamation
                        PonerFoco txtCodigo(33)
                        b = False
                    End If
                End If
            End If
            
            If b Then
                If txtCodigo(40).Text = "" Then
                    MsgBox "Debe introducir la Forma de Pago Positiva. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(40)
                    b = False
                Else
                    If vParamAplic.ContabilidadNueva Then
                        vDevuelve = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtCodigo(40).Text, "N")
                    Else
                        vDevuelve = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtCodigo(40).Text, "N")
                    End If
                    If vDevuelve = "" Then
                        MsgBox "La Forma de Pago Positiva no existe en Contabilidad. Reintroduzca.", vbExclamation
                        PonerFoco txtCodigo(40)
                        b = False
                    End If
                End If
            End If
            
            If b Then
                If txtCodigo(42).Text = "" Then
                    MsgBox "Debe introducir la Forma de Pago Negativa. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(42)
                    b = False
                Else
                    If vParamAplic.ContabilidadNueva Then
                        vDevuelve = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtCodigo(42).Text, "N")
                    Else
                        vDevuelve = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtCodigo(42).Text, "N")
                    End If
                    If vDevuelve = "" Then
                        MsgBox "La Forma de Pago Negativa no existe en Contabilidad. Reintroduzca.", vbExclamation
                        PonerFoco txtCodigo(42)
                        b = False
                    End If
                End If
            End If
            
        Case 9 ' integracion en tesoreria de alta de socios solo para mogente
            If txtCodigo(60).Text = "" Then
                MsgBox "Debe introducir el precio kilo. Reintroduzca.", vbExclamation
                PonerFoco txtCodigo(60)
                b = False
            End If
            
            If b Then
                If txtCodigo(51).Text = "" Then
                    MsgBox "Debe introducir la Fecha de Vencimiento. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(51)
                    b = False
                End If
            End If
            
            If b Then
                If txtCodigo(50).Text = "" Then
                    MsgBox "Debe introducir la Cuenta de Banco Prevista. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(50)
                    b = False
                Else
                    If PonerNombreCuenta(txtCodigo(50), 2) = "" Then
                        PonerFoco txtCodigo(50)
                        b = False
                    End If
                End If
            End If
            
            If b Then
                If txtCodigo(52).Text = "" Then
                    MsgBox "Debe introducir la Forma de Pago Positiva. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(52)
                    b = False
                Else
                    If vParamAplic.ContabilidadNueva Then
                        vDevuelve = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtCodigo(52).Text, "N")
                    Else
                        vDevuelve = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtCodigo(52).Text, "N")
                    End If
                    If vDevuelve = "" Then
                        MsgBox "La Forma de Pago Positiva no existe en Contabilidad. Reintroduzca.", vbExclamation
                        PonerFoco txtCodigo(52)
                        b = False
                    End If
                End If
            End If
            
            If b Then
                If txtCodigo(53).Text = "" Then
                    MsgBox "Debe introducir la Forma de Pago Negativa. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(53)
                    b = False
                Else
                    If vParamAplic.ContabilidadNueva Then
                        vDevuelve = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtCodigo(53).Text, "N")
                    Else
                        vDevuelve = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtCodigo(53).Text, "N")
                    End If
                    If vDevuelve = "" Then
                        MsgBox "La Forma de Pago Negativa no existe en Contabilidad. Reintroduzca.", vbExclamation
                        PonerFoco txtCodigo(53)
                        b = False
                    End If
                End If
            End If
        
        Case 10 ' integracion en tesoreria de baja de socios solo para mogente
            If txtCodigo(54).Text = "" Then
                MsgBox "Debe introducir la Fecha de Devolución. Reintroduzca.", vbExclamation
                PonerFoco txtCodigo(54)
                b = False
            End If
            
            If b Then
                If txtCodigo(57).Text = "" Then
                    MsgBox "Debe introducir la Fecha de Vencimiento. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(57)
                    b = False
                End If
            End If
            
            If b Then
                If txtCodigo(58).Text = "" Then
                    MsgBox "Debe introducir la Cuenta de Banco Prevista. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(58)
                    b = False
                Else
                    If PonerNombreCuenta(txtCodigo(58), 2) = "" Then
                        PonerFoco txtCodigo(58)
                        b = False
                    End If
                End If
            End If
            
            If b Then
                If txtCodigo(56).Text = "" Then
                    MsgBox "Debe introducir la Forma de Pago Positiva. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(56)
                    b = False
                Else
                    If vParamAplic.ContabilidadNueva Then
                        vDevuelve = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtCodigo(56).Text, "N")
                    Else
                        vDevuelve = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtCodigo(56).Text, "N")
                    End If
                    If vDevuelve = "" Then
                        MsgBox "La Forma de Pago Positiva no existe en Contabilidad. Reintroduzca.", vbExclamation
                        PonerFoco txtCodigo(56)
                        b = False
                    End If
                End If
            End If
            
            If b Then
                If txtCodigo(55).Text = "" Then
                    MsgBox "Debe introducir la Forma de Pago Negativa. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(55)
                    b = False
                Else
                    If vParamAplic.ContabilidadNueva Then
                        vDevuelve = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtCodigo(55).Text, "N")
                    Else
                        vDevuelve = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtCodigo(55).Text, "N")
                    End If
                    If vDevuelve = "" Then
                        MsgBox "La Forma de Pago Negativa no existe en Contabilidad. Reintroduzca.", vbExclamation
                        PonerFoco txtCodigo(55)
                        b = False
                    End If
                End If
            End If
        
            ' vemos si el socio al que vamos a dar de baja tiene concepto de aportacion 0
            If b Then
                Sql = "select * from raportacion where raportacion.codsocio = " & DBSet(txtCodigo(59).Text, "N")
                If TotalRegistrosConsulta(Sql) = 0 Then
                    MsgBox "El socio a dar de baja no tiene registro de regularizacion. Revise.", vbExclamation
                    PonerFoco txtCodigo(59)
                    b = False
                End If
                ' vemos si el socio tiene fecha de baja
                If b Then
                    Sql = "select * from rsocios  "
                    Sql = Sql & " where codsocio = " & DBSet(txtCodigo(59).Text, "N") & " and not fechabaja is null "
                    If TotalRegistrosConsulta(Sql) = 0 Then
                        MsgBox "El socio a dar de baja no tiene fecha de baja. Revise.", vbExclamation
                        PonerFoco txtCodigo(59)
                        b = False
                    End If
                End If
                ' vemos si el socio esta en la seccion de almazara
                If b Then
                    Sql = "select * from rsocios_seccion where codsocio = " & DBSet(txtCodigo(59).Text, "N")
                    Sql = Sql & " and codsecci = " & DBSet(vParamAplic.SeccionAlmaz, "N")
                    If TotalRegistrosConsulta(Sql) = 0 Then
                        MsgBox "El socio a dar de baja no es de la sección de almazara. Revise.", vbExclamation
                        PonerFoco txtCodigo(59)
                        b = False
                    End If
                End If
                ' comprobamos que a este socio no se le haya hecho ya la devolucion
                If b Then
                    Sql = "select sum(importe) from raportacion where codsocio = " & DBSet(txtCodigo(59).Text, "N")
                    Sql = Sql & " and fecaport >= (select max(fecaport) from raportacion where codsocio = " & DBSet(txtCodigo(59).Text, "N")
                    Sql = Sql & " and codaport = 0) "
                    If DevuelveValor(Sql) = 0 Then
                        MsgBox "Al socio ya se le ha hecho la devolución de la aportación. Revise.", vbExclamation
                        PonerFoco txtCodigo(59)
                        b = False
                    End If
                End If
            End If
        
        Case 11 ' insercion de aportaciones Bolbaite
            ' descripcion
            If txtCodigo(63).Text = "" Then
                MsgBox "Debe introducir la descripción. Revise.", vbExclamation
                PonerFoco txtCodigo(63)
                b = False
            End If
            ' tipo de aportacion
            If b Then
                If txtCodigo(68).Text = "" Then
                    MsgBox "Debe introducir el Tipo de Aportación. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(68)
                    b = False
                Else
                    vDevuelve = DevuelveDesdeBDNew(cAgro, "rtipoapor", "nomaport", "codaport", txtCodigo(68).Text, "N")
                    If vDevuelve = "" Then
                        MsgBox "El tipo de Aportación no existe. Reintroduzca.", vbExclamation
                        PonerFoco txtCodigo(68)
                        b = False
                    End If
                End If
            End If
        
        Case 12 ' Impresion de recibos de aportaciones de bolbaite
            If txtCodigo(70).Text = "" Then
                MsgBox "Debe introducir la fecha de Impresión de Recibo. Reintroduzca.", vbExclamation
                PonerFoco txtCodigo(70)
                b = False
            End If
        
        Case 13 ' Aportacion obligatoria de bolbaite
            If txtCodigo(74).Text = "" Then
                MsgBox "Debe introducir la fecha de Aportación. Reintroduzca.", vbExclamation
                PonerFoco txtCodigo(74)
                b = False
            End If
            ' descripcion
            If txtCodigo(72).Text = "" Then
                MsgBox "Debe introducir la descripción. Revise.", vbExclamation
                PonerFoco txtCodigo(72)
                b = False
            End If
            ' tipo de aportacion
            If b Then
                If txtCodigo(71).Text = "" Then
                    MsgBox "Debe introducir el Tipo de Aportación. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(71)
                    b = False
                Else
                    vDevuelve = DevuelveDesdeBDNew(cAgro, "rtipoapor", "nomaport", "codaport", txtCodigo(71).Text, "N")
                    If vDevuelve = "" Then
                        MsgBox "El tipo de Aportación no existe. Reintroduzca.", vbExclamation
                        PonerFoco txtCodigo(71)
                        b = False
                    End If
                End If
            End If
        
        Case 14 ' integracion en tesoreria de bolbaite
            If txtCodigo(86).Text = "" Then
                MsgBox "Debe introducir la fecha de Vencimiento. Reintroduzca.", vbExclamation
                PonerFoco txtCodigo(86)
                b = False
            End If
        
            ' tipo de aportacion
            If b Then
                If txtCodigo(75).Text = "" Then
                    MsgBox "Debe introducir el Tipo de Aportación. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(75)
                    b = False
                End If
            End If
        
            If b Then
                If txtCodigo(83).Text = "" Then
                    MsgBox "Debe introducir la Cuenta de Banco Prevista. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(83)
                    b = False
                Else
                    If PonerNombreCuenta(txtCodigo(83), 2) = "" Then
                        PonerFoco txtCodigo(83)
                        b = False
                    End If
                End If
            End If
            
            If b Then
                If txtCodigo(85).Text = "" Then
                    MsgBox "Debe introducir la Forma de Pago Positiva. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(85)
                    b = False
                Else
                    If vParamAplic.ContabilidadNueva Then
                        vDevuelve = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtCodigo(85).Text, "N")
                    Else
                        vDevuelve = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtCodigo(85).Text, "N")
                    End If
                    If vDevuelve = "" Then
                        MsgBox "La Forma de Pago Positiva no existe en Contabilidad. Reintroduzca.", vbExclamation
                        PonerFoco txtCodigo(85)
                        b = False
                    End If
                End If
            End If
            
            If b Then
                If txtCodigo(84).Text = "" Then
                    MsgBox "Debe introducir la Forma de Pago Negativa. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(84)
                    b = False
                Else
                    If vParamAplic.ContabilidadNueva Then
                        vDevuelve = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtCodigo(84).Text, "N")
                    Else
                        vDevuelve = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtCodigo(84).Text, "N")
                    End If
                    If vDevuelve = "" Then
                        MsgBox "La Forma de Pago Negativa no existe en Contabilidad. Reintroduzca.", vbExclamation
                        PonerFoco txtCodigo(84)
                        b = False
                    End If
                End If
            End If
        
        Case 15 ' Certificado
            If b Then
                If txtCodigo(90).Text = "" Then
                    MsgBox "Debe introducir la fecha desde de Aportación. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(90)
                    b = False
                End If
            End If
            If b Then
                If txtCodigo(91).Text = "" Then
                    MsgBox "Debe introducir la fecha hasta de Aportación. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(91)
                    b = False
                End If
            End If
            If b Then
                If txtCodigo(76).Text = "" Then
                    MsgBox "Debe introducir la fecha de Certificado. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(76)
                    b = False
                End If
            End If
            ' tipo de aportacion
            If b Then
                If txtCodigo(87).Text = "" Then
                    MsgBox "Debe introducir el Tipo de Aportación. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(87)
                    b = False
                Else
                    '[Monica]05/12/2013
                    If txtNombre(87).Text = "" Then
                        MsgBox "El Tipo de Aportación no existe. Reintroduzca.", vbExclamation
                        PonerFoco txtCodigo(87)
                        b = False
                    End If
                End If
            End If
                    
            ' Presidente
            If b Then
                If txtCodigo(92).Text = "" Then
                    MsgBox "Debe introducir el Presidente. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(92)
                    b = False
                End If
            End If
            ' Secretario
            If b Then
                If txtCodigo(93).Text = "" Then
                    MsgBox "Debe introducir el Secretario. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(93)
                    b = False
                End If
            End If
            ' Tesorero
            If b Then
                If txtCodigo(94).Text = "" Then
                    MsgBox "Debe introducir el Tesorero. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(94)
                    b = False
                End If
            End If
            
        Case 16 ' devolucion de aportacion
            If b Then
                If txtCodigo(100).Text = "" Then
                    MsgBox "Debe introducir la fecha de Devolución. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(100)
                    b = False
                End If
            End If
            ' tipo de aportacion origen
            If b Then
                If txtCodigo(105).Text = "" Then
                    MsgBox "Debe introducir el Tipo de Aportación. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(105)
                    b = False
                End If
            End If
            ' tipo de aportacion destino
            If b Then
                If txtCodigo(96).Text = "" Then
                    MsgBox "Debe introducir el Nuevo Tipo de Aportación. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(96)
                    b = False
                End If
            End If
            ' descripcion
            If b Then
                If txtCodigo(99).Text = "" Then
                    MsgBox "Debe introducir la descripción. Revise.", vbExclamation
                    PonerFoco txtCodigo(99)
                    b = False
                End If
            End If
    
    
        Case 18 ' Certificado de coopic
'            If B Then
'                If txtCodigo(117).Text = "" Then
'                    MsgBox "Debe introducir la fecha desde de Aportación. Reintroduzca.", vbExclamation
'                    PonerFoco txtCodigo(117)
'                    B = False
'                End If
'            End If
'            If B Then
'                If txtCodigo(118).Text = "" Then
'                    MsgBox "Debe introducir la fecha hasta de Aportación. Reintroduzca.", vbExclamation
'                    PonerFoco txtCodigo(118)
'                    B = False
'                End If
'            End If
            If b Then
                If txtCodigo(116).Text = "" Then
                    MsgBox "Debe introducir la fecha de Certificado. Reintroduzca.", vbExclamation
                    PonerFoco txtCodigo(116)
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
Dim i As Byte
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
Dim vvIban As String

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
                    Case 3 ' Devolucion Capital Social
                        Text33csb = "'Devolución Capital Social de " & Format(DBLet(fecfactu, "F"), "dd/mm/yy") & "'"
                        Text41csb = "de " & DBSet(Importe, "N")
                End Select
                        
                CC = DBLet(vSocio.Digcontrol, "T")
                If DBLet(vSocio.Digcontrol, "T") = "**" Then CC = "00"
                
                '[Monica]03/07/2013: añado trim(codmacta)
                CadValuesAux2 = "(" & DBSet(letraser, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(fecfactu, "F") & ", 1," & DBSet(Trim(vSocio.CtaClien), "T") & ","
                CadValues2 = CadValuesAux2 & DBSet(ForpaPos, "N") & "," & DBSet(FecVen, "F") & "," & DBSet(Importe, "N") & "," & DBSet(CtaBan, "T") & ","
                If Not vParamAplic.ContabilidadNueva Then
                    CadValues2 = CadValues2 & DBSet(vSocio.Banco, "N", "S") & "," & DBSet(vSocio.Sucursal, "N", "S") & ","
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
                Else
                
                    CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & Text33csb & "," & DBSet(Text41csb, "T") & ",1"
                    
                    vvIban = MiFormat(vSocio.Iban, "") & MiFormat(vSocio.Banco, "0000") & MiFormat(vSocio.Sucursal, "0000") & MiFormat(CC, "00") & MiFormat(vSocio.CuentaBan, "0000000000")
                    
                    CadValues2 = CadValues2 & "," & DBSet(vvIban, "T") & ","
                    CadValues2 = CadValues2 & DBSet(vSocio.Nombre, "T") & "," & DBSet(vSocio.Direccion, "T") & "," & DBSet(vSocio.Poblacion, "T") & ","
                    CadValues2 = CadValues2 & DBSet(vSocio.CPostal, "T") & "," & DBSet(vSocio.Provincia, "T") & "," & DBSet(vSocio.nif, "T") & ",'ES') "
                    
                    Sql = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, "
                    Sql = Sql & "ctabanc1, fecultco, impcobro, "
                    Sql = Sql & " text33csb, text41csb, agente,iban, " ') "
                    Sql = Sql & "nomclien, domclien, pobclien, cpclien, proclien, nifclien, codpais) "
                
                
                End If
                
                Sql = Sql & " VALUES " & CadValues2
                ConnConta.Execute Sql
            
            Else
                '********** si la factura es negativa se inserta en la spago con valor positivo
                CadValues2 = ""
            
            
                CadValuesAux2 = "("
                If vParamAplic.ContabilidadNueva Then CadValuesAux2 = CadValuesAux2 & "'" & SerieFraPro & "',"
                CadValuesAux2 = CadValuesAux2 & "'" & vSocio.CtaProv & "', " & DBSet(numfactu, "N") & ", '" & Format(fecfactu, FormatoFecha) & "', "
                '------------------------------------------------------------
                
                CC = DBLet(vSocio.Digcontrol, "T")
                If DBLet(vSocio.Digcontrol, "T") = "**" Then CC = "00"
                
                i = 1
                CadValues2 = CadValuesAux2 & i
                CadValues2 = CadValues2 & ", " & DBSet(ForpaNeg, "N") & ", " & DBSet(FecVen, "F") & ", "
                CadValues2 = CadValues2 & DBSet(DBLet(Importe, "N") * (-1), "N") & ", " & DBSet(CtaBan, "T") & ","
            
                If Not vParamAplic.ContabilidadNueva Then
                    'David. Para que ponga la cuenta bancaria (SI LA tiene)
                    CadValues2 = CadValues2 & DBSet(vSocio.Banco, "T", "S") & "," & DBSet(vSocio.Sucursal, "T", "S") & ","
                    CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(vSocio.CuentaBan, "T", "S") & ","
                End If
                
                'David. JUNIO 07.   Los dos textos de grabacion de datos de csb
                Select Case Tipo
                    Case 0
                        Sql = "Regularización de Aportación"
                    Case 1
                        Sql = "Aportación de Alta Socio"
                    Case 2
                        Sql = "Devolución Aportación Baja Socio"
                    Case 3
                        Sql = "Devolución Capital Social"
                End Select
                    
                CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "',"
                
                Sql = " de " & Format(DBLet(fecfactu, "F"), "dd/mm/yyyy")
                CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "'" ')"


                If vParamAplic.ContabilidadNueva Then
                    vvIban = MiFormat(vSocio.Iban, "") & MiFormat(CStr(vSocio.Banco), "0000") & MiFormat(CStr(vSocio.Sucursal), "0000") & MiFormat(CC, "00") & MiFormat(vSocio.CuentaBan, "0000000000")
                    
                    CadValues2 = CadValues2 & ", " & DBSet(vvIban, "T") & ","
                    'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
                    CadValues2 = CadValues2 & DBSet(vSocio.Nombre, "T") & "," & DBSet(vSocio.Direccion, "T") & "," & DBSet(vSocio.Poblacion, "T") & "," & DBSet(vSocio.CPostal, "T") & ","
                    CadValues2 = CadValues2 & DBSet(vSocio.Provincia, "T") & "," & DBSet(vSocio.nif, "T") & ",'ES') "
                Else
                    '[Monica]22/11/2013: Tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                        CadValues2 = CadValues2 & ", " & DBSet(vSocio.Iban, "T", "S") & ") "
                    Else
                        CadValues2 = CadValues2 & ") "
                    End If
                End If
                
                'Grabar tabla spagop de la CONTABILIDAD
                '-------------------------------------------------
                If CadValues2 <> "" Then
                    If vParamAplic.ContabilidadNueva Then
                        Sql = "INSERT INTO pagos (numserie, codmacta, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,text1csb,text2csb, iban,"
                        Sql = Sql & "nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais)"
                        
                    Else
                        'Insertamos en la tabla spagop de la CONTA
                        'David. Cuenta bancaria y descripcion textos
                        Sql = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,entidad,oficina,cc,cuentaba,text1csb,text2csb" ' ) "
                        '[Monica]22/11/2013: Tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            Sql = Sql & ", iban) "
                        Else
                            Sql = Sql & ") "
                        End If
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
Dim Rs As ADODB.Recordset
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
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    b = True

    While Not Rs.EOF And b
        If Opcion < 4 Then
            Sql = SQLcuentas & " AND codmacta= " & DBSet(Rs!Codmacta, "T")
        End If

        If Not (RegistrosAListar(Sql, cConta) > 0) Then
        'si no lo encuentra
            b = False 'no encontrado
            If Opcion = 1 Or Opcion = 2 Then
                Sql = DBLet(Rs!Codmacta, "T") & " del Socio " & Format(Rs!Codsocio, "000000")
            End If
        End If

        Rs.MoveNext
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


Private Function IntegracionAportacionesTesoreria(tabla As String, vWhere As String)
Dim Sql As String
Dim Rs As ADODB.Recordset
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
    LOG.Insertar 3, vUsu, "Int.Tesoreria Aportaciones: " & vbCrLf & tabla & vbCrLf & vWhere
    Set LOG = Nothing
    '-----------------------------------------------------------------------------
        
    conn.BeginTrans
    ConnConta.BeginTrans

    Sql = "select raporhco.codsocio, sum(impaport) as importe from " & tabla
    If vWhere <> "" Then Sql = Sql & " WHERE " & vWhere
    Sql = Sql & " group by 1 "
    Sql = Sql & " order by 1 "
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    b = True

    Pb4.visible = True
    Pb4.Max = TotalRegistrosConsulta(Sql)
    Pb4.Value = 0
    
    While Not Rs.EOF And b
        IncrementarProgresNew Pb4, 1
    
        MensError = "Insertando cobro en tesoreria" & vbCrLf & vbCrLf
        b = InsertarEnTesoreriaAPOQua(MensError, Rs!Codsocio, DBLet(Rs!Importe, "N"))
    
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    If b Then
        MensError = "Actualizando Aportaciones" & vbCrLf & vbCrLf
        b = ActualizarAportaciones(MensError, tabla, vWhere)
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
Dim i As Byte
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
Dim vvIban As String

    On Error GoTo EInsertarTesoreriaNew

    b = False
    InsertarEnTesoreriaAPOQua = False
    CadValues = ""
    CadValues2 = ""
    
    Seccion = vParamAplic.Seccionhorto
    
    
    Set vSocio = New cSocio
    If vSocio.LeerDatos(CStr(Socio)) Then
        If vSocio.LeerDatosSeccion(CStr(Socio), CStr(Seccion)) Then
            FecVen = txtCodigo(34).Text
            ForpaNeg = txtCodigo(40).Text
            ForpaPos = txtCodigo(42).Text
            CtaBan = txtCodigo(33).Text
            fecfactu = txtCodigo(49).Text
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
                CadValues2 = CadValues2 & DBSet(CtaBan, "T") & ","
                
                If Not vParamAplic.ContabilidadNueva Then
                    CadValues2 = CadValues2 & DBSet(vSocio.Banco, "N", "S") & "," & DBSet(vSocio.Sucursal, "N", "S") & ","
                    CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(vSocio.CuentaBan, "T", "S") & ","
                    CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
                    CadValues2 = CadValues2 & Text33csb & "," & DBSet(Text41csb, "T") & ",1" ')"
                    
                    '[Monica]22/11/2013: Tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                        CadValues2 = CadValues2 & ", " & DBSet(vSocio.Iban, "T", "S") & ") "
                    Else
                        CadValues2 = CadValues2 & ") "
                    End If
                Else
                    CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & Text33csb & "," & DBSet(Text41csb, "T") & ",1"
                    
                    vvIban = MiFormat(vSocio.Iban, "") & MiFormat(vSocio.Banco, "0000") & MiFormat(vSocio.Sucursal, "0000") & MiFormat(CC, "00") & MiFormat(vSocio.CuentaBan, "0000000000")
                    
                    CadValues2 = CadValues2 & "," & DBSet(vvIban, "T") & ","
                    CadValues2 = CadValues2 & DBSet(vSocio.Nombre, "T") & "," & DBSet(vSocio.Direccion, "T") & "," & DBSet(vSocio.Poblacion, "T") & ","
                    CadValues2 = CadValues2 & DBSet(vSocio.CPostal, "T") & "," & DBSet(vSocio.Provincia, "T") & "," & DBSet(vSocio.nif, "T") & ",'ES') "
                End If
                
                'Insertamos en la tabla scobro de la CONTA
                If vParamAplic.ContabilidadNueva Then
                    Sql = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, "
                    Sql = Sql & "ctabanc1, fecultco, impcobro, "
                    Sql = Sql & " text33csb, text41csb, agente,iban, " ') "
                    Sql = Sql & "nomclien, domclien, pobclien, cpclien, proclien, nifclien, codpais) "
                Else
                    Sql = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci, "
                    Sql = Sql & "ctabanc1, codbanco, codsucur, digcontr, cuentaba, ctabanc2, fecultco, impcobro, "
                    Sql = Sql & " text33csb, text41csb, agente" ') "
                    '[Monica]22/11/2013: Tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                        Sql = Sql & ", iban) "
                    Else
                        Sql = Sql & ") "
                    End If
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
    
                CadValuesAux2 = "("
                If vParamAplic.ContabilidadNueva Then CadValuesAux2 = CadValuesAux2 & "'" & SerieFraPro & "',"
    
                CadValuesAux2 = CadValuesAux2 & DBSet(vSocio.CtaClien, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(fecfactu, "F") & ", 9,"
                CadValues2 = CadValuesAux2 & DBSet(ForpaNeg, "N") & "," & DBSet(FecVen, "F") & "," & DBSet(Importe, "N") & ","
                CadValues2 = CadValues2 & DBSet(CtaBan, "T") & ","
                If Not vParamAplic.ContabilidadNueva Then
                    CadValues2 = CadValues2 & DBSet(vSocio.Banco, "N", "S") & "," & DBSet(vSocio.Sucursal, "N", "S") & ","
                    CadValues2 = CadValues2 & DBSet(CC, "T", "S") & "," & DBSet(vSocio.CuentaBan, "T", "S") & ","
                    CadValues2 = CadValues2 & Text1csb & "," & DBSet(Text2csb, "T") '& ")"
                    
                    '[Monica]22/11/2013: Tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                        CadValues2 = CadValues2 & ", " & DBSet(vSocio.Iban, "T", "S") & ") "
                    Else
                        CadValues2 = CadValues2 & ") "
                    End If
                Else
                    CadValues2 = CadValues2 & Text1csb & "," & DBSet(Text2csb, "T")
                    
                    vvIban = MiFormat(vSocio.Iban, "") & MiFormat(CStr(vSocio.Banco), "0000") & MiFormat(CStr(vSocio.Sucursal), "0000") & MiFormat(CC, "00") & MiFormat(vSocio.CuentaBan, "0000000000")
                    
                    CadValues2 = CadValues2 & ", " & DBSet(vvIban, "T") & ","
                    'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
                    CadValues2 = CadValues2 & DBSet(vSocio.Nombre, "T") & "," & DBSet(vSocio.Direccion, "T") & "," & DBSet(vSocio.Poblacion, "T") & "," & DBSet(vSocio.CPostal, "T") & ","
                    CadValues2 = CadValues2 & DBSet(vSocio.Provincia, "T") & "," & DBSet(vSocio.nif, "T") & ",'ES') "
                
                End If
                
                If vParamAplic.ContabilidadNueva Then
                    Sql = "INSERT INTO pagos (numserie, codmacta, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,text1csb,text2csb, iban,"
                    Sql = Sql & "nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais)"
                
                Else
                    'Insertamos en la tabla scobro de la CONTA
                    Sql = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,entidad,oficina,cc,cuentaba,text1csb,text2csb" ') "
                    '[Monica]22/11/2013: Tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                        Sql = Sql & ", iban) "
                    Else
                        Sql = Sql & ") "
                    End If
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
Dim Rs As ADODB.Recordset
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

Private Function BorradoMasivoAporQua(tabla As String, vWhere As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim cadWhere As String
Dim Mens As String
Dim Nregs As Long

    On Error GoTo eBorradoMasivoAporQua

    BorradoMasivoAporQua = False

    Sql = "select raporhco.* from " & tabla
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
    Nregs = TotalRegistrosConsulta("select raporhco.* from " & tabla & cadWhere)
    
    If MsgBox("Va a eliminar " & Nregs & " registros no contabilizados. ¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Function
    
    conn.Execute Sql2 & cadWhere
   
    BorradoMasivoAporQua = True
    Exit Function
   
eBorradoMasivoAporQua:
    
End Function


Private Function CargarTablaTemporal2(nTabla1 As String, nSelect1 As String, Precio1 As String, ByRef Pb1 As ProgressBar) As Boolean
Dim Rs As ADODB.Recordset
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
Dim Rs As ADODB.Recordset
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
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Sql2 = "insert into raportacion (codsocio,fecaport,codaport,descripcion,campanya,kilos,importe) values "

    Campanya = Mid(Format(Year(CDate(vParam.FecIniCam)), "0000"), 3, 2) & "/" & Mid(Format(Year(CDate(vParam.FecFinCam)), "0000"), 3, 2)
    Descripc = "ACUMULADA " & Campanya

    b = True

    Pb6.visible = True
    Pb6.Max = TotalRegistrosConsulta(Sql)
    Pb6.Value = 0
    
    Fecha = vParam.FecIniCam 'DateAdd("d", (-1), vParam.FecIniCam)
    
    While Not Rs.EOF And b
        IncrementarProgresNew Pb6, 1
    
        SqlValues = ""
        
        Importe = Round2(DBLet(Rs!importe1, "N") * Precio, 2)
    
        SqlExiste = "select count(*) from raportacion where codsocio = " & DBSet(Rs!Codigo1, "N") & " and codaport=0 and fecaport=" & DBSet(Fecha, "F")
        b = (TotalRegistros(SqlExiste) = 0)
        
        If Not b Then
            MsgBox "El registro para el socio " & Format(DBLet(Rs!Codigo1, "N"), "000000") & " de fecha " & DBSet(Fecha, "F") & " y tipo 0 existe. Revise.", vbExclamation
        Else
            SqlValues = "(" & DBSet(Rs!Codigo1, "N") & "," & DBSet(Fecha, "F") & ",0," & DBSet(Descripc, "T") & ","
            SqlValues = SqlValues & DBSet(Campanya, "T") & "," & DBSet(Rs!importe1, "N") & "," & DBSet(Importe, "N") & ")"
            
            conn.Execute Sql2 & SqlValues
            
            MensError = "Insertando cobro en tesoreria"
            b = InsertarEnTesoreriaNewAPO(MensError, Rs!Codigo1, DBLet(Importe, "N"), txtCodigo(51).Text, txtCodigo(52).Text, txtCodigo(53).Text, txtCodigo(50).Text, CStr(Fecha), 1)
            If Not b Then
                MsgBox "Error: " & MensError, vbExclamation
            End If
            
        End If
    
    
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
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



Private Function CargarTablaTemporal3(nTabla1 As String, nSelect1 As String, Precio1 As String, ByRef Pb1 As ProgressBar) As Boolean
Dim Rs As ADODB.Recordset
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
Dim SQLinsert As String

Dim rs3 As ADODB.Recordset
Dim Sql3 As String
Dim CadValues As String


    On Error GoTo eCargarTablaTemporal

    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    SQLinsert = "insert into tmpinformes (codusu, codigo1, nombre1, importe1) values "
    
    Sql2 = "select " & vUsu.Codigo & ", rsocios.codsocio, rsocios.nomsocio from rsocios "
    If nSelect1 <> "" Then Sql2 = Sql2 & " where  " & nSelect1
    Sql2 = Sql2 & " order by 1,2"
    
    CadValues = ""
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Sql3 = "select importe from raportacion where codaport = 0 and codsocio = " & DBSet(Rs!Codsocio, "N")
        Sql3 = Sql3 & " and fecaport in (select max(fecaport) from raportacion where codaport = 0 and codsocio = " & DBSet(Rs!Codsocio, "N") & ")"
        
        Set rs3 = New ADODB.Recordset
        rs3.Open Sql3, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            CadValues = CadValues & "(" & vUsu.Codigo & "," & DBSet(Rs!Codsocio, "N") & "," & DBSet(Rs!nomsocio, "T") & "," & DBSet(rs3!Importe * (-1), "N") & "),"
        End If
        Set rs3 = Nothing
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    If CadValues <> "" Then
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        conn.Execute SQLinsert & CadValues
    End If
    
    CargarTablaTemporal3 = True
    
    Exit Function
    
eCargarTablaTemporal:
    MuestraError Err.Number, "Cargar Tabla Temporal", Err.Description
End Function



Private Function ActualizarRegularizacionBajaSocio()
Dim Sql As String
Dim Rs As ADODB.Recordset
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
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Sql2 = "insert into raportacion (codsocio,fecaport,codaport,descripcion,campanya,kilos,importe) values "

    Campanya = Mid(Format(Year(CDate(vParam.FecIniCam)), "0000"), 3, 2) & "/" & Mid(Format(Year(CDate(vParam.FecFinCam)), "0000"), 3, 2)
    Descripc = "BAJA SOCIO"

    b = True

    Pb7.visible = True
    Pb7.Max = TotalRegistrosConsulta(Sql)
    Pb7.Value = 0
    
    Fecha = txtCodigo(54).Text
    
    While Not Rs.EOF And b
        IncrementarProgresNew Pb7, 1
    
        SqlValues = ""
        
        Importe = DBLet(Rs!importe1, "N")
    
        SqlExiste = "select count(*) from raportacion where codsocio = " & DBSet(Rs!Codigo1, "N") & " and codaport=9 and fecaport=" & DBSet(Fecha, "F")
        b = (TotalRegistros(SqlExiste) = 0)
        
        If Not b Then
            MsgBox "El registro para el socio " & Format(DBLet(Rs!Codigo1, "N"), "000000") & " de fecha " & DBSet(Fecha, "F") & " y tipo 0 existe. Revise.", vbExclamation
        Else
            SqlValues = "(" & DBSet(Rs!Codigo1, "N") & "," & DBSet(Fecha, "F") & ",9," & DBSet(Descripc, "T") & ","
            SqlValues = SqlValues & DBSet(Campanya, "T") & ",0," & DBSet(Importe, "N") & ")"
            
            conn.Execute Sql2 & SqlValues
            
            MensError = "Insertando pago en tesoreria"
            b = InsertarEnTesoreriaNewAPO(MensError, Rs!Codigo1, DBLet(Importe, "N"), txtCodigo(57).Text, txtCodigo(55).Text, txtCodigo(56).Text, txtCodigo(58).Text, CStr(Fecha), 2)
            If Not b Then
                MsgBox "Error: " & MensError, vbExclamation
            End If
            
        End If
    
    
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
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
Dim Rs As ADODB.Recordset
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
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Set ItmX = ListView1(Index).ListItems.Add
        ItmX.Text = Rs.Fields(1).Value ' Format(Rs.Fields(0).Value)
        ItmX.Key = Rs.Fields(0).Value
'        ItmX.SubItems(1) = Rs.Fields(1).Value
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing

ECargarList:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Cargar Tipo de Facturas.", Err.Description
    End If
End Sub


Private Function InsertarAportacionesBolbaite(tabla As String, vWhere As String)
Dim Sql As String
Dim Rs As ADODB.Recordset
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
    LOG.Insertar 3, vUsu, "Insertar Aportaciones: " & vbCrLf & tabla & vbCrLf & vWhere
    Set LOG = Nothing
    '-----------------------------------------------------------------------------
        
    conn.BeginTrans

    Sql = "select * from " & tabla
    If vWhere <> "" Then Sql = Sql & " WHERE " & vWhere
    Sql = Sql & " order by codtipom, numfactu, fecfactu "
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    b = True


    SqlValues = ""

    Pb8.visible = True
    Pb8.Max = TotalRegistrosConsulta(Sql)
    Pb8.Value = 0
    
   
    Sql = "insert into raportacion (codsocio,fecaport,codaport,descripcion,campanya,kilos,importe,codtipom,numfactu,intconta,porcaport) values "
    
    While Not Rs.EOF And b
        IncrementarProgresNew Pb8, 1
    
        Sql2 = "select * from raportacion where fecaport = " & DBSet(Rs!fecfactu, "F")
        Sql2 = Sql2 & " and codtipom = " & DBSet(Rs!CodTipom, "T")
        Sql2 = Sql2 & " and numfactu = " & DBSet(Rs!numfactu, "N")
        
        If TotalRegistros(Sql2) = 0 Then
            Importe = Round2(DBLet(Rs!BaseReten) * ImporteSinFormato(ComprobarCero(txtCodigo(69).Text)) / 100, 2)
        
            SqlValues = SqlValues & "(" & DBSet(Rs!Codsocio, "N") & "," & DBSet(Rs!fecfactu, "F") & "," & DBSet(txtCodigo(68).Text, "N") & ","
            SqlValues = SqlValues & DBSet(txtCodigo(63).Text, "T") & ",' ',0," & DBSet(Importe, "N") & "," & DBSet(Rs!CodTipom, "T") & ","
            SqlValues = SqlValues & DBSet(Rs!numfactu, "N") & ",0," & DBSet(txtCodigo(69).Text, "N") & "),"
        End If
    
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
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



Private Function InsertarAportacionesObligatoriasBolbaite(tabla As String, vWhere As String)
Dim Sql As String
Dim Rs As ADODB.Recordset
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
    LOG.Insertar 3, vUsu, "Insertar Aport.Obligatorias: " & vbCrLf & tabla & vbCrLf & vWhere
    Set LOG = Nothing
    '-----------------------------------------------------------------------------
        
    conn.BeginTrans

    Sql = "select * from " & tabla
    If vWhere <> "" Then Sql = Sql & " WHERE " & vWhere
    Sql = Sql & " order by codsocio"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    b = True


    SqlValues = ""

    Pb9.visible = True
    Pb9.Max = TotalRegistrosConsulta(Sql)
    Pb9.Value = 0
    
    Sql = "insert into raportacion (codsocio,fecaport,codaport,descripcion,campanya,kilos,importe,codtipom,numfactu,intconta,porcaport) values "
    
    While Not Rs.EOF And b
        IncrementarProgresNew Pb9, 1
    
        Sql2 = "select * from raportacion where fecaport = " & DBSet(txtCodigo(74).Text, "F")
        Sql2 = Sql2 & " and codsocio = " & DBSet(Rs!Codsocio, "N")
        Sql2 = Sql2 & " and codaport = " & DBSet(txtCodigo(71).Text, "N")
        
        If TotalRegistros(Sql2) = 0 Then
            Importe = ImporteSinFormato(txtCodigo(73).Text)
        
            SqlValues = SqlValues & "(" & DBSet(Rs!Codsocio, "N") & "," & DBSet(txtCodigo(74).Text, "F") & "," & DBSet(txtCodigo(71).Text, "N") & ","
            SqlValues = SqlValues & DBSet(txtCodigo(72).Text, "T") & ",' ',0," & DBSet(Importe, "N") & "," & ValorNulo & ","
            SqlValues = SqlValues & "0,0,0),"
        End If
    
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
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


Private Function IntegracionAportacionesTesoreriaBolbaite(tabla As String, vWhere As String)
Dim Sql As String
Dim Rs As ADODB.Recordset
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
    LOG.Insertar 3, vUsu, "Int.Tesoreria Aportaciones: " & vbCrLf & tabla & vbCrLf & vWhere
    Set LOG = Nothing
    '-----------------------------------------------------------------------------
        
    conn.BeginTrans
    ConnConta.BeginTrans

    Sql = "select * from " & tabla
    If vWhere <> "" Then Sql = Sql & " WHERE " & vWhere
    Sql = Sql & " order by 1 "
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    b = True

    Pb10.visible = True
    Pb10.Max = TotalRegistrosConsulta(Sql)
    Pb10.Value = 0
    
    While Not Rs.EOF And b
        IncrementarProgresNew Pb10, 1
    
        MensError = "Insertando cobro en tesoreria" & vbCrLf & vbCrLf
        b = InsertarEnTesoreriaAPOBol(MensError, Rs)  'Rs!Codsocio, DBLet(Rs!Importe, "N"))
    
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    If b Then
        MensError = "Actualizando Aportaciones" & vbCrLf & vbCrLf
        b = ActualizarAportacionesBol(MensError, tabla, vWhere)
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
Dim Rs As ADODB.Recordset
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


Private Function InsertarEnTesoreriaAPOBol(MenError As String, ByRef Rs As ADODB.Recordset) As Boolean
'Guarda datos de Tesoreria en tablas: conta.scobros
' Tipo: 0 = almazara
'       1 = bodega

Dim b As Boolean
Dim Sql As String, Text33csb As String, Text41csb As String
Dim Sql4 As String
Dim Rs4 As ADODB.Recordset

Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim i As Byte
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
Dim vvIban As String


    On Error GoTo EInsertarTesoreriaNew

    b = False
    InsertarEnTesoreriaAPOBol = False
    CadValues = ""
    CadValues2 = ""
    
    Seccion = vParamAplic.Seccionhorto
    
    Set vSocio = New cSocio
    If vSocio.LeerDatos(CStr(Rs!Codsocio)) Then
        If vSocio.LeerDatosSeccion(CStr(Rs!Codsocio), CStr(Seccion)) Then
            FecVen = txtCodigo(86).Text
            ForpaNeg = txtCodigo(84).Text
            ForpaPos = txtCodigo(85).Text
            CtaBan = txtCodigo(83).Text
            fecfactu = Rs!fecaport
            If Rs!numfactu = 0 Then
                letraser = ""
                letraser = DevuelveDesdeBDNew(cAgro, "usuarios.stipom", "letraser", "codtipom", "APO", "T")
                numfactu = Mid(Format(Year(CDate(fecfactu)), "0000"), 3, 2) & Format(vSocio.Codigo, "000000")
            Else
                letraser = ""
                letraser = DevuelveDesdeBDNew(cAgro, "usuarios.stipom", "letraser", "codtipom", DBLet(Rs!CodTipom), "T")
                numfactu = DBLet(Rs!numfactu, "N")
            End If
            
            Importe = DBLet(Rs!Importe)
            
            If DBLet(Importe, "N") >= 0 Then
                Text33csb = "'Cargo Aportaciones de " & Format(DBLet(fecfactu, "F"), "dd/mm/yy") & "'"
                Text41csb = "de " & DBSet(Importe, "N")
    
                CC = DBLet(vSocio.Digcontrol, "T")
                If DBLet(vSocio.Digcontrol, "T") = "**" Then CC = "00"
    
                CadValuesAux2 = "(" & DBSet(letraser, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(fecfactu, "F") & ", 9," & DBSet(vSocio.CtaProv, "T") & ","
                CadValues2 = CadValuesAux2 & DBSet(ForpaPos, "N") & "," & DBSet(FecVen, "F") & "," & DBSet(Importe, "N") & ","
                CadValues2 = CadValues2 & DBSet(CtaBan, "T")
                
                If Not vParamAplic.ContabilidadNueva Then
                    CadValues2 = CadValues2 & "," & DBSet(vSocio.Banco, "N", "S") & "," & DBSet(vSocio.Sucursal, "N", "S") & ","
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
                Else
                    CadValues2 = CadValues2 & ValorNulo & "," & ValorNulo & "," & Text33csb & "," & DBSet(Text41csb, "T") & ",1"
                    
                    vvIban = MiFormat(vSocio.Iban, "") & MiFormat(vSocio.Banco, "0000") & MiFormat(vSocio.Sucursal, "0000") & MiFormat(CC, "00") & MiFormat(vSocio.CuentaBan, "0000000000")
                    
                    CadValues2 = CadValues2 & "," & DBSet(vvIban, "T") & ","
                    CadValues2 = CadValues2 & DBSet(vSocio.Nombre, "T") & "," & DBSet(vSocio.Direccion, "T") & "," & DBSet(vSocio.Poblacion, "T") & ","
                    CadValues2 = CadValues2 & DBSet(vSocio.CPostal, "T") & "," & DBSet(vSocio.Provincia, "T") & "," & DBSet(vSocio.nif, "T") & ",'ES') "
                
                    Sql = "INSERT INTO cobros (numserie, numfactu, fecfactu, numorden, codmacta, codforpa, fecvenci, impvenci, "
                    Sql = Sql & "ctabanc1, fecultco, impcobro, "
                    Sql = Sql & " text33csb, text41csb, agente,iban, " ') "
                    Sql = Sql & "nomclien, domclien, pobclien, cpclien, proclien, nifclien, codpais) "
                
                
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
    
                CadValuesAux2 = "("
                If vParamAplic.ContabilidadNueva Then CadValuesAux2 = CadValuesAux2 & "'" & SerieFraPro & "',"
    
                CadValuesAux2 = CadValuesAux2 & DBSet(vSocio.CtaProv, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(fecfactu, "F") & ", 9,"
                CadValues2 = CadValuesAux2 & DBSet(ForpaNeg, "N") & "," & DBSet(FecVen, "F") & "," & DBSet(Importe, "N") & ","
                CadValues2 = CadValues2 & DBSet(CtaBan, "T") & ","
                
                If Not vParamAplic.ContabilidadNueva Then
                    CadValues2 = CadValues2 & DBSet(vSocio.Banco, "N", "S") & "," & DBSet(vSocio.Sucursal, "N", "S") & ","
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
                Else
                    CadValues2 = CadValues2 & Text1csb & "," & DBSet(Text2csb, "T")
                
                    vvIban = MiFormat(vSocio.Iban, "") & MiFormat(CStr(vSocio.Banco), "0000") & MiFormat(CStr(vSocio.Sucursal), "0000") & MiFormat(CC, "00") & MiFormat(vSocio.CuentaBan, "0000000000")
                    
                    CadValues2 = CadValues2 & ", " & DBSet(vvIban, "T") & ","
                    'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
                    CadValues2 = CadValues2 & DBSet(vSocio.Nombre, "T") & "," & DBSet(vSocio.Direccion, "T") & "," & DBSet(vSocio.Poblacion, "T") & "," & DBSet(vSocio.CPostal, "T") & ","
                    CadValues2 = CadValues2 & DBSet(vSocio.Provincia, "T") & "," & DBSet(vSocio.nif, "T") & ",'ES') "
                
                
                    Sql = "INSERT INTO pagos (numserie, codmacta, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,text1csb,text2csb, iban,"
                    Sql = Sql & "nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais)"
                
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



Private Function InsertarDevolucionesQua(vtabla As String, vSelect As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim b As Boolean
Dim CadValues As String
Dim SqlExiste As String
Dim SQLinsert As String
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
    SQLinsert = "insert into ariagro.raporhco (numaport,codsocio,codcampo,poligono,parcela,codparti,codvarie,impaport," & _
                "fecaport,anoplant,observac,supcoope,ejercicio,intconta) values "
    
    Sql = "select raporhco.* from " & vtabla
    Sql = Sql & " where " & vSelect
    
    CargarProgres pb12, TotalRegistrosConsulta(Sql)
    pb12.visible = True
    
    
    CadValues = ""
    CodTipoMov = "APO"
    
    Set vTipoMov = New CTiposMov
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF And b
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
                
        CadValues = "(" & DBSet(NumApor, "N") & "," & DBSet(Rs!Codsocio, "N") & "," & DBSet(Rs!codCampo, "N") & ","
        CadValues = CadValues & DBSet(Rs!Poligono, "N") & "," & DBSet(Rs!Parcela, "N") & "," & DBSet(Rs!codparti, "N") & ","
        CadValues = CadValues & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!ImpAport * (-1), "N") & "," & DBSet(txtCodigo(112).Text, "F") & ","
        CadValues = CadValues & DBSet(Rs!anoplant, "N") & "," & ValorNulo & "," & DBSet(Rs!supcoope, "N") & ","
        CadValues = CadValues & DBSet(txtCodigo(98).Text, "N") & ",0)"
                
        conn.Execute SQLinsert & CadValues
                
        b = vTipoMov.IncrementarContador(CodTipoMov)
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
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


Private Function InsertarTemporalDevolQua(vtabla As String, vSelect As String) As Boolean
Dim Sql As String

    On Error GoTo eInsertarTemporal
    
    InsertarTemporalDevolQua = False
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
                                            'socio,  campo,    variedad,  importe
    Sql = "insert into tmpinformes (codusu, codigo1, importe1, importe2, importe3)"
    Sql = Sql & " select " & vUsu.Codigo & ", rapohco.codsocio, raporhco.codcampo, raporhco.codvarie, sum(coalesce(raporhco.impaport,0)) importe "
    Sql = Sql & " from " & vtabla
    Sql = Sql & " where " & vSelect
    Sql = Sql & " group by 1,2,3,4 "
    Sql = Sql & " order by 1,2,3,4 "
    
    conn.Execute Sql

    InsertarTemporalDevolQua = True
    Exit Function

eInsertarTemporal:
    MuestraError Err.Number, "Insertar Temporal", Err.Description
End Function

