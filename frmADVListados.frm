VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmADVListados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6975
   Icon            =   "frmADVListados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FramePagoPartesADV 
      Height          =   4455
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   6345
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
         Left            =   4875
         TabIndex        =   29
         Top             =   3690
         Width           =   1035
      End
      Begin VB.CommandButton CmdAcepPagoPartes 
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
         Index           =   0
         Left            =   3750
         TabIndex        =   27
         Top             =   3690
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
         Index           =   5
         Left            =   1755
         MaxLength       =   7
         TabIndex        =   23
         Tag             =   "Nº Parte|N|S|||rpartes|nroparte|0000000|S|"
         Top             =   1650
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
         Index           =   4
         Left            =   1755
         MaxLength       =   7
         TabIndex        =   22
         Tag             =   "Nº Parte|N|S|||rpartes|nroparte|0000000|S|"
         Top             =   1260
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
         Index           =   14
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   24
         Top             =   2340
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
         Index           =   15
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   25
         Top             =   2745
         Width           =   1320
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
         Index           =   0
         Left            =   825
         TabIndex        =   34
         Top             =   1320
         Width           =   780
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
         Index           =   1
         Left            =   825
         TabIndex        =   33
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Parte"
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
         Left            =   600
         TabIndex        =   32
         Top             =   990
         Width           =   525
      End
      Begin VB.Label Label3 
         Caption         =   "Pago de Partes ADV"
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
         Left            =   600
         TabIndex        =   31
         Top             =   450
         Width           =   4350
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
         Index           =   3
         Left            =   825
         TabIndex        =   30
         Top             =   2355
         Width           =   645
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
         Index           =   4
         Left            =   825
         TabIndex        =   28
         Top             =   2715
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   5
         Left            =   600
         TabIndex        =   26
         Top             =   2025
         Width           =   600
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1485
         Picture         =   "frmADVListados.frx":000C
         Top             =   2340
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1485
         Picture         =   "frmADVListados.frx":0097
         Top             =   2745
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameAsignacionPrecios 
      Height          =   5415
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   6345
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
         Left            =   2595
         Locked          =   -1  'True
         TabIndex        =   56
         Text            =   "Text5"
         Top             =   1740
         Width           =   3345
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
         Left            =   2595
         Locked          =   -1  'True
         TabIndex        =   55
         Text            =   "Text5"
         Top             =   1320
         Width           =   3345
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
         Left            =   1860
         MaxLength       =   4
         TabIndex        =   37
         Top             =   1725
         Width           =   645
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
         Left            =   1860
         MaxLength       =   4
         TabIndex        =   36
         Top             =   1320
         Width           =   645
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
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   41
         Top             =   3705
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
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   40
         Top             =   3300
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
         Left            =   1860
         MaxLength       =   7
         TabIndex        =   39
         Tag             =   "Nº Parte|N|S|||rpartes|nroparte|0000000|S|"
         Top             =   2700
         Width           =   1335
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
         Left            =   1860
         MaxLength       =   7
         TabIndex        =   38
         Tag             =   "Nº Parte|N|S|||rpartes|nroparte|0000000|S|"
         Top             =   2310
         Width           =   1335
      End
      Begin VB.CommandButton CmdAcepAsigPrecios 
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
         Left            =   3720
         TabIndex        =   42
         Top             =   4770
         Width           =   1035
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
         Left            =   4890
         TabIndex        =   43
         Top             =   4770
         Width           =   1035
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   570
         TabIndex        =   51
         Top             =   4380
         Visible         =   0   'False
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1545
         MouseIcon       =   "frmADVListados.frx":0122
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar Tipo Venta"
         Top             =   1740
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1560
         MouseIcon       =   "frmADVListados.frx":0274
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar Tipo Venta"
         Top             =   1320
         Width           =   240
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
         Index           =   14
         Left            =   795
         TabIndex        =   54
         Top             =   1350
         Width           =   645
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
         Index           =   13
         Left            =   795
         TabIndex        =   53
         Top             =   1710
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Venta"
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
         Left            =   570
         TabIndex        =   52
         Top             =   1020
         Width           =   1395
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   5
         Left            =   1545
         Picture         =   "frmADVListados.frx":03C6
         Top             =   3705
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   4
         Left            =   1545
         Picture         =   "frmADVListados.frx":0451
         Top             =   3300
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   11
         Left            =   570
         TabIndex        =   50
         Top             =   3075
         Width           =   600
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
         Index           =   10
         Left            =   795
         TabIndex        =   49
         Top             =   3675
         Width           =   600
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
         Index           =   9
         Left            =   795
         TabIndex        =   48
         Top             =   3360
         Width           =   645
      End
      Begin VB.Label Label5 
         Caption         =   "Asignación de Precios"
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
         Left            =   600
         TabIndex        =   47
         Top             =   450
         Width           =   4350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Albarán"
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
         Index           =   8
         Left            =   570
         TabIndex        =   46
         Top             =   2055
         Width           =   735
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
         Index           =   7
         Left            =   795
         TabIndex        =   45
         Top             =   2700
         Width           =   600
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
         Index           =   6
         Left            =   795
         TabIndex        =   44
         Top             =   2340
         Width           =   645
      End
   End
   Begin VB.Frame FrameInsercionGastos 
      Height          =   5415
      Left            =   0
      TabIndex        =   57
      Top             =   0
      Width           =   6840
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
         Left            =   5565
         TabIndex        =   65
         Top             =   4770
         Width           =   1035
      End
      Begin VB.CommandButton CmdAcepInsGastos 
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
         Left            =   4305
         TabIndex        =   64
         Top             =   4770
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
         Index           =   19
         Left            =   1890
         MaxLength       =   7
         TabIndex        =   61
         Tag             =   "Nº Parte|N|S|||rpartes|nroparte|0000000|S|"
         Top             =   2610
         Width           =   1155
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
         Index           =   18
         Left            =   1890
         MaxLength       =   7
         TabIndex        =   60
         Tag             =   "Nº Parte|N|S|||rpartes|nroparte|0000000|S|"
         Top             =   2220
         Width           =   1155
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
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   63
         Top             =   3690
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
         Index           =   16
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   62
         Top             =   3300
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
         Index           =   13
         Left            =   1860
         MaxLength       =   16
         TabIndex        =   59
         Top             =   1485
         Width           =   1545
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
         Left            =   3465
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   "Text5"
         Top             =   1485
         Width           =   3150
      End
      Begin MSComctlLib.ProgressBar Pb2 
         Height          =   255
         Left            =   570
         TabIndex        =   66
         Top             =   4380
         Visible         =   0   'False
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
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
         Index           =   23
         Left            =   795
         TabIndex        =   74
         Top             =   2250
         Width           =   915
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
         Index           =   22
         Left            =   795
         TabIndex        =   73
         Top             =   2610
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Albarán"
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
         Left            =   570
         TabIndex        =   72
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Inserción de Gastos"
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
         Left            =   600
         TabIndex        =   71
         Top             =   450
         Width           =   4350
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
         Index           =   20
         Left            =   795
         TabIndex        =   70
         Top             =   3360
         Width           =   645
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
         Index           =   19
         Left            =   795
         TabIndex        =   69
         Top             =   3705
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   18
         Left            =   570
         TabIndex        =   68
         Top             =   3030
         Width           =   600
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   7
         Left            =   1560
         Picture         =   "frmADVListados.frx":04DC
         Top             =   3690
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   6
         Left            =   1560
         Picture         =   "frmADVListados.frx":0567
         Top             =   3300
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Artículo Gasto"
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
         Index           =   17
         Left            =   570
         TabIndex        =   67
         Top             =   1155
         Width           =   1395
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1560
         MouseIcon       =   "frmADVListados.frx":05F2
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar Artículo"
         Top             =   1485
         Width           =   240
      End
   End
   Begin VB.Frame FrameCobros 
      Height          =   7440
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6870
      Begin VB.Frame Frame1 
         Caption         =   "Clasificado por"
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
         Height          =   975
         Left            =   360
         TabIndex        =   85
         Top             =   5310
         Width           =   6075
         Begin VB.OptionButton Option3 
            Caption         =   "Tratamiento"
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
            Left            =   3630
            TabIndex        =   88
            Top             =   450
            Width           =   1665
         End
         Begin VB.OptionButton Option2 
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
            Height          =   255
            Left            =   2190
            TabIndex        =   87
            Top             =   450
            Width           =   1665
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Artículo"
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
            Left            =   480
            TabIndex        =   86
            Top             =   450
            Width           =   1725
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
         Index           =   20
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   81
         Text            =   "Text5"
         Top             =   4380
         Width           =   3780
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
         Index           =   21
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   80
         Text            =   "Text5"
         Top             =   4815
         Width           =   3780
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
         Left            =   1695
         MaxLength       =   6
         TabIndex        =   6
         Top             =   4380
         Width           =   975
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
         Left            =   1695
         MaxLength       =   6
         TabIndex        =   7
         Top             =   4815
         Width           =   975
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
         Index           =   153
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   76
         Text            =   "Text5"
         Top             =   3330
         Width           =   3690
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
         Index           =   154
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   75
         Text            =   "Text5"
         Top             =   3765
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
         Index           =   153
         Left            =   1695
         MaxLength       =   6
         TabIndex        =   4
         Top             =   3330
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
         Index           =   154
         Left            =   1695
         MaxLength       =   6
         TabIndex        =   5
         Top             =   3765
         Width           =   1065
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Sólo resumen"
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
         Left            =   630
         TabIndex        =   20
         Top             =   6450
         Width           =   1965
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
         Left            =   1695
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1680
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
         Left            =   1695
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1275
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
         Left            =   5385
         TabIndex        =   9
         Top             =   6555
         Width           =   1035
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
         Left            =   4215
         TabIndex        =   8
         Top             =   6555
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
         Index           =   0
         Left            =   1695
         MaxLength       =   16
         TabIndex        =   2
         Top             =   2310
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
         Index           =   1
         Left            =   1695
         MaxLength       =   16
         TabIndex        =   3
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
         Index           =   0
         Left            =   3060
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   2310
         Width           =   3450
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
         Left            =   3060
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   2730
         Width           =   3450
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1410
         MouseIcon       =   "frmADVListados.frx":0744
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar tratamiento"
         Top             =   4815
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1410
         MouseIcon       =   "frmADVListados.frx":0896
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar tratamiento"
         Top             =   4380
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tratamiento"
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
         Index           =   24
         Left            =   420
         TabIndex        =   84
         Top             =   4065
         Width           =   1200
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
         Index           =   16
         Left            =   630
         TabIndex        =   83
         Top             =   4785
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
         Index           =   15
         Left            =   630
         TabIndex        =   82
         Top             =   4380
         Width           =   690
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   104
         Left            =   1395
         MouseIcon       =   "frmADVListados.frx":09E8
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   3330
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   105
         Left            =   1395
         MouseIcon       =   "frmADVListados.frx":0B3A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   3765
         Width           =   240
      End
      Begin VB.Label Label2 
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
         Index           =   214
         Left            =   420
         TabIndex        =   79
         Top             =   3045
         Width           =   540
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
         Index           =   220
         Left            =   630
         TabIndex        =   78
         Top             =   3735
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
         Index           =   221
         Left            =   630
         TabIndex        =   77
         Top             =   3330
         Width           =   690
      End
      Begin VB.Label Label1 
         Caption         =   "Informe de Rendimiento "
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
         Left            =   405
         TabIndex        =   19
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
         Left            =   405
         TabIndex        =   18
         Top             =   975
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
         Left            =   630
         TabIndex        =   17
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
         Index           =   14
         Left            =   630
         TabIndex        =   16
         Top             =   1680
         Width           =   645
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1425
         Picture         =   "frmADVListados.frx":0C8C
         ToolTipText     =   "Buscar fecha"
         Top             =   1275
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1425
         Picture         =   "frmADVListados.frx":0D17
         ToolTipText     =   "Buscar fecha"
         Top             =   1680
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
         Left            =   630
         TabIndex        =   15
         Top             =   2310
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
         Index           =   12
         Left            =   630
         TabIndex        =   14
         Top             =   2730
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Artículo"
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
         Left            =   450
         TabIndex        =   13
         Top             =   2025
         Width           =   750
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1410
         MouseIcon       =   "frmADVListados.frx":0DA2
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar artículo"
         Top             =   2310
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1425
         MouseIcon       =   "frmADVListados.frx":0EF4
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar artículo"
         Top             =   2730
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmADVListados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public Opcionlistado As Byte
'0 = Rendimiento por Articulo
'1 = Pago de Partes de trabajo de ADV
'2 = Asignacion de precios en albaranes de Mogente (partes de adv)
'3 = Insercion de gastos en los albaranes de Mogente
    
Private WithEvents frmSoc As frmManSocios 'Socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmArt As frmADVArticulos 'articulos de adv
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents frmTto As frmADVTrataMoi 'tratamientos de Moixent (tipos de venta)
Attribute frmTto.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me

End Sub

Private Sub CmdAcepInsGastos_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim NomAlmac As String
Dim cTabla As String
Dim vHayReg As Byte

    InicializarVbles
    
    ' obligamos a introducir el codigo de articulo de gasto
    If txtNombre(13).Text = "" Then
        MsgBox "Debe de introducir un Artículo de Gasto. Revise.", vbExclamation
        PonerFoco txtCodigo(13)
        Exit Sub
    End If
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    'D/H Parte
    cDesde = Trim(txtCodigo(18).Text)
    cHasta = Trim(txtCodigo(19).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{advpartes.numparte}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHParte=""") Then Exit Sub
    End If
    
    'D/H Fecha
    cDesde = Trim(txtCodigo(16).Text)
    cHasta = Trim(txtCodigo(17).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{advpartes.fechapar}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
    End If

    cTabla = tabla & " INNER JOIN advpartes_lineas on advpartes.numparte = advpartes_lineas.numparte "
    
    'sólo los tratamientos que no sean 0
    If Not AnyadirAFormula(cadselect, "{advpartes.codtrata} <> '0'") Then Exit Sub
    

    vHayReg = 0
    If HayRegParaInforme(cTabla, cadselect) Then
        '[Monica]22/07/2013: comprobamos que en los albaranes que han seleccionado no haya ninguna linea con el articulo de
        '                    gastos. Damos un aviso para que revise
        If HayArticuloGastos(cTabla, cadselect) Then Exit Sub
    
        If ProcesoInsercionGastos(cTabla, cadselect, vHayReg) Then
            If vHayReg = 1 Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                cmdCancel_Click (0)
                Exit Sub
            Else
                MsgBox "No hay datos entre esos límites.", vbExclamation
                cmdCancel_Click (0)
                Exit Sub
            End If
        Else
            MsgBox "El Proceso no se ha realizado correctamente. Llame a Ariadna.", vbExclamation
            Exit Sub
        End If
    End If


End Sub


Private Function HayArticuloGastos(vtabla As String, vWhere As String) As Boolean
Dim Sql As String
Dim cTabla As String
Dim cWhere As String
Dim CadPartes As String
Dim Rs As ADODB.Recordset
Dim Cad As String

    On Error GoTo eHayArticuloGastos


    HayArticuloGastos = False


    cTabla = QuitarCaracterACadena(vtabla, "{")
    cTabla = QuitarCaracterACadena(vtabla, "}")
    Sql = "Select distinct advpartes.numparte FROM " & QuitarCaracterACadena(cTabla, "_1")
    If vWhere <> "" Then
        cWhere = QuitarCaracterACadena(vWhere, "{")
        cWhere = QuitarCaracterACadena(vWhere, "}")
        cWhere = QuitarCaracterACadena(vWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
        Sql = Sql & " and advpartes_lineas.codartic = " & DBSet(txtCodigo(13).Text, "T")
    End If
    Sql = Sql & " order by 1 "
    
    If TotalRegistrosConsulta(Sql) <> 0 Then
        CadPartes = ""
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs.EOF
            CadPartes = CadPartes & DBLet(Rs!Numparte, "N") & ", "
        
            Rs.MoveNext
        Wend
        Set Rs = Nothing
        
        Cad = "Los siguientes albaranes ya tienen un Artículo de Gastos. Revise. "
        Cad = Cad & vbCrLf & vbCrLf
        Cad = Cad & Mid(CadPartes, 1, Len(CadPartes) - 2)
        
        MsgBox Cad, vbExclamation
        
        HayArticuloGastos = True
    End If
    Exit Function
    
eHayArticuloGastos:
    MuestraError Err.Number, "Comprobación de Artículos de gastos", Err.Description
End Function


Private Sub CmdAcepAsigPrecios_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim NomAlmac As String
Dim cTabla As String
Dim vHayReg As Byte

    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    'D/H Parte
    cDesde = Trim(txtCodigo(6).Text)
    cHasta = Trim(txtCodigo(7).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{advpartes.numparte}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHParte=""") Then Exit Sub
    End If
    
    'D/H Fecha
    cDesde = Trim(txtCodigo(8).Text)
    cHasta = Trim(txtCodigo(9).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{advpartes.fechapar}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
    End If

    'D/H Tratamientos
    cDesde = Trim(txtCodigo(10).Text)
    cHasta = Trim(txtCodigo(11).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{advpartes.codtrata}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHTrata=""") Then Exit Sub
    End If


    cTabla = "((" & tabla & " INNER JOIN advpartes_lineas ON advpartes.numparte = advpartes_lineas.numparte) "
    cTabla = cTabla & " INNER JOIN advartic ON advpartes_lineas.codartic = advartic.codartic) "
    cTabla = cTabla & " INNER JOIN advtrata ON advpartes.codtrata = advtrata.codtrata "

    vHayReg = 0
    If HayRegParaInforme(cTabla, cadselect) Then
        If ProcesoAsignacionPrecios(cTabla, cadselect, vHayReg) Then
            If vHayReg = 1 Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                cmdCancel_Click (0)
                Exit Sub
            Else
                MsgBox "No hay datos entre esos límites.", vbExclamation
                cmdCancel_Click (0)
                Exit Sub
            End If
        Else
            MsgBox "El Proceso no se ha realizado correctamente. Llame a Ariadna.", vbExclamation
            Exit Sub
        End If
    End If

End Sub

Private Sub CmdAcepPagoPartes_Click(Index As Integer)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim NomAlmac As String
Dim cTabla As String
Dim vHayReg As Byte

    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    ' Proceso de pago de partes de campo
    NomAlmac = ""
    NomAlmac = DevuelveDesdeBDNew(cAgro, "salmpr", "nomalmac", "codalmac", vParamAplic.AlmacenNOMI, "N")
    If NomAlmac = "" Then
        MsgBox "Debe introducir un código de almacén de Nóminas en parámetros. Revise.", vbExclamation
        Exit Sub
    End If

    'D/H Parte
    cDesde = Trim(txtCodigo(4).Text)
    cHasta = Trim(txtCodigo(5).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{advfacturas_partes.numparte}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
    End If
    
    'D/H Fecha
    cDesde = Trim(txtCodigo(14).Text)
    cHasta = Trim(txtCodigo(15).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{advfacturas_partes.fechapar}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
    End If

    cTabla = tabla & " INNER JOIN advfacturas_trabajador ON advfacturas_partes.codtipom = advfacturas_trabajador.codtipom "
    cTabla = cTabla & " and advfacturas_partes.numfactu = advfacturas_trabajador.numfactu and advfacturas_partes.fecfactu = advfacturas_trabajador.fecfactu "
    cTabla = cTabla & " and advfacturas_partes.numparte = advfacturas_trabajador.numparte "

    vHayReg = 0
    If HayRegParaInforme(cTabla, cadselect) Then
        If ProcesoCargaHoras(cTabla, cadselect, vHayReg) Then
            If vHayReg = 1 Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                cmdCancel_Click (0)
                Exit Sub
            Else
                MsgBox "No hay datos entre esos límites.", vbExclamation
                cmdCancel_Click (0)
                Exit Sub
            End If
        Else
            MsgBox "El Proceso no se ha realizado correctamente. Llame a Ariadna.", vbExclamation
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

InicializarVbles
    
    tabla = "advfacturas_lineas"
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H Articulos
    cDesde = Trim(txtCodigo(0).Text)
    cHasta = Trim(txtCodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codartic}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHArticulo= """) Then Exit Sub
    End If
    
    'D/H Fecha factura
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    'D/H Socio
    cDesde = Trim(txtCodigo(153).Text)
    cHasta = Trim(txtCodigo(154).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{advfacturas.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H Tratamiento
    cDesde = Trim(txtCodigo(20).Text)
    cHasta = Trim(txtCodigo(21).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{advfacturas_partes.codtrata}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHTrata= """) Then Exit Sub
    End If
    
    cadParam = cadParam & "pResumen=" & Check1.Value & "|"
    numParam = numParam + 1
    
    tabla = "(" & tabla & ") inner join advfacturas on advfacturas_lineas.codtipom = advfacturas.codtipom and  advfacturas_lineas.numfactu = advfacturas.numfactu and advfacturas_lineas.fecfactu = advfacturas.fecfactu "
    tabla = "(" & tabla & ") inner join advfacturas_partes on advfacturas_partes.codtipom = advfacturas.codtipom and advfacturas_partes.numfactu = advfacturas.numfactu and advfacturas_partes.fecfactu = advfacturas.fecfactu "
    
    If HayRegistros(tabla, cadselect) Then
        'Nombre fichero .rpt a Imprimir
        If Option1.Value = True Then
            frmImprimir.NombreRPT = "rADVRdtoArticulo.rpt"
            cadTitulo = "Rendimiento por Artículo"
        End If
        If Option2.Value = True Then
            frmImprimir.NombreRPT = "rADVRdtoSocio.rpt"
            cadTitulo = "Rendimiento por Socio"
        End If
        If Option3.Value = True Then
            frmImprimir.NombreRPT = "rADVRdtoTrata.rpt"
            cadTitulo = "Rendimiento por Tratamiento"
        End If
        LlamarImprimir
    End If
End Sub


Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
    End If
    Select Case Opcionlistado
        Case 0 ' rendimiento por articulo
            PonerFoco txtCodigo(2)
        Case 1 ' Proceso de Pago de Partes de adv
            PonerFoco txtCodigo(4)
        Case 2 ' Proceso de asignacion de precios de partes de adv (mogente)
            PonerFoco txtCodigo(10)
        Case 3 ' proceso de insercion de gastos en partes de adv (Mogente)
            PonerFoco txtCodigo(13)
        
    End Select
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection


    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
     For H = 0 To 6
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Next H
     
     For H = 104 To 105
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Next H

         
    indFrame = 5
    FrameCobros.visible = False
    FramePagoPartesADV.visible = False
    Me.FrameAsignacionPrecios.visible = False
    '[Monica]19/07/2013: insercion de gastos en los partes
    Me.FrameInsercionGastos.visible = False
    
    Select Case Opcionlistado
        Case 0 ' rendimiento por articulo
            FrameCobrosVisible True, H, W
            tabla = "advfacturas_lineas"
            Option1.Value = True
        
        Case 1 ' Proceso de Pago de Partes de adv
            FramePagoPartesADVVisible True, H, W
            indFrame = 0
            tabla = "advfacturas_partes"
            
        Case 2 ' Asignacion de precios a albaranes de mogente (partes de adv)
            FrameAsignacionPreciosVisible True, H, W
            indFrame = 0
            tabla = "advpartes"
        
        '[Monica]19/07/2013: insercion de gastos en los partes
        Case 3 ' Insercion de gastos a los albaranes de mogente ( partes de adv )
            FrameInsercionGastosVisible True, H, W
            indFrame = 0
            tabla = "advpartes"
        
        
    End Select
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.CmdCancel(0).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Socios
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTto_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Tratamientos
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
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
            indice = Index + 14
        Case 4, 5
            indice = Index + 4
        Case 6, 7
            indice = Index + 10
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
        Case 0, 1 ' Articulos
            AbrirFrmArticulos (Index)

        Case 4 ' articulo
            AbrirFrmArticulos (Index + 9)

        Case 2, 3 ' Tratatmiento (tipo de venta en moixent)
            AbrirFrmTratamientos (Index + 8)
            
        Case 5, 6 ' tratamiento
            AbrirFrmTratamientos (Index + 15)
        
        Case 104, 105 'socios
            AbrirFrmSocios (Index + 49)
            
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub


Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'14/02/2007 antes
'    KEYpress KeyAscii
    If KeyAscii = teclaBuscar Then
        Select Case Index
            ' informe de rendimiento por calibre
            Case 0: KEYBusqueda KeyAscii, 0 'articulo desde
            Case 1: KEYBusqueda KeyAscii, 1 'articulo hasta
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
            Case 20: KEYBusqueda KeyAscii, 5 'tratamiento desde
            Case 21: KEYBusqueda KeyAscii, 6 'tratamiento hasta
            Case 153: KEYBusqueda KeyAscii, 104 'socio desde
            Case 154: KEYBusqueda KeyAscii, 105 'socio hasta
            
            Case 14: KEYFecha KeyAscii, 0 'fecha desde
            Case 15: KEYFecha KeyAscii, 1 'fecha hasta
            
            Case 8: KEYFecha KeyAscii, 4 'fecha desde
            Case 9: KEYFecha KeyAscii, 5 'fecha hasta
            
            Case 10: KEYBusqueda KeyAscii, 2 'tratamiento desde
            Case 11: KEYBusqueda KeyAscii, 3 'tratamiento hasta
            
            Case 13: KEYBusqueda KeyAscii, 4 'articulo de gastos
            Case 16: KEYFecha KeyAscii, 6 'fecha desde
            Case 17: KEYFecha KeyAscii, 7 'fecha hasta
            
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
            
        Case 0, 1, 13 'articulo
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "advartic", "nomartic", "codartic", "T")
        
        Case 2, 3, 8, 9, 14, 15, 16, 17 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 4, 5, 6, 7 'partes
            If txtCodigo(Index).Text <> "" Then PonerFormatoEntero txtCodigo(Index)
        
        Case 10, 11, 20, 21 'tratamientos
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "advtrata", "nomtrata", "codtrata", "T")
            
        Case 153, 154 ' socios
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
        
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 7440 '4200
        Me.FrameCobros.Width = 6855
        W = Me.FrameCobros.Width
        H = Me.FrameCobros.Height
    End If
End Sub

Private Sub FramePagoPartesADVVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el calculo de horas productivas
    Me.FramePagoPartesADV.visible = visible
    If visible = True Then
        Me.FramePagoPartesADV.Top = -90
        Me.FramePagoPartesADV.Left = 0
        Me.FramePagoPartesADV.Height = 4455
        Me.FramePagoPartesADV.Width = 6870
        W = Me.FramePagoPartesADV.Width
        H = Me.FramePagoPartesADV.Height
    End If
End Sub

Private Sub FrameAsignacionPreciosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el calculo de horas productivas
    Me.FrameAsignacionPrecios.visible = visible
    If visible = True Then
        Me.FrameAsignacionPrecios.Top = -90
        Me.FrameAsignacionPrecios.Left = 0
        Me.FrameAsignacionPrecios.Height = 5415
        Me.FrameAsignacionPrecios.Width = 6345
        W = Me.FrameAsignacionPrecios.Width
        H = Me.FrameAsignacionPrecios.Height
    End If
End Sub


Private Sub FrameInsercionGastosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el calculo de horas productivas
    Me.FrameInsercionGastos.visible = visible
    If visible = True Then
        Me.FrameInsercionGastos.Top = -90
        Me.FrameInsercionGastos.Left = 0
        Me.FrameInsercionGastos.Height = 5415
        Me.FrameInsercionGastos.Width = 6840
        W = Me.FrameInsercionGastos.Width
        H = Me.FrameInsercionGastos.Height
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
        .ConSubInforme = True
        .EnvioEMail = False
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmArticulos(indice As Integer)
    indCodigo = indice
    Set frmArt = New frmADVArticulos
    frmArt.DatosADevolverBusqueda = "0|1|" 'Abrimos en Modo Busqueda
    frmArt.Show vbModal
    Set frmArt = Nothing
End Sub

Private Sub AbrirFrmTratamientos(indice As Integer)
    indCodigo = indice
    Set frmTto = New frmADVTrataMoi
    frmTto.DatosADevolverBusqueda = "0|1|" 'Abrimos en Modo Busqueda
    frmTto.Show vbModal
    Set frmArt = Nothing
End Sub

Private Sub AbrirFrmSocios(indice As Integer)
    indCodigo = indice
    Set frmSoc = New frmManSocios
    frmSoc.DatosADevolverBusqueda = "0|1|"
    frmSoc.Show vbModal
    Set frmSoc = Nothing
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


Private Sub InsertaLineaEnTemporal(ByRef ItmX As ListItem)
Dim Sql As String
Dim Codmacta As String
Dim Rs As ADODB.Recordset
Dim SQL1 As String

        SQL1 = "insert into tmpinformes(codusu, codigo1) values ("
        SQL1 = SQL1 & DBSet(vUsu.Codigo, "N") & "," & DBSet(ItmX.Text, "N") & ")"

        conn.Execute SQL1
    
End Sub

Private Function DatosOk() As Boolean

    DatosOk = True

End Function

Private Function ProcesoCargaHoras(cTabla As String, cWhere As String, vHayReg As Byte) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset

    On Error GoTo eProcesoCargaHoras
    
    Screen.MousePointer = vbHourglass
    
    Sql = "NOMADV" 'nominas de adv
    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se puede realizar el proceso de Carga de Nóminas de ADV. Hay otro usuario realizándolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    ProcesoCargaHoras = False

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "Select advfacturas_trabajador.codtraba, advfacturas_partes.fechapar, sum(advfacturas_trabajador.importel), sum(advfacturas_trabajador.horas) FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    Sql = Sql & " group by 1, 2"
    Sql = Sql & " order by 1, 2"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
    Sql = "insert into horas (codtraba, fechahora, horasdia, horasproduc, compleme,"
    Sql = Sql & "intconta, pasaridoc, codalmac, nroparte) values "
        
    Sql3 = ""
    vHayReg = 0
    While Not Rs.EOF
        Sql2 = "select count(*) from horas where fechahora = " & DBSet(Rs.Fields(1).Value, "F")
        Sql2 = Sql2 & " and codtraba = " & DBSet(Rs.Fields(0).Value, "N")
        Sql2 = Sql2 & " and codalmac = " & DBSet(vParamAplic.AlmacenNOMI, "N")
        
        If TotalRegistros(Sql2) = 0 Then
            vHayReg = 1
            If vParamAplic.Cooperativa = 7 Then
                Sql3 = Sql3 & "(" & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(Rs.Fields(1).Value, "F") & "," & DBSet(Rs.Fields(3).Value, "N") & ",0,"
                Sql3 = Sql3 & "0,0,0," & DBSet(vParamAplic.AlmacenNOMI, "N") & ","
                Sql3 = Sql3 & ValorNulo & "),"
            Else
                Sql3 = Sql3 & "(" & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(Rs.Fields(1).Value, "F") & ",0,0,"
                Sql3 = Sql3 & DBSet(Rs.Fields(2).Value, "N") & ",0,0," & DBSet(vParamAplic.AlmacenNOMI, "N") & ","
                Sql3 = Sql3 & ValorNulo & "),"
            End If
        End If
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    If Sql3 <> "" Then
        ' quitamos la ultima coma
        Sql3 = Mid(Sql3, 1, Len(Sql3) - 1)
        Sql = Sql & Sql3
        
        conn.Execute Sql
    End If
    
    DesBloqueoManual ("CARNOM") 'carga de nominas
    
    Screen.MousePointer = vbDefault
    
    ProcesoCargaHoras = True
    Exit Function
    
eProcesoCargaHoras:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Proceso de Carga de Horas", Err.Description
End Function


Private Function ProcesoAsignacionPrecios(cTabla As String, cWhere As String, vHayReg As Byte) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset
Dim Nregs As Integer
Dim Importe As Currency


    On Error GoTo eProcesoAsignacionPrecios
    
    Screen.MousePointer = vbHourglass
    
    Sql = "PREADV" 'precios de adv
    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se puede realizar el proceso de Asignacion de Precios de ADV. Hay otro usuario realizándolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If


    '[Monica]19/07/2013: insertamos en el log y lo protegemos con transaccion
    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    LOG.Insertar 4, vUsu, "Asignación de Precios de ADV: " & vbCrLf & cTabla & vbCrLf & cWhere
    Set LOG = Nothing
    '-----------------------------------------------------------------------------

    ProcesoAsignacionPrecios = False

    conn.BeginTrans
    
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "Select advpartes_lineas.numparte, advpartes_lineas.numlinea, advpartes_lineas.codartic, advartic.preciove, advtrata.tipoprecio, advpartes_lineas.dosishab bultos, advpartes_lineas.cantidad FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    Sql = Sql & " order by 1, 2"
    
    Nregs = TotalRegistrosConsulta(Sql)
    
    Me.Pb1.visible = True
    CargarProgres Pb1, Nregs
    Me.Refresh
    DoEvents
    
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    vHayReg = 0
    While Not Rs.EOF
        vHayReg = 1
    
        IncrementarProgres Pb1, 1
        Me.Refresh
        DoEvents
        
    
        If DBLet(Rs!TipoPrecio) = 0 Then ' cantidad
            Importe = Round2(DBLet(Rs!preciove) * DBLet(Rs!cantidad), 2)
        Else ' bultos
            Importe = Round2(DBLet(Rs!preciove) * DBLet(Rs!bultos), 2)
        End If
    
        Sql3 = "update advpartes_lineas set importel = " & DBSet(Importe, "N")
        Sql3 = Sql3 & " ,preciove = " & DBSet(Rs!preciove, "N")
        Sql3 = Sql3 & " where numparte = " & DBSet(Rs!Numparte, "N")
        Sql3 = Sql3 & " and numlinea = " & DBSet(Rs!NumLinea, "N")
        
        conn.Execute Sql3
    
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    conn.CommitTrans
    
    DesBloqueoManual ("PREADV") 'Precios de Adv
    
    Screen.MousePointer = vbDefault
    
    ProcesoAsignacionPrecios = True
    Me.Pb1.visible = False
    Exit Function
    
eProcesoAsignacionPrecios:
    conn.RollbackTrans
    Screen.MousePointer = vbDefault
    Me.Pb1.visible = False
    MuestraError Err.Number, "Proceso de Asignacion de Precios", Err.Description
End Function




Private Function ProcesoInsercionGastos(cTabla As String, cWhere As String, vHayReg As Byte) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset
Dim Nregs As Integer
Dim Importe As Currency
Dim Precio As Currency
Dim CodIva As Integer
Dim CadValues As String

    On Error GoTo eProcesoInsercionGastos
    
    Screen.MousePointer = vbHourglass
    
    Sql = "INSADV" 'insercion de gastos de adv
    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se puede realizar el proceso de Inserción de Gastos de ADV. Hay otro usuario realizándolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    LOG.Insertar 4, vUsu, "Insercion de Gastos ADV: " & vbCrLf & cTabla & vbCrLf & cWhere
    Set LOG = Nothing
    '-----------------------------------------------------------------------------


    ProcesoInsercionGastos = False

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "Select advpartes.numparte, max(advpartes_lineas.numlinea) + 1  numlinea, sum(advpartes_lineas.dosishab) bultos FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    Sql = Sql & " group by 1 "
    Sql = Sql & " order by 1 "
    
    Nregs = TotalRegistrosConsulta(Sql)
    
    Me.Pb2.visible = True
    CargarProgres Pb2, Nregs
    Me.Refresh
    DoEvents
    
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    Precio = DevuelveDesdeBDNew(cAgro, "advartic", "preciove", "codartic", txtCodigo(13).Text, "T")
    CodIva = DevuelveDesdeBDNew(cAgro, "advartic", "codigiva", "codartic", txtCodigo(13).Text, "T")
        
    Sql3 = "insert into advpartes_lineas (numparte,numlinea,codalmac,codartic,dosishab,cantidad,preciove,importel,ampliaci,codigiva) values "
        
    vHayReg = 0
    CadValues = ""
    While Not Rs.EOF
        vHayReg = 1
    
        IncrementarProgres Pb2, 1
        DoEvents
    
        Importe = Round2(Precio * DBLet(Rs!bultos), 2)
    
        CadValues = CadValues & "(" & DBSet(Rs!Numparte, "N") & "," & DBSet(Rs!NumLinea, "N") & "," & vParamAplic.AlmacenADV & ","
        CadValues = CadValues & DBSet(txtCodigo(13).Text, "T") & "," & DBSet(Rs!bultos, "N") & "," & DBSet(Rs!bultos, "N") & ","
        CadValues = CadValues & DBSet(Precio, "N") & "," & DBSet(Importe, "N") & "," & ValorNulo & ","
        CadValues = CadValues & DBSet(CodIva, "N") & "),"
    
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    If CadValues <> "" Then
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
    
        conn.Execute Sql3 & CadValues
    End If
    
    DesBloqueoManual ("INSADV") 'Precios de Adv
    
    Screen.MousePointer = vbDefault
    
    ProcesoInsercionGastos = True
    Me.Pb2.visible = False
    Exit Function
    
eProcesoInsercionGastos:
    Screen.MousePointer = vbDefault
    Me.Pb2.visible = False
    MuestraError Err.Number, "Proceso de Inserción de Gastos", Err.Description
End Function

