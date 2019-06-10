VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTercRecFact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturas de Terceros"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   14985
   Icon            =   "frmTercRecFact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   14985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkVistaPrevia 
      Caption         =   "Vista previa"
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
      Height          =   300
      Left            =   12555
      TabIndex        =   76
      Top             =   225
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   135
      TabIndex        =   73
      Top             =   45
      Width           =   2055
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   74
         Top             =   180
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Pedir Datos"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar Dto"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Generar Factura"
               Object.Tag             =   "2"
               Object.Width           =   1e-4
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameFactura 
      Height          =   5325
      Left            =   9120
      TabIndex        =   16
      Top             =   2820
      Width           =   5685
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   3735
         MaxLength       =   15
         TabIndex        =   47
         Tag             =   "Importe Retencion|N|N|0||scafac|imporete|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   4065
         Width           =   1800
      End
      Begin VB.TextBox Text1 
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
         Left            =   1080
         MaxLength       =   5
         TabIndex        =   46
         Tag             =   "% reten|N|S|0|99.90|scafac|porcereten|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   4065
         Width           =   705
      End
      Begin VB.TextBox Text1 
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
         Left            =   2010
         MaxLength       =   15
         TabIndex        =   45
         Tag             =   "Base Imponible 1|N|N|0||scafac|baseimp1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   4065
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   39
         Tag             =   "Importe IVA 1|N|N|0||scafac|imporiv1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   1440
         Width           =   1800
      End
      Begin VB.TextBox Text1 
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
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   38
         Tag             =   "Base Imponible 3|N|N|0||scafac|baseimp3|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   810
         Width           =   1800
      End
      Begin VB.TextBox Text1 
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
         Left            =   3735
         MaxLength       =   15
         TabIndex        =   36
         Tag             =   "Base Imponible 1|N|N|0||scafac|baseimp1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   405
         Width           =   1800
      End
      Begin VB.TextBox Text1 
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
         Left            =   210
         MaxLength       =   5
         TabIndex        =   34
         Tag             =   "% IVA 3|N|S|0|99|scafac|porciva3|00|N|"
         Text            =   "Text1 7"
         Top             =   3240
         Width           =   675
      End
      Begin VB.TextBox Text1 
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
         Left            =   210
         MaxLength       =   5
         TabIndex        =   33
         Tag             =   "& IVA 2|N|S|0|99|scafac|porciva2|00|N|"
         Text            =   "Text1 7"
         Top             =   2820
         Width           =   675
      End
      Begin VB.TextBox Text1 
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
         Left            =   210
         MaxLength       =   5
         TabIndex        =   32
         Tag             =   "% IVA 1|N|S|0|99|scafac|porciva1|00|N|"
         Text            =   "Text1 7"
         Top             =   2385
         Width           =   675
      End
      Begin VB.TextBox Text1 
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
         Left            =   2040
         MaxLength       =   15
         TabIndex        =   26
         Tag             =   "Base Imponible 1|N|N|0||scafac|baseimp1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2385
         Width           =   1485
      End
      Begin VB.TextBox Text1 
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
         Left            =   1065
         MaxLength       =   5
         TabIndex        =   25
         Tag             =   "% IVA 1|N|S|0|99.90|scafac|porciva1|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   2385
         Width           =   705
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   24
         Tag             =   "Importe IVA 1|N|N|0||scafac|imporiv1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2385
         Width           =   1800
      End
      Begin VB.TextBox Text1 
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
         Left            =   2040
         MaxLength       =   15
         TabIndex        =   23
         Tag             =   "Base Imponible 2 |N|N|0||scafac|baseimp2|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2820
         Width           =   1485
      End
      Begin VB.TextBox Text1 
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
         Left            =   1065
         MaxLength       =   5
         TabIndex        =   22
         Tag             =   "& IVA 2|N|S|0|99.90|scafac|porciva2|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   2820
         Width           =   705
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   21
         Tag             =   "Importe IVA 2|N|N|0||scafac|imporiv2|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2820
         Width           =   1800
      End
      Begin VB.TextBox Text1 
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
         Left            =   2040
         MaxLength       =   15
         TabIndex        =   20
         Tag             =   "Base Imponible 3|N|N|0||scafac|baseimp3|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   3240
         Width           =   1485
      End
      Begin VB.TextBox Text1 
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
         Left            =   1080
         MaxLength       =   5
         TabIndex        =   19
         Tag             =   "% IVA 3|N|S|0|99.90|scafac|porciva3|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   3240
         Width           =   705
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   3720
         MaxLength       =   15
         TabIndex        =   18
         Tag             =   "Importe IVA 3|N|N|0||scafac|imporiv3|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   3240
         Width           =   1800
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   22
         Left            =   3285
         MaxLength       =   15
         TabIndex        =   17
         Tag             =   "Total Factura|N|N|0||scafac|totalfac|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   4710
         Width           =   2280
      End
      Begin VB.Label Label1 
         Caption         =   "% Ret"
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
         Index           =   15
         Left            =   1080
         TabIndex        =   50
         Top             =   3780
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Retención"
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
         Index           =   13
         Left            =   3735
         TabIndex        =   49
         Top             =   3765
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Base Retención"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   12
         Left            =   2010
         TabIndex        =   48
         Top             =   3780
         Width           =   1620
      End
      Begin VB.Line Line3 
         X1              =   1035
         X2              =   5535
         Y1              =   3705
         Y2              =   3705
      End
      Begin VB.Line Line2 
         X1              =   1080
         X2              =   5535
         Y1              =   1245
         Y2              =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Imp.Descuento"
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
         Index           =   10
         Left            =   1125
         TabIndex        =   41
         Top             =   810
         Width           =   1620
      End
      Begin VB.Line Line1 
         X1              =   1080
         X2              =   5535
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
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
         Index           =   7
         Left            =   1125
         TabIndex        =   40
         Top             =   1440
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "Bruto Factura"
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
         Index           =   6
         Left            =   1125
         TabIndex        =   37
         Top             =   465
         Width           =   1740
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
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
         Left            =   210
         TabIndex        =   35
         Top             =   2070
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
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
         Left            =   2040
         TabIndex        =   31
         Top             =   2070
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Importe IVA"
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
         Index           =   33
         Left            =   3720
         TabIndex        =   30
         Top             =   2070
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
         Index           =   36
         Left            =   11880
         TabIndex        =   29
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "TOTAL FACTURA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   39
         Left            =   3330
         TabIndex        =   28
         Top             =   4440
         Width           =   1890
      End
      Begin VB.Label Label1 
         Caption         =   "% IVA"
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
         Index           =   41
         Left            =   1065
         TabIndex        =   27
         Top             =   2070
         Width           =   720
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   7
      Left            =   10380
      MaxLength       =   15
      TabIndex        =   70
      Tag             =   "Base Imponible 2 |N|N|0||scafac|baseimp2|#,###,###,##0.00|N|"
      Text            =   "Text1 7"
      Top             =   3375
      Width           =   1485
   End
   Begin VB.Frame FrameIntro 
      Height          =   1860
      Left            =   135
      TabIndex        =   8
      Top             =   855
      Width           =   14700
      Begin VB.CheckBox Check1 
         Caption         =   "Rectificativa"
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
         Index           =   3
         Left            =   5175
         TabIndex        =   80
         Top             =   855
         Width           =   1980
      End
      Begin VB.TextBox Text1 
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
         Left            =   9165
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Fecha Recepción|F|N|||tcafpc|fecrecep|dd/mm/yyyy|N|"
         Top             =   180
         Width           =   1350
      End
      Begin VB.CheckBox Check1 
         Caption         =   "MostrarTodo "
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
         Index           =   2
         Left            =   5175
         TabIndex        =   72
         Top             =   225
         Width           =   1620
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
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
         Left            =   10095
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   68
         Text            =   "Text2"
         Top             =   1380
         Width           =   4455
      End
      Begin VB.TextBox Text1 
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
         Left            =   9165
         MaxLength       =   4
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1380
         Width           =   870
      End
      Begin VB.TextBox Text1 
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
         Left            =   9165
         MaxLength       =   6
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   585
         Width           =   870
      End
      Begin VB.TextBox Text2 
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
         Index           =   26
         Left            =   10065
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   51
         Text            =   "Text2"
         Top             =   585
         Width           =   4485
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Intracomunitario"
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
         Left            =   5175
         TabIndex        =   43
         Top             =   540
         Width           =   1980
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Tesoreria"
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
         Left            =   3195
         TabIndex        =   42
         Top             =   945
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.TextBox Text2 
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
         Left            =   1575
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   1275
         Width           =   5535
      End
      Begin VB.TextBox Text1 
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
         Left            =   3195
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Recepción|F|N|||tcafpc|fecrecep|dd/mm/yyyy|N|"
         Top             =   585
         Width           =   1485
      End
      Begin VB.TextBox Text2 
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
         Index           =   5
         Left            =   10065
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   13
         Text            =   "Text2"
         Top             =   975
         Width           =   4485
      End
      Begin VB.TextBox Text1 
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
         Left            =   9165
         MaxLength       =   5
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   975
         Width           =   870
      End
      Begin VB.TextBox Text1 
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
         Left            =   550
         MaxLength       =   6
         TabIndex        =   7
         Tag             =   "Cod. Transportista|N|N|0|999|tcafpc|codtrans|000|S|"
         Text            =   "Text1"
         Top             =   1275
         Width           =   915
      End
      Begin VB.TextBox Text1 
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
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Factura|F|N|||tcafpc|fecfactu|dd/mm/yyyy|S|"
         Top             =   585
         Width           =   1395
      End
      Begin VB.TextBox Text1 
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
         Left            =   240
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Nº Factura|T|N|||tcafpc|numfactu||S|"
         Text            =   "Text1 7"
         Top             =   585
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Desde F.Albar "
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
         Index           =   18
         Left            =   7440
         TabIndex        =   79
         Top             =   225
         Width           =   1395
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   8880
         Picture         =   "frmTercRecFact.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   225
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Forma Pago"
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
         Index           =   17
         Left            =   7440
         TabIndex        =   67
         Top             =   1410
         Width           =   1215
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   8880
         ToolTipText     =   "Buscar forma pago"
         Top             =   1410
         Width           =   240
      End
      Begin VB.Label Label1 
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
         Height          =   210
         Index           =   16
         Left            =   7440
         TabIndex        =   52
         Top             =   630
         Width           =   1035
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   8880
         ToolTipText     =   "Buscar Variedad"
         Top             =   600
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   4440
         Picture         =   "frmTercRecFact.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   285
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   2700
         Picture         =   "frmTercRecFact.frx":0122
         ToolTipText     =   "Buscar fecha"
         Top             =   270
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   8880
         ToolTipText     =   "Buscar banco propio"
         Top             =   1005
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   240
         ToolTipText     =   "Buscar socio"
         Top             =   1305
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F.Recepción"
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
         Index           =   3
         Left            =   3195
         TabIndex        =   14
         Top             =   285
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Cta.Prev.Pago"
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
         Index           =   2
         Left            =   7440
         TabIndex        =   12
         Top             =   1005
         Width           =   1485
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
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   1020
         Width           =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "F.Factura"
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
         Index           =   29
         Left            =   1560
         TabIndex        =   10
         Top             =   285
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Factura"
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
         Index           =   28
         Left            =   240
         TabIndex        =   9
         Top             =   285
         Width           =   1095
      End
   End
   Begin VB.Frame FrameAux0 
      Height          =   5325
      Left            =   150
      TabIndex        =   53
      Top             =   2820
      Width           =   8910
      Begin VB.Frame FrameToolAux1 
         Height          =   645
         Left            =   135
         TabIndex        =   77
         Top             =   405
         Width           =   690
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   330
            Index           =   0
            Left            =   135
            TabIndex        =   78
            Top             =   180
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Object.ToolTipText     =   "Nuevo"
                  Object.Tag             =   "2"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Modificar"
                  Object.Tag             =   "2"
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Object.ToolTipText     =   "Eliminar"
                  Object.Tag             =   "2"
               EndProperty
            EndProperty
         End
      End
      Begin VB.TextBox txtaux 
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
         Height          =   290
         Index           =   6
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   69
         Tag             =   "Variedad|N|N|||rhisfruta|codvarie|000000|N|"
         Text            =   "var"
         Top             =   3600
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtaux 
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
         Height          =   290
         Index           =   5
         Left            =   6420
         MaxLength       =   35
         TabIndex        =   66
         Tag             =   "Socio|N|N|||rhisfruta|codsocio|000000||"
         Text            =   "Socio"
         Top             =   3600
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtaux 
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
         Height          =   290
         Index           =   3
         Left            =   4350
         MaxLength       =   35
         TabIndex        =   63
         Tag             =   "Kilos Netos|N|N|||rhisfruta|kilosnet|###,##0|N|"
         Text            =   "Kilosnet"
         Top             =   3600
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox Text2 
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
         Height          =   315
         Index           =   2
         Left            =   1770
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   62
         Text            =   "Text2"
         Top             =   3600
         Width           =   2520
      End
      Begin VB.Frame Frame1 
         Height          =   555
         Index           =   0
         Left            =   180
         TabIndex        =   60
         Top             =   4485
         Width           =   2865
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
            Height          =   240
            Left            =   120
            TabIndex        =   61
            Top             =   180
            Width           =   2655
         End
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
         Left            =   6510
         TabIndex        =   58
         Top             =   4605
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
         Left            =   7710
         TabIndex        =   59
         Top             =   4605
         Width           =   1065
      End
      Begin VB.TextBox txtaux 
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
         Height          =   290
         Index           =   4
         Left            =   5790
         MaxLength       =   35
         TabIndex        =   57
         Tag             =   "Importe|N|N|||rhisfruta|impentrada|###,###0.00||"
         Text            =   "Importe"
         Top             =   3600
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtaux 
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
         Height          =   290
         Index           =   2
         Left            =   5190
         MaxLength       =   6
         TabIndex        =   56
         Tag             =   "Precio Estimado|N|S|||rhisfruta|prestimado|###,##0.0000|N|"
         Text            =   "prec"
         Top             =   3600
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtaux 
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
         Height          =   290
         Index           =   0
         Left            =   150
         MaxLength       =   7
         TabIndex        =   55
         Tag             =   "Num.Albaran|N|N|||rhisfruta|numalbar|0000000|S|"
         Text            =   "Albara"
         Top             =   3600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtaux 
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
         Height          =   290
         Index           =   1
         Left            =   600
         MaxLength       =   10
         TabIndex        =   54
         Tag             =   "Fec.Albara|F|N|||rhisfruta|fecalbar|dd/mm/yyyy|N|"
         Text            =   "Fec.Alb"
         Top             =   3600
         Visible         =   0   'False
         Width           =   390
      End
      Begin MSAdodcLib.Adodc AdoAux 
         Height          =   375
         Index           =   0
         Left            =   3840
         Top             =   705
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
      Begin MSDataGridLib.DataGrid DataGridAux 
         Bindings        =   "frmTercRecFact.frx":01AD
         Height          =   3135
         Index           =   0
         Left            =   135
         TabIndex        =   64
         Top             =   1125
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   5530
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
      Begin VB.Label Label2 
         Caption         =   "Albaranes del Socio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   65
         Top             =   180
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   7080
      Top             =   5640
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
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   14415
      TabIndex        =   75
      Top             =   225
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
   Begin VB.Label Label1 
      Caption         =   "Imp. dto. ppago"
      Height          =   255
      Index           =   8
      Left            =   9450
      TabIndex        =   71
      Top             =   3375
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Operador"
      Height          =   255
      Index           =   1
      Left            =   1845
      TabIndex        =   44
      Top             =   900
      Width           =   735
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1845
      Picture         =   "frmTercRecFact.frx":01C5
      ToolTipText     =   "Buscar trabajador"
      Top             =   1125
      Width           =   240
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnPedirDatos 
         Caption         =   "&Pedir Datos"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnModificarDto 
         Caption         =   "&Modificar Dto"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnGenerarFac 
         Caption         =   "&Generar Factura"
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Shortcut        =   ^I
         Visible         =   0   'False
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmTercRecFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 4019 '?????



'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)


'========== VBLES PRIVADAS ====================
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1

Private WithEvents frmSoc As frmManSocios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmVar As frmComVar  'Form Mto clientes
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmBanPr As frmComBanco 'Mto de Bancos propios
Attribute frmBanPr.VB_VarHelpID = -1
Private WithEvents frmFPa As frmComFpa 'Mto de formas de pago
Attribute frmFPa.VB_VarHelpID = -1
'Private WithEvents frmCtas As frmCtasConta 'Cuentas contables

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

'cadena donde se almacena la WHERE para la seleccion de los albaranes
'marcados para facturar
Dim cadWhere As String

'Cuando vuelve del formulario de ver los albaranes seleccionados hay que volver
'a cargar los datos de los albaranes
'Dim VerAlbaranes As Boolean

Dim PrimeraVez As Boolean
Dim Bloquear As Boolean
Dim Indice As Integer

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer
'-------------------------------------------------------------------------

Private vSocio As cSocio

Dim ModoLineas As Byte
'1.- Afegir,  2.- Modificar,  3.- Borrar,  0.-Passar control a Llínies
Dim NumTabMto As Integer 'Indica quin nº de Tab està en modo Mantenimient

Dim vWhere As String


Dim ModificaDescuento As Boolean



Private Sub Check1_LostFocus(Index As Integer)
    If Index = 1 Then
        If Check1(1).Value = 1 Then
            If vParamAplic.CodIvaIntra = 0 Then
                MsgBox "No tiene asignado el código de Iva Intracomunitario en parámetros. Revise.", vbExclamation
                Check1(1).Value = 0
            End If
        End If
    End If
End Sub

Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
    If Not AdoAux(0).Recordset.EOF Then _
        Me.lblIndicador.Caption = AdoAux(0).Recordset.AbsolutePosition & " de " & AdoAux(0).Recordset.RecordCount
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
'    If VerAlbaranes Then RefrescarAlbaranes
'    VerAlbaranes = False
End Sub

Private Sub Form_Load()
Dim i As Integer
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Pedir Datos
        .Buttons(2).Image = 4   'Ver albaranes
        .Buttons(3).Image = 15   'Generar FActura
'        .Buttons(6).Image = 11   'Salir
    End With
    
    ' ******* si n'hi han llínies *******
    'ICONETS DE LES BARRES ALS TABS DE LLÍNIA
    For i = 0 To ToolAux.Count - 1
        With Me.ToolAux(i)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next i
    ' ***********************************
    
    
    
    'cargar IMAGES de busqueda
    Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgBuscar(2).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgBuscar(3).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgBuscar(4).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    
    Me.FrameFactura.Enabled = False
    
    LimpiarCampos   'Limpia los campos TextBox
'    InicializarListView
   
    '## A mano
    NombreTabla = "rhisfruta" ' albaranes de venta
    
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    CadenaConsulta = "Select * from " & NombreTabla
    CadenaConsulta = CadenaConsulta & " where numalbar=-1"
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    End If
    CargaGrid 0, False
    
    PrimeraVez = False
End Sub



Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me
    'Aqui va el especifico de cada form es
    '### a mano
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    DesBloqueoManual "FACTRA"
    TerminaBloquear
'    DesBloqueoManual ("scaalp")
End Sub


Private Sub frmBanPr_DatoSeleccionado(CadenaSeleccion As String)
    'Form de Mantenimiento de Bancos Propios
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
    Text1(5).Text = Format(Text1(5).Text, "0000")
    Text2(5).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
Dim Indice As Byte
    Indice = CByte(Me.imgFecha(0).Tag)
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmFPa_DatoSeleccionado(CadenaSeleccion As String)
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod forpa
    FormateaCampo Text1(4)
    Text2(4).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom forpa
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
Dim Indice As Byte
    Indice = 3
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Socios
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom socio
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
'    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Proveedor
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            Indice = 3
       
       Case 2 'Bancos Propios
            Indice = 5
            Set frmBanPr = New frmComBanco
            frmBanPr.DatosADevolverBusqueda = "0|1|"
            frmBanPr.Show vbModal
            Set frmBanPr = Nothing
            
       Case 3 'Variedades
            Set frmVar = New frmComVar
            frmVar.DatosADevolverBusqueda = "0|2|"
            frmVar.Show vbModal
            Set frmVar = Nothing
            Indice = 26
       
       Case 4 'Forma de pago
            Set frmFPa = New frmComFpa
            frmFPa.DatosADevolverBusqueda = "0|1|"
            frmFPa.Seccion = CStr(vParamAplic.Seccionhorto)
            frmFPa.Show vbModal
            Set frmFPa = Nothing
            Indice = 4
       
    End Select
    
    PonerFoco Text1(Indice)
'    Screen.MousePointer = vbDefault
End Sub

Private Sub imgFecha_Click(Index As Integer)
Dim Indice As Byte
Dim esq As Long
Dim dalt As Long
Dim menu As Long
Dim obj As Object

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
    
   Set frmF = New frmCal
    
   esq = imgFecha(Index).Left
   dalt = imgFecha(Index).Top
    
   Set obj = imgFecha(Index).Container

   While imgFecha(Index).Parent.Name <> obj.Name
       esq = esq + obj.Left
       dalt = dalt + obj.Top
       Set obj = obj.Container
   Wend
    
   menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

   frmF.Left = esq + imgFecha(Index).Parent.Left + 30
   frmF.Top = dalt + imgFecha(Index).Parent.Top + imgFecha(Index).Height + menu - 40
   
   frmF.NovaData = Now
   
   If Index = 2 Then
        Indice = 27
   Else
        Indice = Index + 1
   End If
   
   Me.imgFecha(0).Tag = Indice
   
   PonerFormatoFecha Text1(Indice)
   If Text1(Indice).Text <> "" Then frmF.NovaData = CDate(Text1(Indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(Indice)
End Sub

Private Sub mnModificarDto_Click()
Dim i As Integer


    If Text1(0).Text = "" Then Exit Sub

    PonerModo 4

    Me.FrameFactura.Enabled = True
    
    ' *** bloqueje la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    ' *********************************
    
    BloquearTxt Text1(6), True
    BloquearTxt Text1(8), False
    
    For i = 9 To 22
        BloquearTxt Text1(i), True
    Next i
    
    lblIndicador.Caption = "MODIFICA DESCUENTO"
    
    Me.FrameFactura.Enabled = True
    
    PonerFoco Text1(8)
 
    
End Sub

Private Sub mnGenerarFac_Click()
    BotonFacturar
    Set vSocio = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnPedirDatos_Click()
    BotonPedirDatos
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub


'Private Sub mnVerAlbaran_Click()
'    BotonVerAlbaranes
'End Sub

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
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

    If Index <> 8 And Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1, 2 'Fecha factura, fecha recepcion
            PonerFormatoFecha Text1(Index)
            If Text1(Index) <> "" Then
                ' No debe existir el número de factura para el proveedor en hco
                If ExisteFacturaEnHco Then
'                    InicializarListView
                End If
            End If
            
        '[Monica]06/06/2019
        Case 27 'Desde fecha de albaran
            PonerFormatoFecha Text1(Index)
            
        Case 3 'Cod Socios
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rsocios", "nomsocio", "codsocio")
                
                ' comprobamos que el socio sea tercero
                If Text2(Index).Text <> "" Then
                    Set vSocio = New cSocio
                    If vSocio.Estercero(Text1(Index).Text) Then
                        ' No debe existir el número de factura para el socio en hco
                        If ExisteFacturaEnHco Then
        '                    InicializarListView
                        Else
                            'comprobamos que no haya nadie recepcionando facturas de ese proveedor
        '                    DesBloqueoManual ("FACTRA")
        '                    If Not BloqueoManual("FACTRA", Text1(3).Text) Then
                            vWhere = "codsocio = " & DBSet(Text1(3).Text, "N")
                            If Text1(26).Text <> "" Then vWhere = vWhere & " and codvarie = " & DBSet(Text1(26).Text, "N")
                            If Not BloqueaRegistro("rhisfruta", vWhere) Then
                                MsgBox "No se puede recepcionar factura de ese socio. Hay otro usuario recepcionando.", vbExclamation
                                BotonPedirDatos
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            Else
                                If LimpiarImportes(vWhere) Then
                                    '--monica:080908
                                    TerminaBloquear
                                    If Not BloqueaRegistro("rhisfruta", vWhere) Then
                                        MsgBox "No se puede recepcionar factura de ese socio. Hay otro usuario recepcionando.", vbExclamation
                                        BotonPedirDatos
                                        Screen.MousePointer = vbDefault
                                        Exit Sub
                                    Else
                                        PonerModo 5
                                        '--
                                        CargarAlbaranes vWhere
                                        
                                        CalcularDatosFactura
                                    End If
                                End If
                            End If
                            
                        End If
                   Else
                        PonerFoco Text1(Index)
                   End If
                   Set vSocio = Nothing
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 5 'Cta Prevista de Pago
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "banpropi", "nombanpr", "codbanpr")
                Text1(Index).Text = Format(Text1(Index).Text, "0000")
            Else
                Text2(Index).Text = ""
            End If
            
            
        Case 26 'Variedad
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "variedades", "nomvarie", "codvarie")
                Text1(Index).Text = Format(Text1(Index).Text, "000000")
            Else
                Text2(Index).Text = ""
            End If
        
        Case 4 'Forma de pago
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "forpago", "nomforpa", "codforpa")
                Text1(Index).Text = Format(Text1(Index).Text, "000")
            Else
                Text2(Index).Text = ""
            End If
            
'            '++monica:080908
'            If Not ExisteFacturaEnHco Then
'                PonerModo 5
''                Me.ListView1.SetFocus
'            End If
    
        Case 8 ' Descuento general de la factura
            If PonerFormatoDecimal(Text1(Index), 1) Then
                CalcularDatosFactura
                lblIndicador.Caption = ""
                BloquearTxt Text1(8), True
                Me.FrameFactura.Enabled = False
                PonerModo 5
            End If
            
    End Select
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte, NumReg As Byte
Dim B As Boolean
On Error GoTo EPonerModo

    Modo = Kmodo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    B = (Modo = 2)
        
    cmdAceptar.visible = (ModoLineas = 2)
    cmdAceptar.Enabled = (ModoLineas = 2)
    cmdCancelar.visible = (ModoLineas = 2)
    cmdCancelar.Enabled = (ModoLineas = 2)
    
'    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
'    'Si estamos en Insertar además limpia los campos Text1
'    'si estamos en modificar bloquea las compos que son clave primaria
'    BloquearText1 Me, Modo
    
    For i = 0 To Text1.Count - 1
        BloquearTxt Text1(i), (Modo <> 3)
    Next i
    
    'Importes siempre bloqueados
    For i = 6 To 25
        BloquearTxt Text1(i), True
    Next i
    'Campo B.Imp y Imp. IVA siempre en azul
    Text1(9).BackColor = &HFFFFC0 'Base imponible
    Text1(19).BackColor = &HFFFFC0 'Total Iva 1
    Text1(20).BackColor = &HFFFFC0 'Iva 2
    Text1(21).BackColor = &HFFFFC0 'IVa 3
    Text1(22).BackColor = &HC0C0FF 'Total factura
    Text1(25).BackColor = &HFFFFC0 'Imp.Retencion
    
    
    '---------------------------------------------
    B = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = B
    Next i
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    
    For i = 0 To txtAux.Count - 1
        BloquearTxt txtAux(i), True
        txtAux(i).visible = False
    Next i
        
    Me.FrameIntro.Enabled = (Modo = 3)
    Me.FrameAux0.Enabled = (Modo = 5)
       
    Text2(2).visible = False
       
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
'    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
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
'Comprobar que los datos del frame de introduccion son correctos antes de cargar datos
Dim vtag As CTag
Dim cad As String
Dim i As Byte
Dim vSeccion As CSeccion

    On Error GoTo EDatosOK
    DatosOk = False
    
    ' deben de introducirse todos los datos del frame
    For i = 0 To 5
        If Text1(i).Text = "" Then
            If Text1(i).Tag <> "" Then
                Set vtag = New CTag
                If vtag.Cargar(Text1(i)) Then
                    cad = vtag.Nombre
                Else
                    cad = "Campo"
                End If
                Set vtag = Nothing
            Else
                cad = "Campo"
                If i = 5 Then cad = "Cta. Prev. Pago"
                If i = 4 Then cad = "Forma de Pago"
            End If
            MsgBox cad & " no puede estar vacio. Reintroduzca", vbExclamation
            PonerModo 3
            PonerFoco Text1(i)
            Exit Function
        End If
    Next i
        
    'comprobar que la fecha de la factura sea anterior a la fecha de recepcion
    If Not EsFechaIgualPosterior(Text1(1).Text, Text1(2).Text, True, "La fecha de recepción debe ser igual o posterior a la fecha de la factura.") Then
        Exit Function
    End If
    
    'Comprobar que la fecha de RECEPCION esta dentro de los ejercicios contables
    Set vSeccion = New CSeccion
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If vSeccion.AbrirConta Then
            i = EsFechaOKConta(CDate(Text1(2).Text))
            If i > 0 Then
                MsgBox "Fecha fuera ejercicios contables", vbExclamation
                vSeccion.CerrarConta
                Set vSeccion = Nothing
                Exit Function
            End If
        End If
    End If
    vSeccion.CerrarConta
    Set vSeccion = Nothing

'    If vParamAplic.NumeroConta <> 0 Then
'        i = EsFechaOKConta(CDate(Text1(2).Text))
'        If i > 0 Then
'            'If i = 1 Then
'                MsgBox "Fecha fuera ejercicios contables", vbExclamation
'                Exit Function
'           ' Else
'           '     cad = "La fecha es superior al ejercico contable siguiente. ¿Desea continuar?"
'           '     If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
'           ' End If
'        End If
'    End If
    
'--monica:03/12/2008
'    'comprobar que se han seleccionado lineas para facturar
'    If cadWHERE = "" Then
'        MsgBox "Debe seleccionar albaranes para facturar.", vbExclamation
'        Exit Function
'    End If
    
'++monica:03/12/2008
    'comprobamos que hay lineas para facturar: o albaranes o portes de vuelta
    If cadWhere = "" Then
        If AdoAux(0).Recordset.RecordCount = 0 Then
            MsgBox "No hay albaranes para incluir en la factura. Revise.", vbExclamation
            Exit Function
        End If
    End If
    
    
    '[Monica]27/10/2016: si el importe de la factura es 0 no dejamos facturar
    Dim vSqlNuevo As String
    vSqlNuevo = "select * from rhisfruta where " & cadWhere
    If TotalRegistrosConsulta(vSqlNuevo) = 0 Then
        MsgBox "No ha introducido importe de los albaranes. Revise.", vbExclamation
        Exit Function
    End If
    
    
    ' No debe existir el número de factura para el socio tercero en hco
    If ExisteFacturaEnHco Then Exit Function
    
'--monica
'    'todos los albaranes seleccionados deben tener la misma: forma pago, dto ppago, dto gnral
'    cad = "select count(distinct codforpa,dtoppago,dtognral) from scaalp "
'    cad = cad & " WHERE " & Replace(cadWHERE, "slialp", "scaalp")
'    If RegistrosAListar(cad) > 1 Then
'        MsgBox "No se puede facturar albaranes con distintas: forma de pago, dto gral, dto ppago.", vbExclamation
'        Exit Function
'    End If
    
    
    'Si la forpa es TRANSFERENCIA entonces compruebo la si tiene cta bancaria
'    cad = "select distinct (codforpa) from scaalp "
'    cad = cad & " WHERE " & Replace(cadWHERE, "slialp", "scaalp")
    Set miRsAux = New ADODB.Recordset
'    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    cad = miRsAux.Fields(0)
'    miRsAux.Close
    
    
    
    'Ahora buscamos el tipforpa del codforpa
    cad = "Select tipoforp from forpago where codforpa=" & DBSet(Text1(4).Text, "N")
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    If miRsAux.EOF Then
        MsgBox "Error en el TIPO de forma de pago", vbExclamation
    Else
        i = 1
        cad = miRsAux.Fields(0)
        If Val(cad) = vbFPTransferencia Then
            'Compruebo que la forpa es transferencia
            i = 2
        End If
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
    
    If i = 2 Then
        'La forma de pago es transferencia. Debo comprobar que existe la cuenta bancaria
        'del proveedor
        If vSocio.CuentaBan = "" Or vSocio.Digcontrol = "" Or vSocio.Sucursal = "" Or vSocio.Banco = "" Then
            cad = "Cuenta bancaria incorrecta. Forma de pago: transferencia.    ¿Continuar?"
            If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then i = 0
        End If
    End If
    
    'Si i=0 es que o esta mal la forpa o no quiere seguir pq no tiene cuenta bancaria
    If i > 0 Then DatosOk = True
    Exit Function
    
EDatosOK:
    DatosOk = False
    MuestraError Err.Number, "Comprobar datos correctos", Err.Description
End Function



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Pedir datos
             mnPedirDatos_Click
            
        Case 2
             mnModificarDto_Click
            
        Case 3 'Generar Factura
            mnGenerarFac_Click

    End Select
End Sub


Private Sub PonerOpcionesMenu()
Dim J As Byte

    PonerOpcionesMenuGeneral Me
    
    J = Val(Me.mnPedirDatos.HelpContextID)
    If J < vUsu.Nivel Then Me.mnPedirDatos.Enabled = False
    
    J = Val(Me.mnGenerarFac.HelpContextID)
    If J < vUsu.Nivel Then Me.mnGenerarFac.Enabled = False
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

 
Private Sub BotonPedirDatos()
Dim Nombre As String

    TerminaBloquear

    'Vaciamos todos los Text
    LimpiarCampos
    Check1(1).Value = 0
    'Vaciamos el ListView
'    InicializarListView
    CargaGrid 0, False
    
    'Como no habrá albaranes seleccionados vaciamos la cadwhere
    cadWhere = ""
    
    PonerModo 3
    
    'fecha recepcion
    Text1(2).Text = Format(Now, "dd/mm/yyyy")
    
    
    '[Monica]06/06/2019: fecha desde albaran, la fecha de inicio de campaña
    Dim FechaAnt As String
    Text1(27).Text = DateAdd("d", 1, DateAdd("yyyy", -1, vParam.FecFinCam))
    FechaAnt = Text1(27).Text
    
    
    
    'desbloquear los registros de la saalp (si hay bloquedos)
    TerminaBloquear
    
    'si vamos
    'desBloqueo Manual de las tablas
'    DesBloqueoManual ("scaalp")
    
    PonerFoco Text1(0)
End Sub


Private Sub CargarAlbaranes(cadWhere As String)
'Recupera de la BD y muestra en el Listview todos los albaranes de compra
'que tiene el proveedor introducido.
Dim SQL As String
Dim Rs As ADODB.Recordset

On Error GoTo ECargar
    
    
    CargaGrid 0, True

    If AdoAux(0).Recordset.RecordCount = 0 Then
        MsgBox "No existen albaranes pendientes de facturar para este socio.", vbExclamation
        BotonPedirDatos
    End If


ECargar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando Albaranes", Err.Description
End Sub



Private Sub CalcularDatosFactura()
Dim i As Integer
Dim SQL As String
Dim cadAux As String
Dim ImpBruto As Currency
Dim impiva As Currency
Dim vFactu As CFacturaTer
Dim Rs As ADODB.Recordset
Dim Dto As Currency

    Dto = 0
    If Text1(8).Text <> "" Then
        Dto = CCur(TransformaPuntosComas(ImporteSinFormato(Text1(8).Text)))
    End If
    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For i = 6 To 25
         Text1(i).Text = ""
    Next i

    cadAux = ""
    cadWhere = ""
    ImpBruto = 0
    
    SQL = "select variedades.codigiva, sum(impentrada) from rhisfruta, variedades where codsocio= " & DBSet(Text1(3).Text, "N")
    If Text1(26).Text <> "" Then
        SQL = SQL & " and rhisfruta.codvarie = " & DBSet(Text1(26).Text, "N")
    End If
    SQL = SQL & " and variedades.codvarie = rhisfruta.codvarie "
    If Check1(2).Value = 0 Then
        SQL = SQL & " and rhisfruta.cobradosn = 0 "
    End If
    SQL = SQL & " group by 1 "
    SQL = SQL & " order by 1 "
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If Not Rs.EOF Then ImpBruto = ImpBruto + DBLet(Rs.Fields(0).Value, "N")
    
    cadWhere = "rhisfruta.codsocio=" & Val(Text1(3).Text)
    If Check1(2).Value = 0 Then
        cadWhere = cadWhere & " and rhisfruta.cobradosn = 0 "
    End If
    cadWhere = cadWhere & " and rhisfruta.impentrada <> 0 "

    If Text1(26).Text <> "" Then
        cadWhere = cadWhere & " and rhisfruta.codvarie = " & DBSet(Text1(26).Text, "N")
    End If

    If Not SeleccionaRegistros Then Exit Sub
    
    If Not BloqueaRegistro("rhisfruta", cadWhere) Then
        conn.Execute "update rhisfruta set impentrada = 0 where " & cadWhere
        cadWhere = "rhisfruta.codsocio=" & Val(Text1(3).Text)
        
        If Check1(2).Value = 0 Then
            cadWhere = cadWhere & "  and rhisfruta.cobradosn = 0 "
        End If
        If Text1(26).Text <> "" Then
            cadWhere = cadWhere & " and rhisfruta.codvarie = " & DBSet(Text1(26).Text, "N")
        End If
        CargarAlbaranes cadWhere
    End If
    
    cadWhere = "rhisfruta.codsocio=" & Val(Text1(3).Text)
    cadWhere = cadWhere & " and rhisfruta.impentrada <> 0 "
    
    If Check1(2).Value = 0 Then
        cadWhere = cadWhere & " and rhisfruta.cobradosn = 0 "
    End If
        
    If Text1(26).Text <> "" Then
        cadWhere = cadWhere & " and rhisfruta.codvarie = " & DBSet(Text1(26).Text, "N")
    End If

    Set vFactu = New CFacturaTer
    vFactu.DtoPPago = 0
    vFactu.DtoGnral = 0
    If Dto <> 0 Then
        vFactu.DtoGnral = Dto
    End If
    vFactu.Intracomunitario = Check1(1).Value
    If vFactu.CalcularDatosFactura(cadWhere, Text1(3).Text) Then
        Text1(6).Text = vFactu.BrutoFac
        Text1(7).Text = vFactu.ImpPPago
        Text1(8).Text = vFactu.ImpGnral
        Text1(9).Text = vFactu.BaseImp
        Text1(10).Text = vFactu.TipoIVA1
        Text1(11).Text = vFactu.TipoIVA2
        Text1(12).Text = vFactu.TipoIVA3
        Text1(13).Text = vFactu.PorceIVA1
        Text1(14).Text = vFactu.PorceIVA2
        Text1(15).Text = vFactu.PorceIVA3
        Text1(16).Text = vFactu.BaseIVA1
        Text1(17).Text = vFactu.BaseIVA2
        Text1(18).Text = vFactu.BaseIVA3
        Text1(19).Text = vFactu.ImpIVA1
        Text1(20).Text = vFactu.ImpIVA2
        Text1(21).Text = vFactu.ImpIVA3
        Text1(22).Text = vFactu.TotalFac
        Text1(23).Text = vFactu.BaseReten
        Text1(25).Text = vFactu.ImpReten
        If vFactu.ImpReten = 0 Then
            Text1(24).Text = 0
        Else
            Text1(24).Text = vFactu.PorcReten
        End If
        
        Check1(1).Value = vFactu.Intracomunitario
        
        For i = 6 To 26
            FormateaCampo Text1(i)
        Next i
        'Quitar ceros de linea IVA 2
        If Val(Text1(14).Text) = 0 And Val(Text1(11).Text) = 0 Then
            For i = 11 To 20 Step 3
                Text1(i).Text = QuitarCero(CCur(Text1(i).Text))
            Next i
        End If
        'Quitar ceros de linea IVA 3
        If Val(Text1(15).Text) = 0 And Val(Text1(12).Text) = 0 Then
            For i = 12 To 21 Step 3
                Text1(i).Text = QuitarCero(CCur(Text1(i).Text))
            Next i
        End If
        
    Else
        MuestraError Err.Number, "Calculando Factura", Err.Description
    End If
    Set vFactu = Nothing
    
    
    
End Sub

Private Function SeleccionaRegistros() As Boolean
'Comprueba que se seleccionan albaranes en la base de datos
'es decir que hay albaranes marcados
'cuando se van marcando albaranes se van añadiendo el la cadena cadWhere
Dim SQL As String

    On Error GoTo ESel
    SeleccionaRegistros = False
    
    If cadWhere = "" Then Exit Function
    
    SQL = "Select count(*) FROM rhisfruta"
    SQL = SQL & " WHERE " & cadWhere
    If RegistrosAListar(SQL) <> 0 Then SeleccionaRegistros = True
    Exit Function
    
ESel:
    SeleccionaRegistros = False
    MuestraError Err.Number, "No hay seleccionados Albaranes", Err.Description
End Function


Private Sub BotonFacturar()
Dim vFactu As CFacturaTer
Dim cad As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    '[Monica]08/01/2018: daba error cuando no habia nada, ahora no hace nada si no hay fecha
    If Text1(2).Text = "" Then Exit Sub
    
    '[Monica]20/06/2017: control de fechas que antes no estaba
    ResultadoFechaContaOK = EsFechaOKConta(CDate(Text1(2)))
    If ResultadoFechaContaOK > 0 Then
        If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
        Exit Sub
    End If
    
    cad = ""
    If Text1(3).Text = "" Then
        cad = "Falta socio"
    Else
        If Not IsNumeric(Text1(3).Text) Then cad = "Campo socio debe ser numérico"
    End If
    If cad <> "" Then
        MsgBox cad, vbExclamation
        Exit Sub
    End If
        
    Set vSocio = New cSocio
    
    'Tiene que ller los datos del transportista
    If Not vSocio.LeerDatos(Text1(3).Text) Then Exit Sub
    
    If Not DatosOk Then
        Set vSocio = Nothing
        Exit Sub
    End If

    'Pasar los Albaranes seleccionados con cadWHERE a una factura
    Set vFactu = New CFacturaTer
    vFactu.Tercero = Text1(3).Text
    vFactu.numfactu = Text1(0).Text
    vFactu.fecfactu = Text1(1).Text
    vFactu.FecRecep = Text1(2).Text
    vFactu.Trabajador = Text1(4).Text
    vFactu.BancoPr = Text1(5).Text
    vFactu.BrutoFac = ImporteFormateado(Text1(6).Text)
    vFactu.ForPago = Text1(4).Text
    vFactu.DtoPPago = 0
    vFactu.DtoGnral = 0
    vFactu.ImpPPago = ImporteFormateado(Text1(7).Text)
    vFactu.ImpGnral = ImporteFormateado(Text1(8).Text)
    vFactu.BaseIVA1 = ImporteFormateado(Text1(16).Text)
    vFactu.BaseIVA2 = ImporteFormateado(Text1(17).Text)
    vFactu.BaseIVA3 = ImporteFormateado(Text1(18).Text)
    vFactu.TipoIVA1 = ComprobarCero(Text1(10).Text)
    vFactu.TipoIVA2 = ComprobarCero(Text1(11).Text)
    vFactu.TipoIVA3 = ComprobarCero(Text1(12).Text)
    vFactu.PorceIVA1 = ComprobarCero(Text1(13).Text)
    vFactu.PorceIVA2 = ComprobarCero(Text1(14).Text)
    vFactu.PorceIVA3 = ComprobarCero(Text1(15).Text)
    vFactu.ImpIVA1 = ImporteFormateado(Text1(19).Text)
    vFactu.ImpIVA2 = ImporteFormateado(Text1(20).Text)
    vFactu.ImpIVA3 = ImporteFormateado(Text1(21).Text)
    vFactu.TotalFac = ImporteFormateado(Text1(22).Text)
    vFactu.PorcReten = ImporteFormateado(Text1(24).Text)
    vFactu.ImpReten = ImporteFormateado(Text1(25).Text)
    vFactu.BaseReten = ImporteFormateado(Text1(23).Text)
    
    'Si el proveedor tiene CTA BANCARIA se la asigno
    vFactu.CCC_Entidad = vSocio.Banco
    vFactu.CCC_Oficina = vSocio.Sucursal
    vFactu.CCC_CC = vSocio.Digcontrol
    vFactu.CCC_CTa = vSocio.CuentaBan
    vFactu.CCC_Iban = vSocio.Iban
    
    vFactu.Intracomunitario = Check1(1).Value
    
    '[Monica]07/06/2019: marcamos que la factura es rectificativa
    vFactu.EsRectificativa = Check1(3).Value
        
    
    
    ' sacamos la cuenta de proveedor
    If Not vSocio.LeerDatosSeccion(vSocio.Codigo, vParamAplic.Seccionhorto) Then
        MsgBox "No se han encontrado los datos del socio de la sección Hortofrutícola", vbExclamation
        Set vFactu = Nothing
        Exit Sub
    End If
    
    vFactu.CtaTerce = vSocio.CtaProv
    
    If cadWhere <> "" Then
        If vFactu.TraspasoAlbaranesAFactura(cadWhere) Then BotonPedirDatos
    End If
    Set vFactu = Nothing
    
    
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Function ExisteFacturaEnHco() As Boolean
'Comprobamos si la factura ya existe en la tabla de Facturas a Proveedor: scafpc
Dim cad As String

    ExisteFacturaEnHco = False
    'Tiene que tener valor los 3 campos de clave primaria antes de comprobar
    If Not (Text1(0).Text <> "" And Text1(1).Text <> "" And Text1(3).Text <> "") Then Exit Function
    
    ' No debe existir el número de factura para el proveedor en hco. [06/05/2013]la fecha a mirar es la de recepcion
    cad = "SELECT count(*) FROM rcafter "
    cad = cad & " WHERE codsocio=" & Text1(3).Text & " AND numfactu=" & DBSet(Text1(0).Text, "T") & " AND year(fecfactu)=" & Year(Text1(2).Text)
    If RegistrosAListar(cad) > 0 Then
        MsgBox "Factura de Tercero ya existente. Reintroduzca.", vbExclamation
        ExisteFacturaEnHco = True
        Exit Function
    End If
End Function



'****************************************

'************* LLINIES: ****************************
Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
Dim i As Integer
'-- pon el bloqueo aqui
    'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
    Select Case Button.Index
        Case 2
            BotonModificarLinea Index
        
    End Select
    'End If
End Sub




Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim i As Integer
    Dim J As Integer
    
    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub
    
    ModoLineas = 2 'Modificar llínia
       
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5
    ' *** bloqueje la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    ' *********************************
  
    Select Case Index
        Case 0 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
            If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
                i = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
                DataGridAux(Index).Scroll 0, i
                DataGridAux(Index).Refresh
            End If
              
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
        
    End Select
    
    Select Case Index
        ' *** valor per defecte al modificar dels camps del grid ***
        Case 0 'importes
            txtAux(0).Text = DataGridAux(Index).Columns(0).Text
            txtAux(1).Text = DataGridAux(Index).Columns(1).Text
            txtAux(6).Text = DataGridAux(Index).Columns(2).Text
            Text2(2).Text = DataGridAux(Index).Columns(3).Text
            txtAux(3).Text = DataGridAux(Index).Columns(4).Text
            txtAux(2).Text = DataGridAux(Index).Columns(5).Text
            txtAux(4).Text = DataGridAux(Index).Columns(6).Text
            txtAux(5).Text = DataGridAux(Index).Columns(7).Text
            
            For i = 0 To 3
                BloquearTxt txtAux(i), True
            Next i
            BloquearTxt txtAux(4), False
            BloquearTxt txtAux(5), True
       
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'importes
            PonerFoco txtAux(4)
    End Select
    ' ***************************************************************************************
    lblIndicador.Caption = "INSERTAR IMPORTE"
End Sub

Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
Dim B As Boolean
Dim i As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    
    Select Case Index
        Case 0 'rhisfruta
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "S|txtaux(0)|T|Albarán|950|;S|txtaux(1)|T|Fecha|1250|;"
            tots = tots & "S|txtaux(6)|T|Código|860|;S|Text2(2)|T|Variedad|1520|;"
            tots = tots & "S|txtaux(3)|T|Kilos Neto|1100|;S|txtaux(2)|T|Pr.Estim.|1050|;"
            tots = tots & "S|txtaux(4)|T|Importe|1300|;N|txtaux(5)|T|Socio|1100|;"
            arregla tots, DataGridAux(Index), Me, 350
        
            DataGridAux(0).Columns(1).Alignment = dbgLeft
            DataGridAux(0).Columns(5).Alignment = dbgRight
        
            B = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            BloquearTxt txtAux(3), Not B

    End Select
    DataGridAux(Index).ScrollBars = dbgAutomatic
    
    If Not AdoAux(0).Recordset.EOF Then
        Me.lblIndicador.Caption = AdoAux(0).Recordset.AbsolutePosition & " de " & AdoAux(0).Recordset.RecordCount
    Else
        Me.lblIndicador.Caption = ""
    End If
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub



Private Function MontaSQLCarga(Index As Integer, enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basant-se en la informació proporcionada pel vector de camps
'   crea un SQl per a executar una consulta sobre la base de datos que els
'   torne.
' Si ENLAZA -> Enlaça en el data1
'           -> Si no el carreguem sense enllaçar a cap camp
'--------------------------------------------------------------------
Dim SQL As String
Dim tabla As String
   
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
        Case 0 'historico de entradas
            tabla = "rhisfruta"
            SQL = "SELECT rhisfruta.numalbar,rhisfruta.fecalbar, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.kilosnet, rhisfruta.prestimado, rhisfruta.impentrada, rhisfruta.codsocio "
            SQL = SQL & " FROM " & tabla & " inner join variedades on rhisfruta.codvarie = variedades.codvarie "
            If enlaza Then
'                SQL = SQL & ObtenerWhereCab(True)
                SQL = SQL & " where codsocio =  " & DBSet(Text1(3).Text, "N")
                
                
                '[Monica] 04/02/2010 Todos los albaranes o solo los que no han sido cobrados
                If Check1(2).Value = 0 Then
                    SQL = SQL & " and cobradosn = 0 "   ' que no esten cobradas
                End If
                    
                If Text1(26).Text <> "" Then
                    SQL = SQL & " and rhisfruta.codvarie = " & DBSet(Text1(26).Text, "N")
                End If
                
                '[Monica]06/06/2019: desde fehca de albaran
                If Text1(27).Text <> "" Then
                    SQL = SQL & " and rhisfruta.fecalbar >= " & DBSet(Text1(27).Text, "F")
                End If
                
            Else
                SQL = SQL & " WHERE numalbar  = -1"
            End If
            
            SQL = SQL & " ORDER BY " & tabla & ".numalbar,  " & tabla & ".fecalbar "
            
    End Select
    ' ********************************************************************************
    
    MontaSQLCarga = SQL
End Function


Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    If Text1(0).Text = "" Then Exit Sub
    Select Case Modo
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 2 'modificar llínies
                    ModoLineas = 0
                    PonerModo 5
                    ModificarLinea
                    If Not AdoAux(0).Recordset.EOF Then _
                        Me.lblIndicador.Caption = AdoAux(0).Recordset.AbsolutePosition & " de " & AdoAux(0).Recordset.RecordCount
            End Select
            
        CalcularDatosFactura
        ' **************************
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
Dim i As Integer
Dim V
    
    Select Case Modo
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 2 'modificar llínies
                    ModoLineas = 0
                    
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de llínies ***
                        V = AdoAux(NumTabMto).Recordset.Fields(0) 'el 2 es el nº de llinia
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(0).Name & " =" & V)
                        ' ***************************************************************
                    End If
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                    PonerModo 2
                    If Not AdoAux(0).Recordset.EOF Then _
                         Me.lblIndicador.Caption = AdoAux(0).Recordset.AbsolutePosition & " de " & AdoAux(0).Recordset.RecordCount
                    End Select
    End Select
End Sub

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " numalbar=" & Val(txtAux(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function


Private Sub PosicionarData()
Dim cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    cad = "numalbar = " & AdoAux(0).Recordset!NumAlbar
    ' ***************************************
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    If SituarData(AdoAux(0), cad, Indicador) Then
        lblIndicador.Caption = Indicador
    End If
    ' ***********************************************************************************
End Sub


Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim B As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    If Index <> 3 Then DeseleccionaGrid DataGridAux(Index)
    ' ***************************************************
       
    B = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 0 'albaranes
            txtAux(0).visible = False
            txtAux(1).visible = False
            txtAux(2).visible = False
            txtAux(3).visible = False
            For jj = 4 To 4
                txtAux(jj).visible = B
                txtAux(jj).Top = alto
            Next jj
            
            Text2(2).visible = False
            
            
    End Select
End Sub


Private Sub ModificarLinea()
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Long
Dim cad As String
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'cuentas Bancarias
    End Select
    ' **************************************************************

    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
'??monica
'            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
            If cadWhere <> "" Then BloqueaRegistro "rhisfruta", cadWhere

            If NumTabMto <> 3 Then
                V = AdoAux(NumTabMto).Recordset.Fields(0) 'el 2 es el nº de llinia
                CargaGrid NumTabMto, True
            End If

            ' *** si n'hi han tabs ***
'            SituarTab (NumTabMto)

            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            If NumTabMto <> 3 Then
                DataGridAux(NumTabMto).SetFocus
                AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(0).Name & " =" & V)
            End If
            ' ***********************************************************

            LLamaLineas NumTabMto, 0
        End If
    End If
        
End Sub

Private Sub PonerCampos()
Dim i As Integer
Dim CodPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la capçalera
    
    CargaGrid i, True
    If Not AdoAux(i).Recordset.EOF Then _
        PonerCamposForma2 Me, AdoAux(i), 2, "FrameAux" & i
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
End Sub


Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim B As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte
    
    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False

    B = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not B Then Exit Function
    
    DatosOkLlin = B

EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub PonerModoOpcionesMenu(Modo)
'Actives unes Opcions de Menú i Toolbar según el modo en que estem
Dim B As Boolean, bAux As Boolean
Dim i As Byte
    
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
    B = (Modo = 3 Or Modo = 4 Or Modo = 2)
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = B
        If B Then bAux = (B And Me.AdoAux(i).Recordset.RecordCount > 0)
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i
    ' ****************************************
    
'    ' *** si n'hi han tabs que no tenen grids ***
'    i = 3
'    If AdoAux(i).Recordset.EOF Then
'        ToolAux(i).Buttons(1).Enabled = b
'        ToolAux(i).Buttons(2).Enabled = False
'        ToolAux(i).Buttons(3).Enabled = False
'    Else
'        ToolAux(i).Buttons(1).Enabled = False
'        ToolAux(i).Buttons(2).Enabled = b
'    End If
    ' *******************************************
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index <> 0 And KeyCode <> 38 Then KEYdown KeyCode
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim TipoDto As Byte
Dim ImpDto As String
Dim Unidades As String
Dim cantidad As String

    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtAux(Index), ModoLineas) Then Exit Sub
    
    Select Case Index
        Case 4 ' Importe
            If txtAux(Index).Text <> "" Then PonerFormatoDecimal txtAux(Index), 1
            
    End Select
End Sub

Private Function LimpiarImportes(vWhere As String) As Boolean
On Error GoTo eLimpiarImportes

    LimpiarImportes = False

    'primero limpiamos importes
    conn.Execute "update rhisfruta set impentrada = 0 where " & vWhere

    LimpiarImportes = True
    Exit Function

eLimpiarImportes:
    MuestraError Err.Number, "Limpiar Importes", Err.Description
End Function
                                
Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

