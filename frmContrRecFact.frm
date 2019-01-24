VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmContrRecFact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturación de Entradas"
   ClientHeight    =   11085
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   16170
   Icon            =   "frmContrRecFact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11085
   ScaleWidth      =   16170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   135
      TabIndex        =   101
      Top             =   45
      Width           =   1335
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   102
         Top             =   180
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Pedir Datos"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Generar Facturas"
            EndProperty
         EndProperty
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
      Left            =   13845
      TabIndex        =   18
      Top             =   10395
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
      Left            =   15045
      TabIndex        =   87
      Top             =   10395
      Width           =   1065
   End
   Begin VB.Frame FrameFactura 
      Height          =   5010
      Left            =   11205
      TabIndex        =   28
      Top             =   5190
      Width           =   4895
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
         Index           =   32
         Left            =   3195
         MaxLength       =   15
         TabIndex        =   96
         Tag             =   "Base Imponible 3|N|N|0||scafac|baseimp3|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   630
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
         Index           =   8
         Left            =   3195
         MaxLength       =   15
         TabIndex        =   91
         Tag             =   "Base Imponible 3|N|N|0||scafac|baseimp3|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   1050
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
         Index           =   25
         Left            =   3195
         MaxLength       =   15
         TabIndex        =   57
         Tag             =   "Importe Retencion|N|N|0||scafac|imporete|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   3750
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
         Index           =   24
         Left            =   705
         MaxLength       =   5
         TabIndex        =   56
         Tag             =   "% reten|N|S|0|99.90|scafac|porcereten|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   3750
         Width           =   750
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
         Left            =   1515
         MaxLength       =   15
         TabIndex        =   55
         Tag             =   "Base Imponible 1|N|N|0||scafac|baseimp1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   3735
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
         Left            =   3195
         MaxLength       =   15
         TabIndex        =   51
         Tag             =   "Importe IVA 1|N|N|0||scafac|imporiv1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   1530
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
         Index           =   6
         Left            =   3195
         MaxLength       =   15
         TabIndex        =   49
         Tag             =   "Base Imponible 1|N|N|0||scafac|baseimp1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   225
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
         Index           =   12
         Left            =   120
         MaxLength       =   5
         TabIndex        =   47
         Tag             =   "% IVA 3|N|S|0|99|scafac|porciva3|00|N|"
         Text            =   "Text1 7"
         Top             =   3035
         Width           =   500
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
         Left            =   120
         MaxLength       =   5
         TabIndex        =   46
         Tag             =   "& IVA 2|N|S|0|99|scafac|porciva2|00|N|"
         Text            =   "Text1 7"
         Top             =   2640
         Width           =   500
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
         Left            =   120
         MaxLength       =   5
         TabIndex        =   45
         Tag             =   "% IVA 1|N|S|0|99|scafac|porciva1|00|N|"
         Text            =   "Text1 7"
         Top             =   2250
         Width           =   500
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
         Left            =   1515
         MaxLength       =   15
         TabIndex        =   38
         Tag             =   "Base Imponible 1|N|N|0||scafac|baseimp1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2250
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
         Left            =   705
         MaxLength       =   5
         TabIndex        =   37
         Tag             =   "% IVA 1|N|S|0|99.90|scafac|porciva1|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   2250
         Width           =   750
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
         Left            =   3195
         MaxLength       =   15
         TabIndex        =   36
         Tag             =   "Importe IVA 1|N|N|0||scafac|imporiv1|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2250
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
         Index           =   17
         Left            =   1515
         MaxLength       =   15
         TabIndex        =   35
         Tag             =   "Base Imponible 2 |N|N|0||scafac|baseimp2|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2640
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
         Left            =   705
         MaxLength       =   5
         TabIndex        =   34
         Tag             =   "& IVA 2|N|S|0|99.90|scafac|porciva2|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   2640
         Width           =   750
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
         Left            =   3195
         MaxLength       =   15
         TabIndex        =   33
         Tag             =   "Importe IVA 2|N|N|0||scafac|imporiv2|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   2640
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
         Index           =   18
         Left            =   1515
         MaxLength       =   15
         TabIndex        =   32
         Tag             =   "Base Imponible 3|N|N|0||scafac|baseimp3|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   3035
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
         Left            =   705
         MaxLength       =   5
         TabIndex        =   31
         Tag             =   "% IVA 3|N|S|0|99.90|scafac|porciva3|#0.00|N|"
         Text            =   "Text1 7"
         Top             =   3035
         Width           =   750
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
         Left            =   3195
         MaxLength       =   15
         TabIndex        =   30
         Tag             =   "Importe IVA 3|N|N|0||scafac|imporiv3|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   3035
         Width           =   1485
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
         Left            =   2340
         MaxLength       =   15
         TabIndex        =   29
         Tag             =   "Total Factura|N|N|0||scafac|totalfac|#,###,###,##0.00|N|"
         Text            =   "Text1 7"
         Top             =   4395
         Width           =   2325
      End
      Begin VB.Label Label1 
         Caption         =   "Descuento"
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
         Index           =   25
         Left            =   1515
         TabIndex        =   97
         Top             =   690
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Anticipos"
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
         Left            =   1515
         TabIndex        =   92
         Top             =   1110
         Width           =   1215
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
         Left            =   705
         TabIndex        =   61
         Top             =   3465
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "-"
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
         Index           =   14
         Left            =   3795
         TabIndex        =   60
         Top             =   3060
         Width           =   135
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
         Left            =   3195
         TabIndex        =   59
         Top             =   3450
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
         Left            =   1515
         TabIndex        =   58
         Top             =   3465
         Width           =   1620
      End
      Begin VB.Line Line2 
         X1              =   3150
         X2              =   4635
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   4635
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
         Left            =   1515
         TabIndex        =   52
         Top             =   1530
         Width           =   1665
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
         Left            =   1515
         TabIndex        =   50
         Top             =   285
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Cod."
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
         Left            =   120
         TabIndex        =   48
         Top             =   1980
         Width           =   495
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
         Left            =   1515
         TabIndex        =   44
         Top             =   1980
         Width           =   1530
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
         Left            =   3180
         TabIndex        =   43
         Top             =   1980
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
         Left            =   3600
         TabIndex        =   42
         Top             =   1680
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
         TabIndex        =   41
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "TOTAL FACTURA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   39
         Left            =   2340
         TabIndex        =   40
         Top             =   4125
         Width           =   2340
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
         Left            =   705
         TabIndex        =   39
         Top             =   1980
         Width           =   675
      End
      Begin VB.Line Line3 
         X1              =   135
         X2              =   4635
         Y1              =   3435
         Y2              =   3435
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   7
      Left            =   12360
      MaxLength       =   15
      TabIndex        =   79
      Tag             =   "Base Imponible 2 |N|N|0||scafac|baseimp2|#,###,###,##0.00|N|"
      Text            =   "Text1 7"
      Top             =   5715
      Width           =   1485
   End
   Begin VB.Frame FrameIntro 
      Height          =   4365
      Left            =   120
      TabIndex        =   20
      Top             =   765
      Width           =   15960
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
         Index           =   34
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   3630
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
         Index           =   34
         Left            =   2685
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   98
         Text            =   "Text2"
         Top             =   3630
         Width           =   5400
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
         Index           =   33
         Left            =   9690
         MaxLength       =   15
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   3645
         Width           =   1380
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
         Index           =   31
         Left            =   9690
         MaxLength       =   15
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   3105
         Width           =   1380
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
         Index           =   30
         Left            =   2685
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   93
         Text            =   "Text2"
         Top             =   3090
         Width           =   5400
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
         Index           =   30
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   3090
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
         Index           =   29
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   2550
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
         Index           =   29
         Left            =   2685
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   89
         Text            =   "Text2"
         Top             =   2550
         Width           =   5400
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
         Index           =   28
         Left            =   15015
         MaxLength       =   6
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   2040
         Width           =   570
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
         Left            =   12690
         MaxLength       =   8
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   2040
         Width           =   1170
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
         Index           =   2
         Left            =   13635
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Tag             =   "Tipo IRPF|N|N|0|2|rsocios|tipoirpf||N|"
         Top             =   840
         Width           =   1935
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
         Left            =   13635
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Tag             =   "Tipo IRPF|N|N|0|2|rsocios|tipoirpf||N|"
         Top             =   270
         Width           =   1935
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
         Left            =   9690
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Tag             =   "Tipo IRPF|N|N|0|2|rsocios|tipoirpf||N|"
         Top             =   270
         Width           =   1980
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
         Left            =   2700
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   27
         Text            =   "Text2"
         Top             =   270
         Width           =   5400
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
         Left            =   13650
         TabIndex        =   17
         Top             =   3105
         Width           =   1710
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
         Index           =   4
         Left            =   2685
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   77
         Text            =   "Text2"
         Top             =   1410
         Width           =   5400
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
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1410
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
         Left            =   9690
         MaxLength       =   15
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   2040
         Width           =   1380
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Intracomunitaria"
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
         Left            =   13650
         TabIndex        =   19
         Top             =   2745
         Visible         =   0   'False
         Width           =   1935
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
         Left            =   9690
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Fecha Recepción|F|N|||tcafpc|fecrecep|dd/mm/yyyy|N|"
         Top             =   840
         Width           =   1380
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
         Left            =   2685
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   25
         Text            =   "Text2"
         Top             =   1980
         Width           =   5400
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
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1980
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
         Left            =   1815
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "Cod.Socio|N|N|0|999999|tcafpc|codsocio|000000|S|"
         Text            =   "Text1"
         Top             =   270
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
         Index           =   1
         Left            =   6825
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Factura|F|N|||tcafpc|fecfactu|dd/mm/yyyy|S|"
         Top             =   840
         Width           =   1305
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Nº Factura|T|N|||tcafpc|numfactu||S|"
         Text            =   "Text1 7"
         Top             =   840
         Width           =   1245
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Tesoreria"
         Height          =   255
         Index           =   0
         Left            =   9720
         TabIndex        =   53
         Top             =   810
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Concep.Cargo"
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
         Index           =   27
         Left            =   120
         TabIndex        =   100
         Top             =   3660
         Width           =   1395
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1530
         ToolTipText     =   "Buscar concepto"
         Top             =   3660
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Cargo"
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
         Index           =   26
         Left            =   8250
         TabIndex        =   99
         Top             =   3660
         Width           =   1590
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Dto"
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
         Index           =   24
         Left            =   8250
         TabIndex        =   95
         Top             =   3120
         Width           =   1185
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1530
         ToolTipText     =   "Buscar concepto"
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Concepto Dto"
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
         Index           =   23
         Left            =   120
         TabIndex        =   94
         Top             =   3120
         Width           =   1440
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
         Height          =   315
         Index           =   11
         Left            =   120
         TabIndex        =   90
         Top             =   2580
         Width           =   945
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1530
         ToolTipText     =   "Buscar variedad"
         Top             =   2580
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Documentos del socio"
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
         Index           =   9
         Left            =   12720
         TabIndex        =   88
         Top             =   3705
         Width           =   2250
      End
      Begin VB.Image imgDoc 
         Height          =   465
         Index           =   1
         Left            =   15045
         ToolTipText     =   "Documentos"
         Top             =   3600
         Width           =   510
      End
      Begin VB.Label Label1 
         Caption         =   "%Corredor"
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
         Index           =   22
         Left            =   13890
         TabIndex        =   85
         Top             =   2055
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Precio/Kg"
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
         Index           =   21
         Left            =   11700
         TabIndex        =   84
         Top             =   2055
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Total"
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
         Index           =   20
         Left            =   8250
         TabIndex        =   83
         Top             =   2055
         Width           =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Precio"
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
         Index           =   19
         Left            =   12165
         TabIndex        =   82
         Top             =   870
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Entrada"
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
         Index           =   18
         Left            =   12165
         TabIndex        =   81
         Top             =   285
         Width           =   1440
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
         Height          =   315
         Index           =   17
         Left            =   120
         TabIndex        =   76
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1530
         ToolTipText     =   "Buscar forma pago"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Factura"
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
         Index           =   16
         Left            =   8250
         TabIndex        =   62
         Top             =   285
         Width           =   1305
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   9420
         Picture         =   "frmContrRecFact.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   870
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   6540
         Picture         =   "frmContrRecFact.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1530
         ToolTipText     =   "Buscar banco propio"
         Top             =   2010
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1530
         ToolTipText     =   "Buscar socio"
         Top             =   285
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fec.Recep."
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
         Left            =   8250
         TabIndex        =   26
         Top             =   900
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Cta.Pr.Pago"
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
         Left            =   120
         TabIndex        =   24
         Top             =   2010
         Width           =   1215
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
         Left            =   120
         TabIndex        =   23
         Top             =   300
         Width           =   660
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   29
         Left            =   4980
         TabIndex        =   22
         Top             =   870
         Width           =   1545
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
         Height          =   315
         Index           =   28
         Left            =   120
         TabIndex        =   21
         Top             =   870
         Width           =   1095
      End
   End
   Begin VB.Frame FrameAux0 
      Height          =   5010
      Left            =   120
      TabIndex        =   63
      Top             =   5190
      Width           =   11020
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
         Index           =   2
         Left            =   10035
         MaxLength       =   6
         TabIndex        =   66
         Tag             =   "Precio Estimado|N|S|||rhisfruta|prestimado|###,##0.0000|N|"
         Text            =   "prec"
         Top             =   4320
         Visible         =   0   'False
         Width           =   675
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4185
         Left            =   90
         TabIndex        =   86
         Top             =   540
         Width           =   10795
         _ExtentX        =   19050
         _ExtentY        =   7382
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
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
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   6
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   78
         Tag             =   "Variedad|N|N|||rhisfruta|codvarie|000000|N|"
         Text            =   "var"
         Top             =   3600
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   5
         Left            =   6420
         MaxLength       =   35
         TabIndex        =   75
         Tag             =   "Socio|N|N|||rhisfruta|codsocio|000000||"
         Text            =   "Socio"
         Top             =   3600
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   3
         Left            =   4350
         MaxLength       =   35
         TabIndex        =   71
         Tag             =   "Kilos Netos|N|N|||rhisfruta|kilosnet|###,##0|N|"
         Text            =   "Kilosnet"
         Top             =   3600
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   1770
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   70
         Text            =   "Text2"
         Top             =   3600
         Width           =   2520
      End
      Begin VB.Frame Frame1 
         Height          =   555
         Index           =   0
         Left            =   180
         TabIndex        =   68
         Top             =   4020
         Width           =   2865
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
            TabIndex        =   69
            Top             =   180
            Width           =   2655
         End
      End
      Begin VB.TextBox txtaux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   4
         Left            =   5790
         MaxLength       =   35
         TabIndex        =   67
         Tag             =   "Importe|N|N|||rhisfruta|impentrada|###,###0.00||"
         Text            =   "Importe"
         Top             =   3600
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtaux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   0
         Left            =   150
         MaxLength       =   7
         TabIndex        =   65
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
         Height          =   290
         Index           =   1
         Left            =   600
         MaxLength       =   10
         TabIndex        =   64
         Tag             =   "Fec.Albara|F|N|||rhisfruta|fecalbar|dd/mm/yyyy|N|"
         Text            =   "Fec.Alb"
         Top             =   3600
         Visible         =   0   'False
         Width           =   390
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   0
         Left            =   135
         TabIndex        =   72
         Top             =   675
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
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
         Bindings        =   "frmContrRecFact.frx":0122
         Height          =   2910
         Index           =   0
         Left            =   135
         TabIndex        =   73
         Top             =   1125
         Width           =   7110
         _ExtentX        =   12541
         _ExtentY        =   5133
         _Version        =   393216
         AllowUpdate     =   0   'False
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
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
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   10230
         Picture         =   "frmContrRecFact.frx":013A
         Top             =   210
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   10590
         Picture         =   "frmContrRecFact.frx":0284
         Top             =   210
         Width           =   240
      End
      Begin VB.Image imgDoc 
         Height          =   315
         Index           =   0
         Left            =   2640
         ToolTipText     =   "Asignacion Precios"
         Top             =   150
         Width           =   330
      End
      Begin VB.Label Label2 
         Caption         =   "Albaranes del Socio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   74
         Top             =   240
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   6960
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
   Begin VB.Label Label1 
      Caption         =   "Imp. dto. ppago"
      Height          =   255
      Index           =   8
      Left            =   11250
      TabIndex        =   80
      Top             =   5715
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Operador"
      Height          =   255
      Index           =   1
      Left            =   1845
      TabIndex        =   54
      Top             =   900
      Width           =   735
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1845
      Picture         =   "frmContrRecFact.frx":03CE
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
Attribute VB_Name = "frmContrRecFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
Private WithEvents frmMens As frmMensajes 'Mensajes
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmPrecios As frmContrRecFactPre ' precios por calidades
Attribute frmPrecios.VB_VarHelpID = -1
Private WithEvents frmGas As frmManConcepGasto 'Form Mto de conceptos de gastos
Attribute frmGas.VB_VarHelpID = -1
Private WithEvents frmCargo As frmFVARConceptos 'Form Mto de conceptos de cargos
Attribute frmCargo.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe
Private ConSubInforme As Boolean


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
Dim cadWHERE As String

'Cuando vuelve del formulario de ver los albaranes seleccionados hay que volver
'a cargar los datos de los albaranes
'Dim VerAlbaranes As Boolean

Dim PrimeraVez As Boolean
Dim Bloquear As Boolean
Dim indice As Integer

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
Dim TotalFactAnticipo As Currency

Dim Anticipos As String


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

Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbLightBlue
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Combo1(Index).BackColor = vbLightBlue Then Combo1(Index).BackColor = vbWhite
    '[Monica]24/01/2019: dejo todo abierto
    If vParamAplic.Cooperativa = 18 Then
        Text1(26).Enabled = True
        Text1(27).Enabled = True
    Else
        Text1(26).Enabled = (Combo1(1).ListIndex = 1) Or (Combo1(0).ListIndex = 1)
        Text1(27).Enabled = Not ((Combo1(1).ListIndex = 1) Or (Combo1(0).ListIndex = 1))
    End If
    
    If Text1(26).Enabled Then Text1(27).Text = ""
    If Text1(27).Enabled Then Text1(26).Text = ""
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
    If Not Adoaux(0).Recordset.EOF Then _
        Me.lblIndicador.Caption = Adoaux(0).Recordset.AbsolutePosition & " de " & Adoaux(0).Recordset.RecordCount
End Sub

Private Sub Form_Activate()
'    Screen.MousePointer = vbDefault
'    If VerAlbaranes Then RefrescarAlbaranes
'    VerAlbaranes = False
    If PrimeraVez Then
        mnPedirDatos_Click
    End If
    PrimeraVez = False
    

End Sub

Private Sub Form_Load()
Dim i As Integer
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Pedir Datos
        .Buttons(2).Image = 15   'Generar FActura
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
    For i = 2 To 6
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 1 To 1
        Me.imgDoc(i).Picture = frmPpal.imgListPpal.ListImages(24).Picture
    Next i
    For i = 0 To 0
        Me.imgDoc(i).Picture = frmPpal.imgListPpal.ListImages(9).Picture
    Next i
    
    Me.FrameFactura.Enabled = False
    
    LimpiarCampos   'Limpia los campos TextBox
'    InicializarListView
   
    '## A mano
    NombreTabla = "rhisfruta" ' albaranes de venta
    
    CargaCombo
    
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
'    CargaGrid 0, False
    InicializarListView
End Sub


Private Sub LimpiarCampos()
Dim i As Integer

On Error Resume Next
    limpiar Me
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Check1(2).Value = 0
    For i = 0 To Combo1.Count - 1
        Combo1(i).ListIndex = -1
    Next i
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
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
Dim indice As Byte
    indice = CByte(Me.imgFecha(0).Tag)
    Text1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmFPa_DatoSeleccionado(CadenaSeleccion As String)
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod forpa
    FormateaCampo Text1(4)
    Text2(4).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom forpa
End Sub

Private Sub frmGas_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Conceptos de gastos
    Text1(30).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00") 'Cod concepto de gasto
    Text2(30).Text = RecuperaValor(CadenaSeleccion, 2) 'Descripcion
End Sub


Private Sub frmCargo_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Conceptos de cargo fvarconcep
    Text1(34).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00") 'Cod concepto de cargo
    Text2(34).Text = RecuperaValor(CadenaSeleccion, 2) 'Descripcion
End Sub



Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)

    Anticipos = ""
    If CadenaSeleccion <> "" Then
        Anticipos = "(" & CadenaSeleccion & ")"
    End If

End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
Dim indice As Byte
    indice = 3
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Socios
    FormateaCampo Text1(indice)
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom socio
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    Text1(29).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod variedad
    FormateaCampo Text1(29)
    Text2(29).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre variedad
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
'    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Socio
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            indice = 3
       
       Case 2 'Bancos Propios
            indice = 5
            Set frmBanPr = New frmComBanco
            frmBanPr.DatosADevolverBusqueda = "0|1|"
            frmBanPr.Show vbModal
            Set frmBanPr = Nothing
            
       
       Case 3 'Forma de pago
            Set frmFPa = New frmComFpa
            frmFPa.DatosADevolverBusqueda = "0|1|"
            '[Monica]04/01/2018: para crear la forma de pago en contabilidad
            frmFPa.Seccion = CStr(vParamAplic.Seccionhorto)
            frmFPa.Show vbModal
            Set frmFPa = Nothing
            indice = 4
       
       Case 4 'variedades
            Set frmVar = New frmComVar
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.Show vbModal
            Set frmVar = Nothing
            indice = 4
    
       Case 5 ' concepto de descuento
            Set frmGas = New frmManConcepGasto
            frmGas.DatosADevolverBusqueda = "0|1|"
            frmGas.CodigoActual = Text1(30).Text
            frmGas.Show vbModal
            Set frmGas = Nothing
            PonerFoco Text1(30)
    
       Case 6 ' concepto de cargo
            Set frmCargo = New frmFVARConceptos
            frmCargo.DatosADevolverBusqueda = "0|1|"
            frmCargo.CodigoActual = Text1(34).Text
            frmCargo.Show vbModal
            Set frmCargo = Nothing
            PonerFoco Text1(34)
    
    
    End Select
    
    PonerFoco Text1(indice)
'    Screen.MousePointer = vbDefault
End Sub

Private Sub imgDoc_Click(Index As Integer)
Dim Albaranes As String
Dim i As Integer

    If Text1(3).Text = "" Then Exit Sub
       
    Select Case Index
       Case 0 ' asignacion de precios
            If Text1(27).Text = "" Then Exit Sub
            
            Albaranes = ""
            
            For i = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(i).Checked Then
                    Albaranes = Albaranes & ListView1.ListItems(i).Text & ","
                End If
            Next i
            ' lista de albaranes
            If Albaranes <> "" Then
                Albaranes = Mid(Albaranes, 1, Len(Albaranes) - 1)
            
                Set frmPrecios = New frmContrRecFactPre
                frmPrecios.Precio = Text1(27).Text
                frmPrecios.Albaranes = Albaranes
                frmPrecios.Show vbModal
                
                Set frmPrecios = Nothing
                
                
                CalcularDatosFacturaNew
            End If
            
    
       Case 1 ' imagenes del socio
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
            frmSoc.Socio = Text1(3)
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            indice = 3
    End Select

End Sub

Private Sub imgFecha_Click(Index As Integer)
Dim indice As Byte
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
   indice = Index + 1
   Me.imgFecha(0).Tag = indice
   
   PonerFormatoFecha Text1(indice)
   If Text1(indice).Text <> "" Then frmF.NovaData = CDate(Text1(indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(indice)
End Sub


Private Sub ListView1_DblClick()

'    txtAux(2).visible = True
'    txtAux(2).Enabled = True
'    txtAux(2).Locked = False
'
'
'    txtAux(2).Top = ListView1.SelectedItem.Top + 600
'    txtAux(2).Width = ListView1.ColumnHeaders(7).Width
'    txtAux(2).Left = ListView1.ColumnHeaders(7).Left + 100
'    txtAux(2).ToolTipText = ListView1.SelectedItem.Text
'
'    txtAux(2).Text = Text1(27).Text
'
'    PonerFoco txtAux(2)
    
End Sub

Private Sub ListView1_ItemCheck(ByVal item As MSComctlLib.ListItem)
Dim cantidad As String
Dim Palets As String
Dim Sql As String
Dim Valor As Currency
Dim Valor2 As Currency
Dim i As Long
Dim b As Boolean

    If Modo <> 5 Then Exit Sub

    If Bloquear = True Then
        ListView1.SetFocus
        item.EnsureVisible
        Exit Sub
    End If
        
    i = ListView1.SelectedItem.Index
    
    If i = 1 Then
    
    End If
    
    
'    CargarAlbaranes vWhere

    TerminaBloquear

    Sql = "update rhisfruta_clasif set precio = " & DBSet(Text1(27).Text, "N")
    '[Monica]09/01/2019:
    If vParamAplic.Cooperativa = 18 Then
        '[Monica]19/09/2013: hemos añadido el importe para que puedan modificarlo
        Sql = Sql & " , importe = round(kilostra * " & DBSet(Text1(27).Text, "N") & ", 2)"
    Else
        '[Monica]19/09/2013: hemos añadido el importe para que puedan modificarlo
        Sql = Sql & " , importe = round(kilosnet * " & DBSet(Text1(27).Text, "N") & ", 2)"
    End If
    Sql = Sql & " where numalbar = " & DBSet(item.Text, "N")
    conn.Execute Sql

    CalcularDatosFacturaNew
    
'    ' Crea una variable ListItem.
'    ' Establece la variable al elemento encontrado.
'    If I < ListView1.ListItems.Count Then
'        ListView1.SelectedItem = ListView1.ListItems.item(I + 1)
'    Else
'        ListView1.SelectedItem = ListView1.ListItems.item(I)
'    End If
'    ListView1.SetFocus
'    item.EnsureVisible
        

End Sub

Private Sub ListView1_ItemClick(ByVal item As MSComctlLib.ListItem)
'Dim cantidad As String
'Dim Palets As String
'Dim Sql As String
'Dim Valor As Currency
'Dim Valor2 As Currency
'Dim I As Long
'Dim b As Boolean
'
'    If Modo <> 5 Then Exit Sub
'
'    If Bloquear = True Then
'        ListView1.SetFocus
'        item.EnsureVisible
'        Exit Sub
'    End If
'
'    I = ListView1.SelectedItem.Index
'
''    CargarAlbaranes vWhere
'    TerminaBloquear
'
'    CalcularDatosFactura
'
'    ' Crea una variable ListItem.
'    ' Establece la variable al elemento encontrado.
'    If I < ListView1.ListItems.Count Then
'        ListView1.SelectedItem = ListView1.ListItems.item(I + 1)
'    Else
'        ListView1.SelectedItem = ListView1.ListItems.item(I)
'    End If
'    ListView1.SetFocus
'    item.EnsureVisible
'
'
End Sub


Private Sub mnModificarDto_Click()
Dim i As Integer


    If Text1(6).Text = "" Then Exit Sub

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
    BotonPedirDatos True
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
        Case 0 ' numero de factura, si el socio es tercero es requerida
            If Text1(3).Text = "" Then Exit Sub
            Set vSocio = New cSocio
            If vSocio.LeerDatos(Text1(3).Text) Then
                If vSocio.Estercero(Text1(3).Text, True) Then
                    If Text1(0).Text = "" Then
                        MsgBox "Debe de introducir el nro de factura. Reintroduzca", vbExclamation
                        PonerFoco Text1(0)
                    End If
                End If
            End If
            Set vSocio = Nothing
    
        Case 1, 2 'Fecha factura, fecha recepcion
            If PonerFormatoFecha(Text1(Index)) Then
                If Text1(Index) <> "" Then
                    ' No debe existir el número de factura para el proveedor en hco
                    If ExisteFacturaEnHco Then
    '                    InicializarListView
                    End If
                    If Index = 1 Then Text1(2).Text = Text1(1).Text
                End If
            End If
            
        Case 3 'Cod Socios
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rsocios", "nomsocio", "codsocio")
                Text1(0).Enabled = True
                '[Monica]23/09/2013: solo si es tercero permito cargo adicional
                Text1(34).Enabled = True
                Text1(33).Enabled = True
                imgBuscar(6).Enabled = True
                
                If Text2(3).Text = "" Then
                    MsgBox "Socio no existe. Reintroduzca", vbExclamation
                    PonerFoco Text1(Index)
                Else
                    Set vSocio = New cSocio
                    Text1(0).Enabled = vSocio.Estercero(Text1(3).Text, True)
                    
                    '[Monica]23/09/2013: solo si es tercero permito cargo adicional
                    Text1(34).Enabled = vSocio.Estercero(Text1(3).Text, True)
                    Text1(33).Enabled = vSocio.Estercero(Text1(3).Text, True)
                    imgBuscar(6).Enabled = vSocio.Estercero(Text1(3).Text, True)
                    If Not vSocio.Estercero(Text1(3).Text, True) Then
                        Text1(34).Text = ""
                        Text2(34).Text = ""
                        Text1(33).Text = ""
                    End If
                    Set vSocio = Nothing
                    
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 5 'Cta Prevista de PAgo
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "banpropi", "nombanpr", "codbanpr")
                Text1(Index).Text = Format(Text1(Index).Text, "0000")
            Else
                Text2(Index).Text = ""
            End If
            
            
        Case 26 'importe de la factura
            PonerFormatoDecimal Text1(Index), 3
        
        Case 27
            PonerFormatoDecimal Text1(Index), 8
        
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
                CalcularDatosFacturaNew
                lblIndicador.Caption = ""
                BloquearTxt Text1(8), True
                Me.FrameFactura.Enabled = False
                PonerModo 5
            End If
            
        Case 29 ' Variedades
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "variedades", "nomvarie", "codvarie")
                Text1(Index).Text = Format(Text1(Index).Text, "000000")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 30 ' concepto de gasto
            If Text1(Index) <> "" Then
                Text2(Index) = DevuelveDesdeBDNew(cAgro, "rconcepgasto", "nomgasto", "codgasto", Text1(Index), "N")
                If Text2(Index).Text = "" Then
                    MsgBox "Código no existe. Revise.", vbExclamation
                    PonerFoco Text1(Index)
                Else
                    If EsGastodeFactura(Text1(Index).Text) = True Then
                        MsgBox "Este concepto de gasto es de factura. Reintroduzca.", vbExclamation
                        PonerFoco Text1(Index)
                    End If
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        Case 31 ' importe de dto
            PonerFormatoDecimal Text1(Index), 1
        
         
        '[Monica]23/092013: insertamos un concepto de cargo que aumenta la factura (puede q con otro tipo de iva)
        Case 34 ' concepto de cargo
            If Text1(Index) <> "" Then
                Text2(Index) = DevuelveDesdeBDNew(cAgro, "fvarconce", "nomconce", "codconce", Text1(Index), "N")
                If Text2(Index).Text = "" Then
                    MsgBox "Código no existe. Revise.", vbExclamation
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        Case 33 ' importe de dto
            PonerFormatoDecimal Text1(Index), 1
       
    End Select
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte, NumReg As Byte
Dim b As Boolean
On Error GoTo EPonerModo

    Modo = Kmodo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
        
    CmdAceptar.visible = (ModoLineas = 2)
    CmdAceptar.Enabled = (ModoLineas = 2)
    cmdCancelar.visible = (ModoLineas = 2)
    cmdCancelar.Enabled = (ModoLineas = 2)
    
'    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
'    'Si estamos en Insertar además limpia los campos Text1
'    'si estamos en modificar bloquea las compos que son clave primaria
'    BloquearText1 Me, Modo
    
'    BloquearCombo Me, Modo

    'b = (Modo = 3 Or Modo = 4 Or Modo = 1 Or Modo = 5)
    b = (Modo = 3 Or Modo = 4 Or Modo = 1) '06/09/2005, lleve el modo 5 per a que no es puga modificar la capçalera mentre treballe en les llínies
    
    For i = 0 To Combo1.Count - 1
        Combo1(i).Enabled = b
        If b Then
            Combo1(i).BackColor = vbWhite
        Else
            Combo1(i).BackColor = &H80000018 'Amarillo Claro
        End If
    Next i

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
    b = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = b
    Next i
                    
'    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    
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
Dim Cad As String
Dim i As Byte
Dim vSeccion As CSeccion

    On Error GoTo EDatosOK
    DatosOk = False
    
    ' deben de introducirse todos los datos del frame
    For i = 1 To 5
        If Text1(i).Text = "" Then
            If Text1(i).Tag <> "" Then
                Set vtag = New CTag
                If vtag.Cargar(Text1(i)) Then
                    Cad = vtag.Nombre
                Else
                    Cad = "Campo"
                End If
                Set vtag = Nothing
            Else
                Cad = "Campo"
                If i = 5 Then Cad = "Cta. Prev. Pago"
                If i = 4 Then Cad = "Forma de Pago"
            End If
            MsgBox Cad & " no puede estar vacio. Reintroduzca", vbExclamation
            PonerFoco Text1(i)
            Exit Function
        End If
    Next i
        
    'comprobar que la fecha de la factura sea anterior a la fecha de recepcion
    If Not EsFechaIgualPosterior(Text1(1).Text, Text1(2).Text, True, "La fecha de recepción debe ser igual o posterior a la fecha de la factura.") Then
        PonerFoco Text1(1)
        Exit Function
    End If
    
    'Comprobar que la fecha de RECEPCION esta dentro de los ejercicios contables
    Set vSeccion = New CSeccion
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If vSeccion.AbrirConta Then
            '[Monica]20/06/2017: control de fechas que antes no estaba
            ResultadoFechaContaOK = EsFechaOKConta(CDate(Text1(2).Text))
            If ResultadoFechaContaOK > 0 Then
                If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
                vSeccion.CerrarConta
                Set vSeccion = Nothing
                PonerFoco Text1(2)
                Exit Function
            End If
        End If
    End If
    vSeccion.CerrarConta
    Set vSeccion = Nothing

    
    
    If Combo1(1).ListIndex = 1 Or Combo1(0).ListIndex = 1 Then
        If Text1(29).Text = "" Then
            MsgBox "Si va a liquidar sin entradas o a anticipar, ha de introducir la variedad. Revise.", vbExclamation
            PonerFoco Text1(29)
            Exit Function
        Else
            If Text2(29).Text = "" Then
                MsgBox "La variedad introducida ha de existir. Revise.", vbExclamation
                PonerFoco Text1(29)
                Exit Function
            End If
        End If
    End If


    If Combo1(1).ListIndex = 1 Or Combo1(0).ListIndex = 1 Then
        ' venta campo o liq sin entradas
        If ComprobarCero(Text1(26).Text) = 0 Then
            MsgBox "Si va a facturar entradas de venta campo o liq sin entradas, ha de introducir el Importe Total. Revise.", vbExclamation
            PonerFoco Text1(26)
            Exit Function
        End If
    End If
    If Not (Combo1(1).ListIndex = 1 Or Combo1(0).ListIndex = 1) Then
        If ComprobarCero(Text1(27).Text) = 0 Then
            MsgBox "Si va a facturar entradas normales, ha de introducir el precio/kilo. Revise.", vbExclamation
            PonerFoco Text1(27)
            Exit Function
        End If
    End If
    
    Set vSocio = New cSocio
    If vSocio.LeerDatos(Text1(3).Text) Then
        If vSocio.Estercero(Text1(3).Text, True) Then
            If Text1(0).Text = "" Then
                MsgBox "El socio es un tercero. Debe de introducir el Nro de Factura. Reintroduzca.", vbExclamation
                PonerFoco Text1(0)
            Else
                ' No debe existir el número de factura para el socio en hco
                If ExisteFacturaEnHco Then
                    PonerFoco Text1(0)
                    Exit Function
                End If
            End If
        End If
    End If
    
    DatosOk = True
    Exit Function
EDatosOK:
    DatosOk = False
    MuestraError Err.Number, "Comprobar datos correctos", Err.Description
End Function



Private Function DatosOkFact() As Boolean
'Comprobar que los datos del frame de introduccion son correctos antes de cargar datos
Dim vtag As CTag
Dim Cad As String
Dim i As Byte
Dim vSeccion As CSeccion


    On Error GoTo EDatosOK
    DatosOkFact = False
    
'[Monica]20/06/2017: he quitado esto pq lo tenemos que ver una vez abierta la conta de hortofruticola
'    If vParamAplic.NumeroConta <> 0 Then
'        I = EsFechaOKConta(CDate(Text1(2).Text))
'        If I > 0 Then
'            'If i = 1 Then
'                MsgBox "Fecha fuera ejercicios contables", vbExclamation
'                Exit Function
'           ' Else
'           '     cad = "La fecha es superior al ejercico contable siguiente. ¿Desea continuar?"
'           '     If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
'           ' End If
'        End If
'    End If
    
    
    'Comprobar que la fecha de RECEPCION esta dentro de los ejercicios contables
    Set vSeccion = New CSeccion
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If vSeccion.AbrirConta Then
            '[Monica]20/06/2017: control de fechas que antes no estaba
            If Not vSocio.Estercero(Text1(3).Text, True) Then
                ResultadoFechaContaOK = EsFechaOKConta(CDate(Text1(1).Text))
            Else
                ResultadoFechaContaOK = EsFechaOKConta(CDate(Text1(2).Text))
            End If
            If ResultadoFechaContaOK > 0 Then
                If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
                vSeccion.CerrarConta
                Set vSeccion = Nothing
                Exit Function
            End If
        End If
    End If
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
    
    
'--monica:03/12/2008
    'comprobar que se han seleccionado lineas para facturar
    If cadWHERE = "" And (Combo1(0).ListIndex = 0 And Combo1(1).ListIndex = 0) Then
        MsgBox "Debe seleccionar albaranes para facturar.", vbExclamation
        Exit Function
    End If
    
'++monica:03/12/2008
    'comprobamos que hay lineas para facturar: o albaranes o portes de vuelta
    If cadWHERE = "" And (Combo1(0).ListIndex = 0 And Combo1(1).ListIndex = 0) Then
        If Adoaux(0).Recordset.RecordCount = 0 Then
            MsgBox "No hay albaranes para incluir en la factura. Revise.", vbExclamation
            Exit Function
        End If
    End If

    ' No debe existir el número de factura para el socio tercero en hco
    If ExisteFacturaEnHco Then Exit Function


    'Si la forpa es TRANSFERENCIA entonces compruebo la si tiene cta bancaria
'    cad = "select distinct (codforpa) from scaalp "
'    cad = cad & " WHERE " & Replace(cadWHERE, "slialp", "scaalp")
    Set miRsAux = New ADODB.Recordset
'    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    cad = miRsAux.Fields(0)
'    miRsAux.Close



    'Ahora buscamos el tipforpa del codforpa
    Cad = "Select tipoforp from forpago where codforpa=" & DBSet(Text1(4).Text, "N")
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    If miRsAux.EOF Then
        MsgBox "Error en el TIPO de forma de pago", vbExclamation
    Else
        i = 1
        Cad = miRsAux.Fields(0)
        If Val(Cad) = vbFPTransferencia Then
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
            Cad = "Cuenta bancaria incorrecta. Forma de pago: transferencia.    ¿Continuar?"
            If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then i = 0
        End If
    End If

    'Si i=0 es que o esta mal la forpa o no quiere seguir pq no tiene cuenta bancaria
    If i > 0 Then DatosOkFact = True
    Exit Function
    DatosOkFact = True
    Exit Function
EDatosOK:
    DatosOkFact = False
    MuestraError Err.Number, "Comprobar datos correctos", Err.Description
End Function


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Pedir datos
             mnPedirDatos_Click
            
'        Case 2
'             mnModificarDto_Click
            
        Case 2 'Generar Factura
            mnGenerarFac_Click

'        Case 6    'Salir
'            mnSalir_Click
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

 
Private Sub BotonPedirDatos(Preguntar As Boolean)
Dim Nombre As String
Dim i As Integer

    TerminaBloquear

    'Vaciamos todos los Text
    If Text1(3).Text <> "" And Preguntar Then
        If MsgBox("¿ Desea limpiar datos ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            LimpiarCampos
            'fecha recepcion
            Text1(2).Text = Format(Now, "dd/mm/yyyy")
            'fecha de factura
            Text1(1).Text = Format(Now, "dd/mm/yyyy")
        
            Combo1(0).ListIndex = 0
            Combo1(1).ListIndex = 0
            Combo1(2).ListIndex = 0
        
        End If
    Else
        LimpiarCampos
        'fecha recepcion
        Text1(2).Text = Format(Now, "dd/mm/yyyy")
        'fecha de factura
        Text1(1).Text = Format(Now, "dd/mm/yyyy")
        
        Combo1(0).ListIndex = 0
        Combo1(1).ListIndex = 0
        Combo1(2).ListIndex = 0
    End If
    
    Check1(1).Value = 0
    'Vaciamos el ListView
    InicializarListView
    LimpiarImportes "(1=1)"
    For i = 6 To 25
         Text1(i).Text = ""
    Next i

    'Como no habrá albaranes seleccionados vaciamos la cadwhere
    cadWHERE = ""
    
    PonerModo 3
    
    'desbloquear los registros de la rhisfruta
    TerminaBloquear
    
    'si vamos
    'desBloqueo Manual de las tablas
'    DesBloqueoManual ("scaalp")
    
    Me.CmdAceptar.visible = True
    Me.CmdAceptar.Enabled = True
    Me.cmdCancelar.visible = True
    Me.cmdCancelar.Enabled = True
    
    
    PonerFoco Text1(3)
End Sub


Private Sub CargarAlbaranes(cadWHERE As String)
'Recupera de la BD y muestra en el Listview todos los albaranes de compra
'que tiene el proveedor introducido.
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Tabla As String
Dim ItmX As ListItem
Dim i As Integer

On Error GoTo ECargar
    
    ListView1.ListItems.Clear

    Tabla = "rhisfruta"
    
    '[Monica]09/01/2019: para el caso de frutas inma, los kilos son los de transporte
    If vParamAplic.Cooperativa = 18 Then
        Sql = "SELECT rhisfruta.numalbar,rhisfruta.fecalbar, variedades.nomvarie, rhisfruta.codsocio, rsocios.nomsocio, rhisfruta.kilostra kilosnet " ', " & DBSet(Text1(27).Text, "N") & " prestimado "
        Sql = Sql & " FROM (" & Tabla & " inner join variedades on rhisfruta.codvarie = variedades.codvarie) "
        Sql = Sql & " inner join rsocios on rhisfruta.codsocio = rsocios.codsocio "
        Sql = Sql & " where " & cadWHERE
    
    Else
        Sql = "SELECT rhisfruta.numalbar,rhisfruta.fecalbar, variedades.nomvarie, rhisfruta.codsocio, rsocios.nomsocio, rhisfruta.kilosnet " ', " & DBSet(Text1(27).Text, "N") & " prestimado "
        Sql = Sql & " FROM (" & Tabla & " inner join variedades on rhisfruta.codvarie = variedades.codvarie) "
        Sql = Sql & " inner join rsocios on rhisfruta.codsocio = rsocios.codsocio "
        Sql = Sql & " where " & cadWHERE
    End If
        
    ' quitamos los albaranes q hayan sido cobrados
    If Check1(2).Value = 0 Then
        Sql = Sql & " and not rhisfruta.numalbar in (select numalbar from rfactsoc_albaran union select numalbar from rlifter)"
    End If
    
    Sql = Sql & " ORDER BY " & Tabla & ".numalbar,  " & Tabla & ".fecalbar "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    InicializarListView
    
    If Rs.EOF Then
        MsgBox "No existen albaranes pendientes de facturar para este socio.", vbExclamation
        PonerFoco Text1(3)
        Exit Sub
        'BotonPedirDatos
    End If
    
    While Not Rs.EOF
        Set ItmX = ListView1.ListItems.Add()
        ItmX.Text = Rs!NumAlbar
        ItmX.SubItems(1) = Format(DBLet(Rs!Fecalbar, "F"), "dd/mm/yyyy")
        ItmX.SubItems(2) = DBLet(Rs!nomvarie, "T")
        ItmX.SubItems(3) = Format(DBLet(Rs!Codsocio, "N"), "000000")
        ItmX.SubItems(4) = DBLet(Rs!nomsocio, "T")
        ItmX.SubItems(5) = Format(DBLet(Rs!KilosNet, "N"), "###,##0")
'        ItmX.SubItems(6) = Format(DBLet(Rs!PrEstimado, "N"), "##0.0000")
        
        If EstaFacturado(Rs!NumAlbar) Then
            ItmX.ForeColor = vbRed
            For i = 1 To 5 '6
                ItmX.ListSubItems(i).ForeColor = vbRed
            Next i
        End If
        
        'Sig
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing

    PonerModo 5
'    CargaGrid 0, True
'
'    If AdoAux(0).Recordset.RecordCount = 0 Then
'        MsgBox "No existen albaranes pendientes de facturar para este socio.", vbExclamation
'        BotonPedirDatos
'    End If


ECargar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando Albaranes", Err.Description
End Sub


Private Sub InicializarListView()
'Inicializa las columnas del List view

    ListView1.ListItems.Clear
    
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Albarán", 1200
    ListView1.ColumnHeaders.Add , , "Fecha", 1400, 2
    ListView1.ColumnHeaders.Add , , "Variedad", 1700
    ListView1.ColumnHeaders.Add , , "Socio", 1150
    ListView1.ColumnHeaders.Add , , "Nombre", 3850
    ListView1.ColumnHeaders.Add , , "Kilos", 1150, 1
'    ListView1.ColumnHeaders.Add , , "Precio", 800, 1
    
End Sub


Private Sub CalcularDatosFactura()
Dim i As Integer
Dim Sql As String
Dim cadAux As String
Dim ImpBruto As Currency
Dim impiva As Currency
Dim vFactu As CFacturaTer
Dim Rs As ADODB.Recordset
Dim Dto As Currency
Dim TotalKilos As Currency
Dim vPrecio As Currency
Dim PorcIva As Currency
Dim vImporte As Currency
Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim Variedad As String
Dim TipoIVA As Integer

Dim BrutoFact As Currency
Dim BaseImp As Currency
Dim BaseIva As Currency
Dim BaseReten As Currency
Dim PorcReten As Currency
Dim ImpReten As Currency
Dim TotalFac As Currency
Dim Diferencia As Currency

Dim vbase As Currency
Dim Ultimo As Long
Dim vtotfac As Currency
Dim vtotcal As Currency
Dim Albaranes As String

    On Error GoTo eCalcularDatosFactura


    Set vSocio = New cSocio
    If vSocio.LeerDatos(Text1(3).Text) Then
        If vSocio.LeerDatosSeccion(Text1(3).Text, vParamAplic.Seccionhorto) Then
        
        
            Dto = 0
'            If Text1(8).Text <> "" Then
'                Dto = CCur(TransformaPuntosComas(ImporteSinFormato(Text1(8).Text)))
'            End If
            'Limpiar en el form los datos calculados de la factura
            'y volvemos a recalcular
            For i = 6 To 25
                 Text1(i).Text = ""
            Next i
        
            cadAux = ""
            cadWHERE = ""
            ImpBruto = 0
            
            vPrecio = 0
            vImporte = 0
            
            'calculo el total de kilos
            TotalKilos = 0
            For i = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(i).Checked Then
                    TotalKilos = TotalKilos + DBSet(ListView1.ListItems(i).SubItems(5), "N")
                End If
            Next i
            
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
                If vSeccion.AbrirConta Then
                    PorcIva = 0
                    PorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(vSocio.CodIva), "N")
                End If
            End If
            Set vSeccion = Nothing
                
            If Combo1(1).ListIndex = 1 Then ' entradas de venta campo
                Select Case Combo1(2).ListIndex
                    Case 0 'precio normal
                        vImporte = CCur(ComprobarCero(Text1(26).Text))
                        vbase = vImporte
                    Case 1 'precio iva incluido con retencion
                        vImporte = CCur(ComprobarCero(Text1(26).Text))
                        vtotfac = vImporte
                        Select Case vSocio.TipoIRPF
                            Case 0 'retencion sobre base + iva
                                ' le quito la retencion
                                vbase = Round2(vImporte / ((1 + (PorcIva / 100)) * (1 - (vParamAplic.PorcreteFacSoc / 100))), 2)
                                
                            Case 1 'retencion sobre base
                                ' le quito la retencion
                                vbase = Round2(vImporte / (1 + (PorcIva / 100) - (vParamAplic.PorcreteFacSoc / 100)), 2)
                            
                            Case 2 ' sin retencion
                                ' le quito el iva
                                vbase = Round2(vImporte / (1 + (PorcIva / 100)), 2)
                        End Select
                        
                    Case 2 'precio iva incluido sin retencion
                        vImporte = CCur(ComprobarCero(Text1(26).Text))
                        'le añado la retencion
                        vImporte = Round2(vImporte * (1 - (vParamAplic.PorcreteFacSoc / 100)), 2)
                        vtotfac = vImporte
                        Select Case vSocio.TipoIRPF
                            Case 0 'retencion sobre base + iva
                                ' le quito la retencion
                                vbase = Round2(vImporte / ((1 + (PorcIva / 100)) * (1 - (vParamAplic.PorcreteFacSoc / 100))), 2)
                                
                            Case 1 'retencion sobre base
                                ' le quito la retencion
                                vbase = Round2(vImporte / (1 + (PorcIva / 100) - (vParamAplic.PorcreteFacSoc / 100)), 2)
                            
                            Case 2 ' sin retencion
                                ' le quito el iva
                                vbase = Round2(vImporte / (1 + (PorcIva / 100)), 2)
                        End Select
                        
                End Select
            
            Else ' entradas normales
                Select Case Combo1(2).ListIndex
                    Case 0 'precio normal
                        vPrecio = CCur(ComprobarCero(Text1(27).Text))
                    
                    Case 1 'precio iva incluido con retencion
                        vPrecio = CCur(ComprobarCero(Text1(27).Text))
                        vImporte = Round2(vPrecio * TotalKilos, 2)
                        vtotfac = vImporte
                        Select Case vSocio.TipoIRPF
                            Case 0 'retencion sobre base + iva
                                ' le quito la retencion
                                vbase = Round2(vImporte / ((1 + (PorcIva / 100)) * (1 - (vParamAplic.PorcreteFacSoc / 100))), 2)
                                
                            Case 1 'retencion sobre base
                                ' le quito la retencion
                                vbase = Round2(vImporte / (1 + (PorcIva / 100) - (vParamAplic.PorcreteFacSoc / 100)), 2)
                            
                            Case 2 ' sin retencion
                                ' le quito el iva
                                vbase = Round2(vImporte / (1 + (PorcIva / 100)), 2)
                        End Select
                        
                    Case 2 'precio iva incluido sin retencion
                        vPrecio = CCur(ComprobarCero(Text1(27).Text))
                        vImporte = Round2(vPrecio * TotalKilos, 2)
                        'le añado la retencion
                        vImporte = Round2(vImporte * (1 - (vParamAplic.PorcreteFacSoc / 100)), 2)
                        vtotfac = vImporte
                        
                        Select Case vSocio.TipoIRPF
                            Case 0 'retencion sobre base + iva
                                ' le quito la retencion
                                vbase = Round2(vImporte / ((1 + (PorcIva / 100)) * (1 - (vParamAplic.PorcreteFacSoc / 100))), 2)
                                
                            Case 1 'retencion sobre base
                                ' le quito la retencion
                                vbase = Round2(vImporte / (1 + (PorcIva / 100) - (vParamAplic.PorcreteFacSoc / 100)), 2)
                            
                            Case 2 ' sin retencion
                                ' le quito el iva
                                vbase = Round2(vImporte / (1 + (PorcIva / 100)), 2)
                        End Select
                End Select
            End If
                
            Albaranes = ""
            
            For i = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(i).Checked Then
                    Albaranes = Albaranes & ListView1.ListItems(i).Text & ","
                    Ultimo = i
                End If
            Next i
            
            'limpiamos los albaranes
            If Albaranes <> "" Then
                Albaranes = Mid(Albaranes, 1, Len(Albaranes) - 1)
            
                If Combo1(1).ListIndex = 1 Then ' entradas ventacampo
                    Sql = "update rhisfruta set impentrada = 0,prestimado = 0 where numalbar in (" & Albaranes & ")"
                    conn.Execute Sql
                End If
            End If
            
            For i = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(i).Checked Then
                
                    Ultimo = i
                
                    If Combo1(1).ListIndex = 0 Then ' entradas normales
                        Select Case Combo1(2).ListIndex
                            Case 0 'precio normal
                                vPrecio = CCur(ComprobarCero(Text1(27).Text))

'                                '[Monica]01/07/2013: el precio segun albaran
'                                vPrecio = CCur(ListView1.ListItems(I).SubItems(6))
'                                '

                                vImporte = vPrecio * ListView1.ListItems(i).SubItems(5)

                            Case 1, 2 'precio iva incluido con retencion
                                vImporte = vbase * ListView1.ListItems(i).SubItems(5) / TotalKilos
                                vPrecio = vImporte / ListView1.ListItems(i).SubItems(5)
                        End Select

                        Sql = "update rhisfruta set prestimado = " & DBSet(vPrecio, "N") & ", impentrada = " & DBSet(vImporte, "N")
                        Sql = Sql & " where numalbar = " & DBSet(ListView1.ListItems(i).Text, "N")
                        conn.Execute Sql

                        ImpBruto = ImpBruto + vImporte 'Round2(DBSet(ListView1.ListItems(I).SubItems(5), "N") * vPrecio, 2)
                    
                    Else ' entradas ventacampo
                        vImporte = Round2(CCur(vbase) / TotalKilos * DBSet(ListView1.ListItems(i).SubItems(5), "N"), 2)
                        vPrecio = Round2(vImporte / DBSet(ListView1.ListItems(i).SubItems(5), "N"), 4)
                        
                        Sql = "update rhisfruta set prestimado = " & DBSet(vPrecio, "N") & ", impentrada = " & DBSet(Round2(DBSet(ListView1.ListItems(i).SubItems(5), "N") * vPrecio, 2), "N")
                        Sql = Sql & " where numalbar = " & DBSet(ListView1.ListItems(i).Text, "N")
                        conn.Execute Sql
                        
                        ImpBruto = ImpBruto + vImporte 'Round2(DBSet(ListView1.ListItems(I).SubItems(5), "N") * vPrecio, 2)
                    End If
                    
                    Sql = "(rhisfruta.numalbar=" & DBSet(ListView1.ListItems(i).Text, "N") & ") "
                    If cadAux = "" Then
                        cadAux = Sql
                    Else
                        cadAux = cadAux & " OR " & Sql
                    End If
                End If
            Next i
            
            If cadAux <> "" Then
            'se han seleccionado albaranes para facturar
            'Esta el la cadena WHERE de los albaranes seleccionados para obtener
            'el bruto de las lineas de los albaranes agrupadas por tipo de iva
                cadWHERE = "(1=1) "
                cadWHERE = cadWHERE & " AND (" & cadAux & ")"
            
                If Not SeleccionaRegistros Then Exit Sub
                
                If Not BloqueaRegistro("rhisfruta", cadWHERE) Then
                    CargarAlbaranes vWhere
                End If
                
                TerminaBloquear
                
                '[Monica]30/04/2013: el calculo de la factura es sobre el iva del socio
                Sql = "SELECT sum(impentrada) as bruto"
                Sql = Sql & " FROM rhisfruta "
                Sql = Sql & " WHERE " & cadWHERE
            
                BrutoFact = DevuelveValor(Sql)
                BaseImp = BrutoFact
                
                'Obtener el % de IVA
                cadAux = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                
                '---- Laura: 24/10/2006
                If cadAux = "" Then cadAux = "0"
                impiva = CalcularPorcentaje(BaseImp, CCur(cadAux), 2)
                
                TipoIVA = vSocio.CodIva
                        
                PorcIva = cadAux '% de IVA
                     
                
                ' dependiendo del tipoirpf del socio se calcula la retencion
                Select Case vSocio.TipoIRPF
                    Case 0
                        BaseReten = BaseImp + impiva
                    Case 1
                        BaseReten = BaseImp
                    Case 2
                        BaseReten = 0
                End Select
                
                ' calculo de la retencion
                PorcReten = CCur(ComprobarCero(vParamAplic.PorcreteFacSoc))
                ImpReten = Round2(BaseReten * PorcReten / 100, 2)
            
                'TOTAL de la factura
                TotalFac = BaseImp + impiva - ImpReten
                
                
                If TotalFac <> vtotfac And (Combo1(2).ListIndex <> 0) Then
                    Diferencia = (vtotfac - TotalFac)
                    TotalFac = TotalFac + Diferencia
                    BaseImp = BaseImp + Diferencia
                    BaseReten = BaseImp + impiva
                
                    ' lista de albaranes
                    If Albaranes <> "" Then
                        Sql = "select sum(impentrada) from rhisfruta where numalbar in (" & Albaranes & ")"
                        If BaseImp <> DevuelveValor(Sql) Then
                            Diferencia = BaseImp - DevuelveValor(Sql)
                            Sql = "update rhisfruta set impentrada = impentrada + " & DBSet(Diferencia, "N")
                            Sql = Sql & " where numalbar = " & DBSet(ListView1.ListItems(Ultimo).Text, "N")
                            conn.Execute Sql
                        
                            Sql = "update rhisfruta set prestimado = round(impentrada / kilosnet,4)"
                            Sql = Sql & " where numalbar = " & DBSet(ListView1.ListItems(Ultimo).Text, "N")
                            conn.Execute Sql
                        End If
                    End If
                End If
                
                'descontamos las baseimponibles de las variedades de lo albaranes de facturas de anticipos
                VisualizarAnticipos vSocio
                ' se recalcula el total factura
                
                BrutoFact = BaseImp
                Dto = 0
                If ComprobarCero(Text1(8).Text) <> 0 Then
                    Dto = CCur(TransformaPuntosComas(ImporteSinFormato(Text1(8).Text)))
                    
                    BaseImp = BaseImp - Dto
                    impiva = CalcularPorcentaje(BaseImp, CCur(cadAux), 2)
                    Select Case vSocio.TipoIRPF
                        Case 0
                            BaseReten = BaseImp + impiva
                        Case 1
                            BaseReten = BaseImp
                        Case 2
                            BaseReten = 0
                    End Select
                    
                    ' calculo de la retencion
                    PorcReten = CCur(ComprobarCero(vParamAplic.PorcreteFacSoc))
                    ImpReten = Round2(BaseReten * PorcReten / 100, 2)
                
                    'TOTAL de la factura
                    TotalFac = BaseImp + impiva - ImpReten
                End If
                
                
                'Hasta aqui
                Text1(6).Text = BrutoFact
                Text1(9).Text = BaseImp
                Text1(10).Text = TipoIVA
                Text1(11).Text = 0
                Text1(12).Text = 0
                Text1(13).Text = PorcIva
                Text1(14).Text = 0
                Text1(15).Text = 0
                Text1(16).Text = BaseImp
                Text1(17).Text = 0
                Text1(18).Text = 0
                Text1(19).Text = impiva
                Text1(20).Text = 0
                Text1(21).Text = 0
                Text1(22).Text = TotalFac
                Text1(23).Text = BaseReten
                Text1(25).Text = ImpReten
                If ImpReten = 0 Then
                    Text1(24).Text = 0
                Else
                    Text1(24).Text = PorcReten
                End If
                
                Check1(1).Value = 0
                
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
            End If
        End If
    End If
    Exit Sub
    
eCalcularDatosFactura:
    MuestraError Err.Number, "Calcular Datos Factura", Err.Description
End Sub



Private Sub CalcularDatosFacturaNew()
Dim i As Integer
Dim Sql As String
Dim cadAux As String
Dim ImpBruto As Currency
Dim impiva As Currency
Dim vFactu As CFacturaTer
Dim Rs As ADODB.Recordset
Dim Dto As Currency
Dim Dto1 As Currency
Dim TotalKilos As Currency
Dim vPrecio As Currency
Dim PorcIva As Currency
Dim vImporte As Currency
Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim Variedad As String
Dim TipoIVA As Integer

Dim BrutoFact As Currency
Dim BaseImp As Currency
Dim BaseIva As Currency
Dim BaseReten As Currency
Dim PorcReten As Currency
Dim ImpReten As Currency
Dim TotalFac As Currency
Dim Diferencia As Currency

Dim vbase As Currency
Dim Ultimo As Long
Dim vtotfac As Currency
Dim vtotcal As Currency
Dim Albaranes As String
Dim KilosAlb As Currency

' datos de concepto de cargo
Dim cImporte As Currency
Dim cCodIva As Integer
Dim cPorIva As Currency
Dim cBaseIva As Currency
Dim cImpIva As Currency

    On Error GoTo eCalcularDatosFacturaNew


    Set vSocio = New cSocio
    If vSocio.LeerDatos(Text1(3).Text) Then
        If vSocio.LeerDatosSeccion(Text1(3).Text, vParamAplic.Seccionhorto) Then
        
        
            Dto = 0
'            If Text1(8).Text <> "" Then
'                Dto = CCur(TransformaPuntosComas(ImporteSinFormato(Text1(8).Text)))
'            End If
            'Limpiar en el form los datos calculados de la factura
            'y volvemos a recalcular
            For i = 6 To 25
                 Text1(i).Text = ""
            Next i
        
            cadAux = ""
            cadWHERE = ""
            ImpBruto = 0
            
            vPrecio = 0
            vImporte = 0
            
            'calculo el total de kilos
            TotalKilos = 0
            For i = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(i).Checked Then
                    TotalKilos = TotalKilos + DBSet(ListView1.ListItems(i).SubItems(5), "N")
                End If
            Next i
            
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
                If vSeccion.AbrirConta Then
                    PorcIva = 0
                    PorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(vSocio.CodIva), "N")
                    
                    '[Monica]23/09/2013: si hay concepto de cargo
                    If ComprobarCero(Text1(33).Text) <> 0 Then
                        cImporte = ComprobarCero(Text1(33).Text)
                        cCodIva = DevuelveDesdeBDNew(cAgro, "fvarconce", "tipoiva", "codconce", Text1(34).Text, "N")
                        cPorIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(cCodIva), "N")
                        
                        cBaseIva = ComprobarCero(Text1(33).Text)
                        cImpIva = Round2(cBaseIva * cPorIva / 100, 2)
                    End If
                End If
            End If
            Set vSeccion = Nothing
                
                
            If Combo1(1).ListIndex = 1 Then ' entradas de venta campo
                Select Case Combo1(2).ListIndex
                    Case 0 'precio normal
                        vImporte = CCur(ComprobarCero(Text1(26).Text))
                        vbase = vImporte
                    Case 1 'precio iva incluido con retencion
                        vImporte = CCur(ComprobarCero(Text1(26).Text))
                        vtotfac = vImporte
                        Select Case vSocio.TipoIRPF
                            Case 0 'retencion sobre base + iva
                                ' le quito la retencion
                                vbase = Round2(vImporte / ((1 + (PorcIva / 100)) * (1 - (vParamAplic.PorcreteFacSoc / 100))), 2)
                                
                            Case 1 'retencion sobre base
                                ' le quito la retencion
                                vbase = Round2(vImporte / (1 + (PorcIva / 100) - (vParamAplic.PorcreteFacSoc / 100)), 2)
                            
                            Case 2 ' sin retencion
                                ' le quito el iva
                                vbase = Round2(vImporte / (1 + (PorcIva / 100)), 2)
                        End Select
                        
                    Case 2 'precio iva incluido sin retencion
                        vImporte = CCur(ComprobarCero(Text1(26).Text))
                        'le añado la retencion
                        vImporte = Round2(vImporte * (1 - (vParamAplic.PorcreteFacSoc / 100)), 2)
                        vtotfac = vImporte
                        Select Case vSocio.TipoIRPF
                            Case 0 'retencion sobre base + iva
                                ' le quito la retencion
                                vbase = Round2(vImporte / ((1 + (PorcIva / 100)) * (1 - (vParamAplic.PorcreteFacSoc / 100))), 2)
                                
                            Case 1 'retencion sobre base
                                ' le quito la retencion
                                vbase = Round2(vImporte / (1 + (PorcIva / 100) - (vParamAplic.PorcreteFacSoc / 100)), 2)
                            
                            Case 2 ' sin retencion
                                ' le quito el iva
                                vbase = Round2(vImporte / (1 + (PorcIva / 100)), 2)
                        End Select
                        
                End Select
            
            Else ' entradas normales
                
                Albaranes = ""
                
                For i = 1 To ListView1.ListItems.Count
                    If ListView1.ListItems(i).Checked Then
                        Albaranes = Albaranes & ListView1.ListItems(i).Text & ","
                        Ultimo = i
                    End If
                Next i
                
                'limpiamos los albaranes
                If Albaranes <> "" Then
                    Albaranes = Mid(Albaranes, 1, Len(Albaranes) - 1)
                End If
                    
                Select Case Combo1(2).ListIndex
                    Case 0 'precio normal
                        TotalKilos = 0
                        vImporte = 0
                        CalculoImporteTot Albaranes, TotalKilos, vImporte

                    Case 1 'precio iva incluido con retencion
                        TotalKilos = 0
                        vImporte = 0
                        CalculoImporteTot Albaranes, TotalKilos, vImporte
                        vtotfac = vImporte
                        Select Case vSocio.TipoIRPF
                            Case 0 'retencion sobre base + iva
                                ' le quito la retencion
                                vbase = Round2(vImporte / ((1 + (PorcIva / 100)) * (1 - (vParamAplic.PorcreteFacSoc / 100))), 2)

                            Case 1 'retencion sobre base
                                ' le quito la retencion
                                vbase = Round2(vImporte / (1 + (PorcIva / 100) - (vParamAplic.PorcreteFacSoc / 100)), 2)

                            Case 2 ' sin retencion
                                ' le quito el iva
                                vbase = Round2(vImporte / (1 + (PorcIva / 100)), 2)
                        End Select

                    Case 2 'precio iva incluido sin retencion
                        TotalKilos = 0
                        vImporte = 0
                        CalculoImporteTot Albaranes, TotalKilos, vImporte
                        ' se le añade la retencion
                        vImporte = Round2(vImporte * (1 - (vParamAplic.PorcreteFacSoc / 100)), 2)
                        vtotfac = vImporte

                        Select Case vSocio.TipoIRPF
                            Case 0 'retencion sobre base + iva
                                ' le quito la retencion
                                vbase = Round2(vImporte / ((1 + (PorcIva / 100)) * (1 - (vParamAplic.PorcreteFacSoc / 100))), 2)

                            Case 1 'retencion sobre base
                                ' le quito la retencion
                                vbase = Round2(vImporte / (1 + (PorcIva / 100) - (vParamAplic.PorcreteFacSoc / 100)), 2)

                            Case 2 ' sin retencion
                                ' le quito el iva
                                vbase = Round2(vImporte / (1 + (PorcIva / 100)), 2)
                        End Select
                End Select
            
            End If
                
            Albaranes = ""
            
            For i = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(i).Checked Then
                    Albaranes = Albaranes & ListView1.ListItems(i).Text & ","
                    Ultimo = i
                End If
            Next i
            
            'limpiamos los albaranes
            If Albaranes <> "" Then
                Albaranes = Mid(Albaranes, 1, Len(Albaranes) - 1)
            
                If Combo1(1).ListIndex = 1 Then ' entradas ventacampo
                    Sql = "update rhisfruta set impentrada = 0,prestimado = 0 where numalbar in (" & Albaranes & ")"
                    conn.Execute Sql
                End If
            End If
            
            For i = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(i).Checked Then
                
                    Ultimo = i
                
                    If Combo1(1).ListIndex = 0 Then ' entradas normales
                        Select Case Combo1(2).ListIndex
                            Case 0 'precio normal
                                KilosAlb = 0
                                CalculoImporteTot ListView1.ListItems(i).Text, KilosAlb, vImporte
                            
                                'DAVID###   explotaba por que el calculo al ser EOF daba 0 y dividir por cero overfow
                                'vPrecio = vImporte / KilosAlb  estaba asin
                                If KilosAlb = 0 Then
                                    vPrecio = 0
                                Else
                                    vPrecio = vImporte / KilosAlb
                                End If
                            Case 1, 2 'precio iva incluido con retencion
                                KilosAlb = 0
                                CalculoImporteTot ListView1.ListItems(i).Text, KilosAlb, vImporte
                                
                                vImporte = vbase * KilosAlb / TotalKilos
                                
                                'vPrecio = vImporte / KilosAlb  estaba asin
                                'DAVID###   explotaba por que el calculo al ser EOF daba 0 y dividir por cero overfow
                                If KilosAlb = 0 Then
                                    vPrecio = 0
                                Else
                                    vPrecio = vImporte / KilosAlb
                                End If
                        End Select

                        Sql = "update rhisfruta set prestimado = " & DBSet(vPrecio, "N") & ", impentrada = " & DBSet(vImporte, "N") & ", kilosfactu = " & DBSet(KilosAlb, "N")
                        Sql = Sql & " where numalbar = " & DBSet(ListView1.ListItems(i).Text, "N")
                        conn.Execute Sql

                        ImpBruto = ImpBruto + vImporte 'Round2(DBSet(ListView1.ListItems(I).SubItems(5), "N") * vPrecio, 2)
                    
                    Else ' entradas ventacampo
                        vImporte = Round2(CCur(vbase) / TotalKilos * DBSet(ListView1.ListItems(i).SubItems(5), "N"), 2)
                        vPrecio = Round2(vImporte / DBSet(ListView1.ListItems(i).SubItems(5), "N"), 4)
                        
                        Sql = "update rhisfruta set prestimado = " & DBSet(vPrecio, "N") & ", impentrada = " & DBSet(Round2(DBSet(ListView1.ListItems(i).SubItems(5), "N") * vPrecio, 2), "N")
                        Sql = Sql & " kilosfactu = " & DBSet(ListView1.ListItems(i).SubItems(5), "N")
                        Sql = Sql & " where numalbar = " & DBSet(ListView1.ListItems(i).Text, "N")
                        conn.Execute Sql
                        
                        ImpBruto = ImpBruto + vImporte 'Round2(DBSet(ListView1.ListItems(I).SubItems(5), "N") * vPrecio, 2)
                    End If
                    
                    Sql = "(rhisfruta.numalbar=" & DBSet(ListView1.ListItems(i).Text, "N") & ") "
                    If cadAux = "" Then
                        cadAux = Sql
                    Else
                        cadAux = cadAux & " OR " & Sql
                    End If
                End If
            Next i
            
            If cadAux <> "" Then
            'se han seleccionado albaranes para facturar
            'Esta el la cadena WHERE de los albaranes seleccionados para obtener
            'el bruto de las lineas de los albaranes agrupadas por tipo de iva
                cadWHERE = "(1=1) "
                cadWHERE = cadWHERE & " AND (" & cadAux & ")"
            
                If Not SeleccionaRegistros Then Exit Sub
                
                If Not BloqueaRegistro("rhisfruta", cadWHERE) Then
                    CargarAlbaranes vWhere
                End If
                
                TerminaBloquear
                
                '[Monica]30/04/2013: el calculo de la factura es sobre el iva del socio
                Sql = "SELECT sum(impentrada) as bruto"
                Sql = Sql & " FROM rhisfruta "
                Sql = Sql & " WHERE " & cadWHERE
            
                BrutoFact = DevuelveValor(Sql)
                BaseImp = BrutoFact
                
                'Obtener el % de IVA
                cadAux = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                
                '---- Laura: 24/10/2006
                If cadAux = "" Then cadAux = "0"
                impiva = CalcularPorcentaje(BaseImp, CCur(cadAux), 2)
                
                TipoIVA = vSocio.CodIva
                        
                PorcIva = cadAux '% de IVA
                     
                
                ' dependiendo del tipoirpf del socio se calcula la retencion
                Select Case vSocio.TipoIRPF
                    Case 0
                        BaseReten = BaseImp + impiva
                    Case 1
                        BaseReten = BaseImp
                    Case 2
                        BaseReten = 0
                End Select
                
                ' calculo de la retencion
                PorcReten = CCur(ComprobarCero(vParamAplic.PorcreteFacSoc))
                ImpReten = Round2(BaseReten * PorcReten / 100, 2)
            
                'TOTAL de la factura
                TotalFac = BaseImp + impiva - ImpReten
                
                
                If TotalFac <> vtotfac And (Combo1(2).ListIndex <> 0) Then
                    Diferencia = (vtotfac - TotalFac)
                    TotalFac = TotalFac + Diferencia
                    BaseImp = BaseImp + Diferencia
                    BaseReten = BaseImp + impiva
                
                    ' lista de albaranes
                    If Albaranes <> "" Then
                        Sql = "select sum(impentrada) from rhisfruta where numalbar in (" & Albaranes & ")"
                        If BaseImp <> DevuelveValor(Sql) Then
                            Diferencia = BaseImp - DevuelveValor(Sql)
                            Sql = "update rhisfruta set impentrada = impentrada + " & DBSet(Diferencia, "N")
                            Sql = Sql & " where numalbar = " & DBSet(ListView1.ListItems(Ultimo).Text, "N")
                            conn.Execute Sql
                        
                            Sql = "update rhisfruta set prestimado = round(impentrada / kilosnet,4)"
                            Sql = Sql & " where numalbar = " & DBSet(ListView1.ListItems(Ultimo).Text, "N")
                            conn.Execute Sql
                        End If
                    End If
                End If
                
                'descontamos las baseimponibles de las variedades de lo albaranes de facturas de anticipos
                VisualizarAnticipos vSocio
                ' se recalcula el total factura
                
                ' el descuento
                Text1(32).Text = ""
                If Text1(31).Text <> "" Then
                    Text1(32).Text = Format(Text1(31).Text, "###,###,##0.00")
                End If
                
                BrutoFact = BaseImp + cImporte
                Dto = 0
                Dto1 = 0
                If ComprobarCero(Text1(8).Text) <> 0 Or ComprobarCero(Text1(32).Text) <> 0 Then
                    Dto = CCur(TransformaPuntosComas(ImporteSinFormato(ComprobarCero(Text1(8).Text))))
                    Dto1 = CCur(TransformaPuntosComas(ImporteSinFormato(ComprobarCero(Text1(32).Text))))
                    
                    BaseImp = BaseImp - Dto - Dto1
                End If

                '[Monica]31/10/2013: he subido el endif desde '***(1)
                    impiva = CalcularPorcentaje(BaseImp, CCur(cadAux), 2)
                    Select Case vSocio.TipoIRPF
                        Case 0
                            BaseReten = BaseImp + impiva + cBaseIva + cImpIva
                        Case 1
                            BaseReten = BaseImp + cBaseIva
                        Case 2
                            BaseReten = 0
                    End Select
                    
                    ' calculo de la retencion
                    PorcReten = CCur(ComprobarCero(vParamAplic.PorcreteFacSoc))
                    ImpReten = Round2(BaseReten * PorcReten / 100, 2)
                
                    'TOTAL de la factura
                    TotalFac = BaseImp + impiva - ImpReten + cBaseIva + cImpIva
                 '***(1)
'                End If
                 
                '[Monica]23/09/2013: calculo de cargo
'                If cImporte <> 0 Then
'                    TotalFac = TotalFac + cImporte + cImpIva
'                    BaseImp = BaseImp + cImporte
'                End If
                
                
                
                'Hasta aqui
                Text1(6).Text = BrutoFact
                Text1(9).Text = BaseImp + cBaseIva
                Text1(10).Text = TipoIVA
                Text1(11).Text = 0
                Text1(12).Text = 0
                Text1(13).Text = PorcIva
                Text1(14).Text = 0
                Text1(15).Text = 0
                Text1(16).Text = BaseImp
                Text1(17).Text = 0
                Text1(18).Text = 0
                Text1(19).Text = impiva
                Text1(20).Text = 0
                Text1(21).Text = 0
                Text1(22).Text = TotalFac
                Text1(23).Text = BaseReten
                Text1(25).Text = ImpReten
                If ImpReten = 0 Then
                    Text1(24).Text = 0
                Else
                    Text1(24).Text = PorcReten
                End If
                
                Check1(1).Value = 0
                
                '[Monica]23/09/2013: si hay cargo
                If cImporte <> 0 Then
                    If cCodIva = TipoIVA Then
                        Text1(16).Text = BaseImp + cBaseIva
                        Text1(19).Text = impiva + cImpIva
                    Else
                        Text1(11).Text = cCodIva
                        Text1(14).Text = cPorIva
                        Text1(17).Text = cBaseIva
                        Text1(20).Text = cImpIva
                    End If
                End If
                
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
            End If
        '[Monica]03/11/2014: Añadido el else por si el socio no tiene secciones
        Else
            MsgBox "El socio no tiene sección. Revise.", vbExclamation
            
        End If
    End If
    Exit Sub
    
eCalcularDatosFacturaNew:
    MuestraError Err.Number, "Calcular Datos Factura", Err.Description
End Sub




Private Function CalculoImporteTot(Albaran As String, TotalKilos As Currency, TotalImporte As Currency) As Currency
Dim Sql As String
Dim Rs As ADODB.Recordset
    
    On Error Resume Next

    If Albaran = "" Then Exit Function

    '[Monica]09/01/2019: para el caso de frutas inma, sobre kilos transportados
    If vParamAplic.Cooperativa = 18 Then
        '[Monica]19/09/2013: calculamos el precio normal
        If Combo1(2).ListIndex = 0 Then
            Sql = "select kilostra kilosnet, /*round(precio * kilosnet,2)*/ importe from rhisfruta_clasif where numalbar in (" & Albaran & ")"
        Else
            Sql = "select kilostra kilosnet, round(precio * kilostra,2) importe from rhisfruta_clasif where numalbar in (" & Albaran & ")"
        End If
    Else

        '[Monica]19/09/2013: calculamos el precio normal
        If Combo1(2).ListIndex = 0 Then
            Sql = "select kilosnet, /*round(precio * kilosnet,2)*/ importe from rhisfruta_clasif where numalbar in (" & Albaran & ")"
        Else
            Sql = "select kilosnet, round(precio * kilosnet,2) importe from rhisfruta_clasif where numalbar in (" & Albaran & ")"
        End If
    '[Monica]22/04/2015: cogemos todas las lineas marcadas hayan o no puesto precio
    '    SQL = SQL & " and precio <> 0 and not precio is null "
    End If
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    TotalKilos = 0
    TotalImporte = 0
    While Not Rs.EOF
        TotalKilos = TotalKilos + DBLet(Rs!KilosNet, "N")
        TotalImporte = TotalImporte + DBLet(Rs!Importe, "N")
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing

End Function


Private Sub CalcularDatosFacturaSinEntradas()
Dim i As Integer
Dim Sql As String
Dim cadAux As String
Dim ImpBruto As Currency
Dim impiva As Currency
Dim vFactu As CFacturaTer
Dim Rs As ADODB.Recordset
Dim Dto As Currency
Dim Dto2 As Currency

Dim TotalKilos As Currency
Dim vPrecio As Currency
Dim PorcIva As Currency
Dim vImporte As Currency
Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim Variedad As String
Dim TipoIVA As Integer

Dim BrutoFact As Currency
Dim BaseImp As Currency
Dim BaseIva As Currency
Dim BaseReten As Currency
Dim PorcReten As Currency
Dim ImpReten As Currency
Dim TotalFac As Currency

Dim vbase As Currency
Dim Diferencia As Currency

    Set vSocio = New cSocio
    If vSocio.LeerDatos(Text1(3).Text) Then
        If vSocio.LeerDatosSeccion(Text1(3).Text, vParamAplic.Seccionhorto) Then
            
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
                If vSeccion.AbrirConta Then
                    PorcIva = 0
                    PorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(vSocio.CodIva), "N")
                End If
            End If
            Set vSeccion = Nothing
        
            Dto = 0
            If Text1(8).Text <> "" Then
                Dto = CCur(TransformaPuntosComas(ImporteSinFormato(Text1(8).Text)))
            End If
            
            '[Monica]07/10/2013
            Dto2 = 0
            If Text1(31).Text <> "" Then
                Text1(32).Text = Text1(31).Text
                Dto2 = CCur(TransformaPuntosComas(ImporteSinFormato(Text1(32).Text)))
            End If
            
            
            'Limpiar en el form los datos calculados de la factura
            'y volvemos a recalcular
            For i = 6 To 25
                 Text1(i).Text = ""
            Next i
        
            cadAux = ""
            cadWHERE = ""
            ImpBruto = 0
            
            vImporte = ComprobarCero(Text1(26).Text) - Dto2
            
            Select Case Combo1(2).ListIndex
                Case 0 'precio normal
                    vImporte = ComprobarCero(Text1(26).Text) - Dto2
                    vbase = vImporte
                
                Case 1 'precio iva incluido con retencion
                    vImporte = ComprobarCero(Text1(26).Text) - Dto2
                    
                    Select Case vSocio.TipoIRPF
                        Case 0 'retencion sobre base + iva
                            ' le quito la retencion
                            vbase = Round2(vImporte / ((1 + (PorcIva / 100)) * (1 - (vParamAplic.PorcreteFacSoc / 100))), 2)
                            
                        Case 1 'retencion sobre base
                            ' le quito la retencion
                            vbase = Round2(vImporte / (1 + (PorcIva / 100) - (vParamAplic.PorcreteFacSoc / 100)), 2)
                        
                        Case 2 ' sin retencion
                            ' le quito el iva
                            vbase = Round2(vImporte / (1 + (PorcIva / 100)), 2)
                    End Select
                    
                Case 2 'precio iva incluido sin retencion
                    vImporte = ComprobarCero(Text1(26).Text) - Dto2
                    'le añado la retencion
                    vImporte = Round2(vImporte * (1 - (vParamAplic.PorcreteFacSoc / 100)), 2)
                    
                    Select Case vSocio.TipoIRPF
                        Case 0 'retencion sobre base + iva
                            ' le quito la retencion
                            vbase = Round2(vImporte / ((1 + (PorcIva / 100)) * (1 - (vParamAplic.PorcreteFacSoc / 100))), 2)
                            
                        Case 1 'retencion sobre base
                            ' le quito la retencion
                            vbase = Round2(vImporte / (1 + (PorcIva / 100) - (vParamAplic.PorcreteFacSoc / 100)), 2)
                        
                        Case 2 ' sin retencion
                            ' le quito el iva
                            vbase = Round2(vImporte / (1 + (PorcIva / 100)), 2)
                    End Select
                    
            End Select
            
            BrutoFact = vbase
            BaseImp = BrutoFact
            
            'Obtener el % de IVA
            cadAux = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
            
            '---- Laura: 24/10/2006
            If cadAux = "" Then cadAux = "0"
            impiva = CalcularPorcentaje(BaseImp, CCur(cadAux), 2)
            
            TipoIVA = vSocio.CodIva
                    
            PorcIva = cadAux '% de IVA
                 
            ' dependiendo del tipoirpf del socio se calcula la retencion
            Select Case vSocio.TipoIRPF
                Case 0
                    BaseReten = BaseImp + impiva
                Case 1
                    BaseReten = BaseImp
                Case 2
                    BaseReten = 0
            End Select
            
            ' calculo de la retencion
            PorcReten = CCur(ComprobarCero(vParamAplic.PorcreteFacSoc))
            ImpReten = Round2(BaseReten * PorcReten / 100, 2)
        
            'TOTAL de la factura
            TotalFac = BaseImp + impiva - ImpReten
            
            If TotalFac <> vImporte And (Combo1(2).ListIndex <> 0) And Combo1(0).ListIndex = 1 Then
                Diferencia = (vImporte - TotalFac)
                TotalFac = TotalFac + Diferencia
                BaseImp = BaseImp + Diferencia
                BaseReten = BaseImp + impiva
            End If
            
             BrutoFact = BaseImp
            ' si es liquidacion descontamos anticipos
            If Combo1(0).ListIndex = 0 Then
            
                VisualizarAnticipos vSocio
            
                BrutoFact = BaseImp
                Dto = 0
                If ComprobarCero(Text1(8).Text) <> 0 Then
                    Dto = CCur(TransformaPuntosComas(ImporteSinFormato(Text1(8).Text)))
                    
                    BaseImp = BaseImp - Dto
                    impiva = CalcularPorcentaje(BaseImp, CCur(cadAux), 2)
                    Select Case vSocio.TipoIRPF
                        Case 0
                            BaseReten = BaseImp + impiva
                        Case 1
                            BaseReten = BaseImp
                        Case 2
                            BaseReten = 0
                    End Select
                    
                    ' calculo de la retencion
                    PorcReten = CCur(ComprobarCero(vParamAplic.PorcreteFacSoc))
                    ImpReten = BaseReten * PorcReten / 100
                
                    'TOTAL de la factura
                    TotalFac = Round2(BaseImp + impiva - ImpReten, 2)
                    
'                    If TotalFac <> CCur(Text1(26).Text) - totalanticipos Then
'
'                    End If
                
                    If (vImporte - TotalFac) <> TotalFactAnticipo And Combo1(2).ListIndex <> 0 Then
                        Diferencia = (vImporte - TotalFac - TotalFactAnticipo)
                        TotalFac = TotalFac + Diferencia
                        BaseImp = BaseImp + Diferencia
                        BaseReten = BaseImp + impiva
                    End If
                Else
                    If TotalFac <> vImporte And (Combo1(2).ListIndex <> 0) Then
                        Diferencia = (vImporte - TotalFac)
                        TotalFac = TotalFac + Diferencia
                        BaseImp = BaseImp + Diferencia
                        BrutoFact = BaseImp
                        BaseReten = BaseImp + impiva
                    End If
                End If
            End If
            
            
            'Hasta aqui
            Text1(6).Text = BrutoFact
            Text1(9).Text = BaseImp
            Text1(10).Text = TipoIVA
            Text1(11).Text = 0
            Text1(12).Text = 0
            Text1(13).Text = PorcIva
            Text1(14).Text = 0
            Text1(15).Text = 0
            Text1(16).Text = BaseImp
            Text1(17).Text = 0
            Text1(18).Text = 0
            Text1(19).Text = impiva
            Text1(20).Text = 0
            Text1(21).Text = 0
            Text1(22).Text = TotalFac
            Text1(23).Text = BaseReten
            Text1(25).Text = ImpReten
            If ImpReten = 0 Then
                Text1(24).Text = 0
            Else
                Text1(24).Text = PorcReten
            End If
            
            Check1(1).Value = 0
            
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
        End If
    End If
    
End Sub

Private Function SeleccionaRegistros() As Boolean
'Comprueba que se seleccionan albaranes en la base de datos
'es decir que hay albaranes marcados
'cuando se van marcando albaranes se van añadiendo el la cadena cadWhere
Dim Sql As String

    On Error GoTo ESel
    SeleccionaRegistros = False
    
    If cadWHERE = "" Then Exit Function
    
    Sql = "Select count(*) FROM rhisfruta"
    Sql = Sql & " WHERE " & cadWHERE
    If RegistrosAListar(Sql) <> 0 Then SeleccionaRegistros = True
    Exit Function
    
ESel:
    SeleccionaRegistros = False
    MuestraError Err.Number, "No hay seleccionados Albaranes", Err.Description
End Function


Private Sub BotonFacturar()
Dim vFactu As CFacturaTer
Dim Cad As String
Dim CadFact As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    Cad = ""
    If Text1(3).Text = "" Then
        Cad = "Falta socio"
    Else
        If Not IsNumeric(Text1(3).Text) Then Cad = "Campo socio debe ser numérico"
    End If
    If Cad <> "" Then
        MsgBox Cad, vbExclamation
        Exit Sub
    End If
        
    Set vSocio = New cSocio
    
    'Tiene que ller los datos del transportista
    If Not vSocio.LeerDatos(Text1(3).Text) Then Exit Sub
    
    If Not DatosOkFact Then
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
    vFactu.BrutoFac = ImporteFormateado(Text1(9).Text) + ImporteFormateado(Text1(32).Text) 'ImporteFormateado(Text1(6).Text) + ImporteFormateado(Text1(32).Text)
    vFactu.ForPago = Text1(4).Text
    vFactu.DtoPPago = 0
    vFactu.DtoGnral = 0
    vFactu.ImpPPago = ImporteFormateado(Text1(7).Text)
    vFactu.ImpGnral = ImporteFormateado(Text1(32).Text)
        
'    vFactu.ImpGnral = ImporteFormateado(Text1(8).Text)
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
    vFactu.CCC_Entidad = ""
    vFactu.CCC_Oficina = ""
    If CInt(vSocio.Banco) <> 0 Then vFactu.CCC_Entidad = CInt(vSocio.Banco)
    If CInt(vSocio.Sucursal) <> 0 Then vFactu.CCC_Oficina = CInt(vSocio.Sucursal)
    vFactu.CCC_CC = vSocio.Digcontrol
    vFactu.CCC_CTa = vSocio.CuentaBan
    
    vFactu.Intracomunitario = Check1(1).Value
    
    vFactu.EsAnticipo = Combo1(0).ListIndex
    vFactu.Anticipos = Anticipos
    
    '[Monica]23/09/2013: guardamos el concepto y el importe de cargo si lo hay en la factura de tercero
    vFactu.ConcepCargo = ComprobarCero(Text1(34).Text)
    vFactu.ImpCargo = ComprobarCero(Text1(33).Text)
    
    '[Monica]27/05/2013: insertamos cual es el porcentaje de corredor para hacer mas de un pago si es distinto de 0
    vFactu.PorcCorredor = ImporteFormateado(ComprobarCero(Text1(28).Text))
'    frmTercRecFact.Check1(0).Value = 0

    ' sacamos la cuenta de proveedor
    If Not vSocio.LeerDatosSeccion(vSocio.Codigo, vParamAplic.Seccionhorto) Then
        MsgBox "No se han encontrado los datos del socio de la sección Hortofrutícola", vbExclamation
        Set vFactu = Nothing
        Exit Sub
    End If
    
    vFactu.CtaTerce = vSocio.CtaProv
    cadFormula = ""
    cadSelect = ""
'    If cadWhere <> "" Then
        If Not vSocio.Estercero(Text1(3).Text, True) Then
            If InsertarFacturaSocioAntLiq(vSocio) Then
                BotonImprimir
                BotonPedirDatos False
            End If
        Else
            vFactu.Variedad = Text1(29).Text
            If vFactu.TraspasoAlbaranesAFactura(cadWHERE) Then BotonPedirDatos False
        End If
'    End If
    
    Set vFactu = Nothing
    
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub BotonImprimir()
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim nTabla As String

    CadParam = ""
    numParam = 0
    
    '[Monica]28/01/2014: preguntamos si quiere imprimir arrobas
    If MsgBox("¿ Desea impresión con Arrobas ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
        CadParam = CadParam & "|pConArrobas=1|"
    Else
        CadParam = CadParam & "|pConArrobas=0|"
    End If
    numParam = numParam + 1
    
    
    indRPT = 23 'Impresion de facturas de socios
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
    'Nombre fichero .rpt a Imprimir
    cadNombreRPT = nomDocu
    'Nombre fichero .rpt a Imprimir
    cadTitulo = "Impresión de Factura Socio"
    ConSubInforme = True

    LlamarImprimir

    If frmVisReport.EstaImpreso Then
        ActualizarRegistrosFac "rfactsoc", cadSelect
    End If

End Sub

Private Function ActualizarRegistrosFac(cTabla As String, cWhere As String) As Boolean
'Actualizar la marca de impreso
Dim Sql As String

    On Error GoTo eActualizarRegistros

    ActualizarRegistrosFac = False
    Sql = "update " & cTabla & ", usuarios.stipom set impreso = 1 "
    Sql = Sql & " where usuarios.stipom.codtipom = rfactsoc.codtipom "
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " and " & cWhere
    End If
    
    conn.Execute Sql
    
    ActualizarRegistrosFac = True
    Exit Function
    
eActualizarRegistros:
    MuestraError Err.Number, "Actualizando registros", Err.Description
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



Private Function InsertarFacturaSocioAntLiq(vSocio As cSocio) As Boolean
Dim Sql As String
Dim vSeccion As CSeccion
Dim CuentaPrev As String
Dim b As Boolean
Dim tipoMov As String
Dim vTipoMov As CTiposMov
Dim numfactu As Long
Dim devuelve As String
Dim Existe As Boolean
Dim MenError As String
Dim Albaranes As String
Dim i As Long

    On Error GoTo eInsertarFacturaSocioAntLiq

    InsertarFacturaSocioAntLiq = False

    conn.BeginTrans


    Set vSeccion = New CSeccion
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If vSeccion.AbrirConta Then
            ConnConta.BeginTrans
        Else
            Exit Function
        End If
    End If

'    If Combo1(0).ListIndex = 1 Then
'        If Combo1(1).ListIndex = 0 Then
'            tipoMov = vSocio.CodTipomAnt
'        Else
'            tipoMov = vSocio.CodTipomAntVC
'        End If
'    Else
'        If Combo1(1).ListIndex = 0 Then
'            tipoMov = vSocio.CodTipomLiq
'        Else
'            tipoMov = vSocio.CodTipomLiqVC
'        End If
'    End If
    
    '[Monica]01/10/2018: para el resto de cooperativas, frutas inma y castelduc
    If vParamAplic.Cooperativa = 12 Or vParamAplic.Cooperativa = 18 Then
        tipoMov = vSocio.CodTipomLiq
    Else
        If Combo1(0).ListIndex = 1 Then
            If Combo1(1).ListIndex = 0 Then
                tipoMov = vSocio.CodTipomAnt
            Else
                tipoMov = vSocio.CodTipomAntVC
            End If
        Else
            If Combo1(1).ListIndex = 0 Then
                tipoMov = vSocio.CodTipomLiq
            Else
                tipoMov = vSocio.CodTipomLiqVC
            End If
        End If
    End If
    
    
    Set vTipoMov = New CTiposMov
    
    numfactu = vTipoMov.ConseguirContador(tipoMov)
    Do
        numfactu = vTipoMov.ConseguirContador(tipoMov)
        devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", Text1(1).Text, "F")
        If devuelve <> "" Then
            'Ya existe el contador incrementarlo
            Existe = True
            vTipoMov.IncrementarContador (tipoMov)
            numfactu = vTipoMov.ConseguirContador(tipoMov)
        Else
            Existe = False
        End If
    Loop Until Not Existe
                

    'Cuenta Prevista de Cobro de las Facturas
    CuentaPrev = DevuelveDesdeBDNew(cAgro, "banpropi", "codmacta", "codbanpr", Text1(5).Text, "N")

    'Insertar la Factura
    MenError = "Insertar Cabecera de Factura"
    b = InsertarCabecera(vSocio, tipoMov, CStr(numfactu), Text1(1).Text, MenError)

    'si se trata de una liquidacion con albaranes
    If Combo1(0).ListIndex = 0 And Combo1(1).ListIndex = 0 Then
    
        Albaranes = ""
        
        For i = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(i).Checked Then
                Albaranes = Albaranes & ListView1.ListItems(i).Text & ","
            End If
        Next i
        ' lista de albaranes
        If Albaranes <> "" Then
            Albaranes = Mid(Albaranes, 1, Len(Albaranes) - 1)
        End If
        
    
        MenError = "Insertar Variedades"
        If b Then b = InsertarLineasFactura(tipoMov, CStr(numfactu), Text1(1).Text, Albaranes, MenError)
    
        MenError = "Insertar Albaranes"
        If b Then b = InsertarLineasAlbaranes(tipoMov, CStr(numfactu), Text1(1).Text, Albaranes, MenError)
    
        MenError = "Insertar Anticipos"
        If b Then
            If TotalFactAnticipo <> 0 Then
                b = InsertarAnticipos(tipoMov, CStr(numfactu), Text1(1).Text, vSocio, MenError)
            End If
        End If
        
        MenError = "Insertar Descuentos"
        If b Then
            If ComprobarCero(Text1(32).Text) <> 0 Then
                b = InsertarGastos(tipoMov, CStr(numfactu), Text1(1).Text, vSocio, MenError)
            End If
        End If


    Else
    ' caso de liquidacion sin albaranes
        MenError = "Insertar Variedades"
        If b Then b = InsertarLineasFacturaSinEntradas(tipoMov, CStr(numfactu), Text1(1).Text, Text1(29).Text, MenError)
    
        ' si es liquidacion puede que tenga anticipos
        If b Then
            If Combo1(0).ListIndex = 0 Then
                If TotalFactAnticipo <> 0 Then
                    b = InsertarAnticipos(tipoMov, CStr(numfactu), Text1(1).Text, vSocio, MenError)
                End If
            End If
        End If
        
        
        '[Monica]22/10/2013
        MenError = "Insertar Descuentos"
        If b Then
            If ComprobarCero(Text1(32).Text) <> 0 Then
                b = InsertarGastos(tipoMov, CStr(numfactu), Text1(1).Text, vSocio, MenError)
            End If
        End If
    
    End If

    ' actualizar contador
    If b Then b = vTipoMov.IncrementarContador(tipoMov)



    If b Then
        cadFormula = "{rfactsoc.numfactu} = " & DBSet(numfactu, "N") & " and {rfactsoc.codtipom} = """ & tipoMov & """"
        cadFormula = cadFormula & " and {rfactsoc.fecfactu}= Date(" & Year(CDate(Text1(1).Text)) & "," & Month(CDate(Text1(1).Text)) & "," & Day(CDate(Text1(1).Text)) & ")"
        
        cadSelect = "rfactsoc.numfactu=" & DBSet(numfactu, "N") & " and rfactsoc.codtipom = " & DBSet(tipoMov, "T")
        cadSelect = cadSelect & " and rfactsoc.fecfactu= " & DBSet(Text1(1).Text, "F")
    End If
    
    

eInsertarFacturaSocioAntLiq:
    If Err.Number <> 0 Then b = False
    If b Then
        conn.CommitTrans
        ConnConta.CommitTrans
        InsertarFacturaSocioAntLiq = True
        MsgBox "La Factura de Socio de los Albaranes seleccionados se generó correctamente.", vbInformation
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        InsertarFacturaSocioAntLiq = False
        MsgBox "ATENCIÓN:" & vbCrLf & "La Factura NO se generó correctamente!!!." & vbCrLf & MenError, vbInformation
    End If

    vSeccion.CerrarConta
    Set vSeccion = Nothing

End Function


Private Function InsertarAnticipos(tipoMov As String, numfactu As String, FecFac As String, vSocio As cSocio, MenError As String) As Boolean
'(rfactsoc_albaran)
'codcampo tiene valor cuando venimos de almazara que hemos tenido que buscarlo porque en el cursor Rs no lo tenemos
Dim Tipo As String

    Dim Sql As String
    Dim ImpLinea As Currency
    
    On Error GoTo eInsertLinea
    
    MensError = ""
    
    InsertarAnticipos = False
    
    'insertamos el albaran
    If vSocio.TipoProd = 0 Then ' socio
        Sql = "insert into rfactsoc_anticipos (codtipom, numfactu, fecfactu, codtipomanti,numfactuanti,fecfactuanti,codvarieanti,codcampoanti,baseimpo) "
        Sql = Sql & " select " & DBSet(tipoMov, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ", codtipom, numfactu, fecfactu, basereten,0, baseimpo "
        Sql = Sql & " from tmprfactsoc "
        Sql = Sql & " where codusu = " & vUsu.Codigo
    
        conn.Execute Sql
    
        ' he de marcar los anticipos que acabo de descontar
        Sql = "update rfactsoc_variedad set descontado = 1 "
        Sql = Sql & " where (codtipom, numfactu, fecfactu, codvarie) in (select codtipom, numfactu, fecfactu, basereten from tmprfactsoc where codusu = " & vUsu.Codigo & ")"
        
        conn.Execute Sql
            
    
    Else 'tercero
        Sql = "insert into rliantifter (codsocio, numfactu, fecfactu, codsocioanti, numfactuanti, fecfactuanti) "
        Sql = Sql & " select " & DBSet(vSocio.Codigo, "N") & "," & DBSet(numfactu, "T") & "," & DBSet(FecFac, "F") & "," & DBSet(vSocio.Codigo, "N")
        Sql = Sql & " codsocio, fecfactu "
        Sql = Sql & " from tmprfactsoc "
        Sql = Sql & " where codusu = " & vUsu.Codigo
    
        conn.Execute Sql
    
        '[Monica]25/09/2013: descontamos los anticipos que hemos marcado
        Sql = "update rlifter set descontado = 1 "
        Sql = Sql & " where (codsocio, numfactu, fecfactu, codvarie) in (select codtipom, numfactu, fecfactu, basereten from tmprfactsoc where codusu = " & vUsu.Codigo & ")"
    
        conn.Execute Sql
    
    End If
    
    
    
    
    
    
    
    InsertarAnticipos = True
    
    Exit Function
    
eInsertLinea:
    If Err.Number <> 0 Then
        MensError = "Se ha producido un error en la inserción de anticipos de factura "
        MuestraError Err.Number, MensError, Err.Description
    End If
End Function





Private Function InsertarLineasAlbaranes(tipoMov As String, numfactu As String, FecFac As String, Albaranes As String, MenError As String) As Boolean
'(rfactsoc_albaran)
'codcampo tiene valor cuando venimos de almazara que hemos tenido que buscarlo porque en el cursor Rs no lo tenemos
Dim GastosAlb As Currency
Dim Tipo As String

    Dim Sql As String
    Dim ImpLinea As Currency
    
    On Error GoTo eInsertLinea
    
    MensError = ""
    
    InsertarLineasAlbaranes = False
    
    
    'insertamos el albaran
    Sql = "insert into rfactsoc_albaran (codtipom, numfactu, fecfactu, numalbar, fecalbar, "
    Sql = Sql & "codvarie, codcampo, kilosbru, kilosnet, grado, precio, importe, imporgasto)  "
'    sql = sql & "select '" & Trim(tipoMov) & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ", numalbar, fecalbar, rhisfruta.codvarie, 0, rhisfruta.kilosbru, rhisfruta.kilosfactu, "
'    sql = sql & " prestimado,impentrada,0"
'    sql = sql & " from  rhisfruta "
'    sql = sql & " where rhisfruta.numalbar in (" & Albaranes & ")"
    
    
    Sql = Sql & " select " & DBSet(tipoMov, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ", rhisfruta_clasif.numalbar, rhisfruta.fecalbar, rhisfruta.codvarie, "
    '[Monica]07/01/2019: antes el codcampo era siempre 0, pq montifrut no usa campos
    '[Monica]24/01/2019: si es frutas inma los kilos son los transportados
    If vParamAplic.Cooperativa = 18 Then
        Sql = Sql & " rhisfruta.codcampo, sum(rhisfruta_clasif.kilostra), sum(rhisfruta_clasif.kilostra), round(sum(precio) / count(*),4), round(sum(precio) / count(*),4), round(sum(rhisfruta_clasif.kilostra) * round(sum(precio) / count(*),4),2), 0  "
    Else
        Sql = Sql & " rhisfruta.codcampo, sum(rhisfruta_clasif.kilosnet), sum(rhisfruta_clasif.kilosnet), round(sum(precio) / count(*),4), round(sum(precio) / count(*),4), round(sum(rhisfruta_clasif.kilosnet) * round(sum(precio) / count(*),4),2), 0  "
    End If
    Sql = Sql & " from rhisfruta_clasif inner join rhisfruta on rhisfruta_clasif.numalbar = rhisfruta.numalbar "
    Sql = Sql & " and rhisfruta.codvarie = rhisfruta_clasif.codvarie "
    Sql = Sql & " where rhisfruta.numalbar in (" & Albaranes & ")"
    Sql = Sql & " group by 1,2,3,4,5,6,7 "
    Sql = Sql & " order by 1,2,3,4,5,6,7 "
    
    
    
    
    conn.Execute Sql
    
    InsertarLineasAlbaranes = True
    
    Exit Function
    
eInsertLinea:
    If Err.Number <> 0 Then
        MensError = "Se ha producido un error en la inserción de las lineas de albaranes de factura "
        MuestraError Err.Number, MensError, Err.Description
    End If
End Function







Private Function InsertarLineasFactura(tipoMov As String, numfactu As String, FecFac As String, Albaranes As String, MenError As String) As Boolean
Dim Precio As Currency
Dim Rs As ADODB.Recordset
Dim CadValues As String
Dim Sql As String

    
    On Error GoTo eInsertLinea
    
    InsertarLineasFactura = False
    
    MensError = ""
    Precio = 0
    
    '[Monica]01/10/2018: metemos el codigo de campo
    Sql = "select codvarie, codcampo, sum(kilosfactu) kilos, sum(impentrada) importe from rhisfruta where numalbar in (" & Albaranes & ")"
    Sql = Sql & " group by 1,2 order by 1,2"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CadValues = ""
    
    While Not Rs.EOF
        If CCur(ImporteSinFormato(Rs!Kilos)) <> 0 Then
            Precio = Round2(CCur(ImporteSinFormato(Rs!Importe)) / CCur(ImporteSinFormato(Rs!Kilos)), 4)
        End If
    
        CadValues = CadValues & "('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
        CadValues = CadValues & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!codCampo, "N") & "," '[Monica]01/10/2018: antes 0
        CadValues = CadValues & DBSet(ImporteSinFormato(Rs!Kilos), "N") & "," & DBSet(Precio, "N") & ","
        CadValues = CadValues & DBSet(ImporteSinFormato(Rs!Importe), "N")
        CadValues = CadValues & ",0,0," & DBSet(Text1(27).Text, "N") & " ),"
    
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    If CadValues <> "" Then
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        
        Sql = "insert into rfactsoc_variedad (codtipom, numfactu, fecfactu, codvarie, codcampo, "
        Sql = Sql & "kilosnet, preciomed, imporvar, imporgasto, kilogrado, preciorea) values "
        
        conn.Execute Sql & CadValues
    End If
    
    Sql = "insert into rfactsoc_calidad(codtipom,numfactu,fecfactu,codvarie,codcampo,codcalid,kilosnet,precio,imporcal) "
    '[Monica]11/07/2013: cambiamos el calculo
'    sql = sql & " select codtipom,numfactu,fecfactu,rfactsoc_variedad.codvarie,codcampo,max(codcalid),kilosnet,preciomed,imporvar "
'    sql = sql & " from rfactsoc_variedad inner join rcalidad on rfactsoc_variedad.codvarie = rcalidad.codvarie "
'    sql = sql & " where codtipom = " & DBSet(tipoMov, "T")
'    sql = sql & " and numfactu = " & DBSet(numfactu, "N")
'    sql = sql & " and fecfactu = " & DBSet(FecFac, "F")
'    sql = sql & " group by 1,2,3,4,5,7,8,9 "
'    sql = sql & " order by 1,2,3,4,5,7,8,9 "

    '[Monica]01/10/2018: antes era 0 codcampo
    '[Monica]24/01/2019: para el caso de frutas inma
    If vParamAplic.Cooperativa = 18 Then
        Sql = Sql & " select " & DBSet(tipoMov, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ", rhisfruta_clasif.codvarie, rhisfruta.codcampo, codcalid, sum(rhisfruta_clasif.kilostra), round(sum(rhisfruta_clasif.precio) / count(*),4), round(sum(rhisfruta_clasif.kilostra) * round(sum(rhisfruta_clasif.precio) / count(*),4),2)  "
    Else
        Sql = Sql & " select " & DBSet(tipoMov, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ", rhisfruta_clasif.codvarie, rhisfruta.codcampo, codcalid, sum(rhisfruta_clasif.kilosnet), round(sum(rhisfruta_clasif.precio) / count(*),4), round(sum(rhisfruta_clasif.kilosnet) * round(sum(rhisfruta_clasif.precio) / count(*),4),2)  "
    End If
    
    Sql = Sql & " from rhisfruta_clasif inner join rhisfruta on rhisfruta_clasif.numalbar = rhisfruta.numalbar "
    Sql = Sql & " where rhisfruta_clasif.numalbar in (" & Albaranes & ")"
    Sql = Sql & " group by 1,2,3,4,5,6 "
    Sql = Sql & " order by 1,2,3,4,5,6 "

    conn.Execute Sql
    
    '[Monica]01/10/2018: añadimos esto
    If vParamAplic.Cooperativa = 5 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19 Then
        Sql = "update rfactsoc_calidad set preciocalidad = precio, imporcalidad = imporcal where codtipom = " & DBSet(tipoMov, "T")
        Sql = Sql & " and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
        
        conn.Execute Sql
    End If
    
    InsertarLineasFactura = True
    Exit Function
    
eInsertLinea:
    If Err.Number <> 0 Then
        MensError = "Se ha producido un error en la inserción de las lineas de factura"
        MuestraError Err.Number, MensError, Err.Description
    End If
End Function


Private Function InsertarLineasFacturaSinEntradas(tipoMov As String, numfactu As String, FecFac As String, Variedad As String, MenError As String) As Boolean
Dim Precio As Currency
Dim Rs As ADODB.Recordset
Dim CadValues As String
Dim Sql As String

    
    On Error GoTo eInsertLinea
    
    InsertarLineasFacturaSinEntradas = False
    
    MensError = ""
    Precio = 0
    
    CadValues = ""
    
    CadValues = CadValues & "('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
    CadValues = CadValues & DBSet(Variedad, "N") & ",0,"
    CadValues = CadValues & DBSet(0, "N") & "," & DBSet(0, "N") & ","
    CadValues = CadValues & DBSet(Text1(6).Text, "N")
    CadValues = CadValues & ",0,0,0)"
    
        
    Sql = "insert into rfactsoc_variedad (codtipom, numfactu, fecfactu, codvarie, codcampo, "
    Sql = Sql & "kilosnet, preciomed, imporvar, imporgasto, kilogrado, preciorea) values "
    
    conn.Execute Sql & CadValues
    
    Sql = "insert into rfactsoc_calidad(codtipom,numfactu,fecfactu,codvarie,codcampo,codcalid,kilosnet,precio,imporcal) "
    Sql = Sql & " select codtipom,numfactu,fecfactu,rfactsoc_variedad.codvarie,codcampo,max(codcalid),kilosnet,preciomed,imporvar "
    Sql = Sql & " from rfactsoc_variedad inner join rcalidad on rfactsoc_variedad.codvarie = rcalidad.codvarie "
    Sql = Sql & " where codtipom = " & DBSet(tipoMov, "T")
    Sql = Sql & " and numfactu = " & DBSet(numfactu, "N")
    Sql = Sql & " and fecfactu = " & DBSet(FecFac, "F")
    Sql = Sql & " group by 1,2,3,4,5,7,8,9 "
    Sql = Sql & " order by 1,2,3,4,5,7,8,9 "
    conn.Execute Sql
    
    
    InsertarLineasFacturaSinEntradas = True
    Exit Function
    
eInsertLinea:
    If Err.Number <> 0 Then
        MensError = "Se ha producido un error en la inserción de las lineas de factura sin entradas"
        MuestraError Err.Number, MensError, Err.Description
    End If
End Function





Private Function InsertarCabecera(vSocio As cSocio, TipoM As String, NumFact As String, FecFact As String, MenError As String) As Boolean
Dim Sql As String
Dim PorcIva As Currency
Dim EsAnticipo As Byte
Dim EsVtaCampo As Byte

    On Error GoTo eInsertCabe
    
    MensError = ""
    InsertarCabecera = False

    EsAnticipo = Combo1(0).ListIndex
    EsVtaCampo = Combo1(1).ListIndex

    Sql = "insert into rfactsoc (codtipom, numfactu, fecfactu, codsocio, baseimpo, tipoiva, porc_iva,"
    Sql = Sql & "imporiva, tipoirpf, basereten, porc_ret, impreten, baseaport, porc_apo, impapor, totalfac, impreso, contabilizado, pasaridoc,"
    Sql = Sql & "esanticipogasto, esretirada, esliqcomplem, codforpa, porccorredor, tipoprecio) "
    Sql = Sql & " values ('" & TipoM & "'," & DBSet(NumFact, "N") & "," & DBSet(FecFact, "F") & "," & DBSet(vSocio.Codigo, "N") & ","
    
    PorcIva = 0
    PorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(vSocio.CodIva), "N")

    Sql = Sql & DBSet(Text1(9).Text, "N") & "," & vSocio.CodIva & "," & DBSet(PorcIva, "N") & ","
    
    Sql = Sql & DBSet(Text1(19).Text, "N") & "," & DBSet(vSocio.TipoIRPF, "N") & "," & DBSet(Text1(23).Text, "N") & ","
    Sql = Sql & DBSet(Text1(24).Text, "N") & "," & DBSet(Text1(25).Text, "N") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & DBSet(Text1(22).Text, "N") & ","
    Sql = Sql & "0,0,0,"
    
    '0,0,0,"
    '[Monica]01/10/2018:caso de frutas inma y castelduc
    If vParamAplic.Cooperativa = 12 Or vParamAplic.Cooperativa = 18 Then
        Sql = Sql & DBSet(EsAnticipo, "N") & "," & DBSet(EsVtaCampo, "N") & ",0,"
    Else
        Sql = Sql & "0,0,0,"
    End If
    
    
    Sql = Sql & DBSet(Text1(4).Text, "N") & "," & DBSet(Text1(28).Text, "N") & "," & Combo1(2).ListIndex & ")"
    
    conn.Execute Sql
    
    InsertarCabecera = True
    
    Exit Function

eInsertCabe:
    MensError = "Error en la inserción en rfactsoc de la factura " & NumFact & " del socio " & vSocio.Codigo
    MuestraError Err.Number, MensError

End Function


Private Function ExisteFacturaEnHco() As Boolean
'Comprobamos si la factura ya existe en la tabla de Facturas a Proveedor: scafpc
Dim Cad As String

    ExisteFacturaEnHco = False
    'Tiene que tener valor los 3 campos de clave primaria antes de comprobar
    If Not (Text1(0).Text <> "" And Text1(1).Text <> "" And Text1(3).Text <> "") Then Exit Function
    
    ' No debe existir el número de factura para el proveedor en hco
    Cad = "SELECT count(*) FROM rcafter "
    Cad = Cad & " WHERE codsocio=" & Text1(3).Text & " AND numfactu=" & DBSet(Text1(0).Text, "T") & " AND year(fecfactu)=" & Year(Text1(2).Text)
    If RegistrosAListar(Cad) > 0 Then
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
    
    If Adoaux(Index).Recordset.EOF Then Exit Sub
    If Adoaux(Index).Recordset.RecordCount < 1 Then Exit Sub
    
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
Dim b As Boolean
Dim i As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    CargaGridGnral Me.DataGridAux(Index), Me.Adoaux(Index), tots, PrimeraVez
    
    
    Select Case Index
        Case 0 'rhisfruta
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "S|txtaux(0)|T|Albarán|750|;S|txtaux(1)|T|Fecha|950|;"
            tots = tots & "S|txtaux(6)|T|Código|660|;S|Text2(2)|T|Variedad|1420|;"
            tots = tots & "S|txtaux(3)|T|Kilos Neto|1000|;S|txtaux(2)|T|Pr.Estim.|850|;"
            tots = tots & "S|txtaux(4)|T|Importe|1100|;N|txtaux(5)|T|Socio|1100|;"
            arregla tots, DataGridAux(Index), Me
        
            DataGridAux(0).Columns(1).Alignment = dbgLeft
            DataGridAux(0).Columns(5).Alignment = dbgRight
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            BloquearTxt txtAux(3), Not b

    End Select
    DataGridAux(Index).ScrollBars = dbgAutomatic
    
    If Not Adoaux(0).Recordset.EOF Then
        Me.lblIndicador.Caption = Adoaux(0).Recordset.AbsolutePosition & " de " & Adoaux(0).Recordset.RecordCount
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
Dim Sql As String
Dim Tabla As String
   
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
        Case 0 'historico de entradas
            Tabla = "rhisfruta"
            Sql = "SELECT rhisfruta.numalbar,rhisfruta.fecalbar, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.kilosnet, rhisfruta.prestimado, rhisfruta.impentrada, rhisfruta.codsocio "
            Sql = Sql & " FROM " & Tabla & " inner join variedades on rhisfruta.codvarie = variedades.codvarie "
            If enlaza Then
'                SQL = SQL & ObtenerWhereCab(True)
                Sql = Sql & " where codsocio =  " & DBSet(Text1(3).Text, "N")
                
                
                '[Monica] 04/02/2010 Todos los albaranes o solo los que no han sido cobrados
                If Check1(2).Value = 0 Then
                    Sql = Sql & " and cobradosn = 0 "   ' que no esten cobradas
                End If
                    
                If Text1(26).Text <> "" Then
                    Sql = Sql & " and rhisfruta.codvarie = " & DBSet(Text1(26).Text, "N")
                End If
            Else
                Sql = Sql & " WHERE numalbar  = -1"
            End If
            
            Sql = Sql & " ORDER BY " & Tabla & ".numalbar,  " & Tabla & ".fecalbar "
            
    End Select
    ' ********************************************************************************
    
    MontaSQLCarga = Sql
End Function


Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    If Not DatosOk Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If vSocio Is Nothing Then Set vSocio = New cSocio
    
    If vSocio.LeerDatos(Text1(3).Text) Then

        vWhere = "rhisfruta.codsocio = " & DBSet(Text1(3).Text, "N")
        If vSocio.Cooperativa <> 1 Then vWhere = vWhere & " or rhisfruta.codsocio in (select codsocio from rsocios where codcoope = " & vSocio.Cooperativa & ")"
        
        ' Las entradas que sean del tipo indicado
        vWhere = "(" & vWhere & ") and rhisfruta.tipoentr = " & Combo1(1).ListIndex
        
        If Not BloqueaRegistro("rhisfruta", vWhere) Then
            MsgBox "No se puede recepcionar factura de ese socio. Hay otro usuario recepcionando.", vbExclamation
            BotonPedirDatos True
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            If LimpiarImportes(vWhere) Then
                '--monica:080908
                TerminaBloquear
                If Not BloqueaRegistro("rhisfruta", vWhere) Then
                    MsgBox "No se puede recepcionar factura de ese socio. Hay otro usuario recepcionando.", vbExclamation
                    BotonPedirDatos True
                    Screen.MousePointer = vbDefault
                    Exit Sub
                Else
                    '--
                    
                    If Combo1(0).ListIndex = 1 Or Combo1(1).ListIndex = 1 Then '  And Combo1(1).ListIndex = 1) Then ' liquidacion sin entradas
                        PonerModo 5
                        
                        If Combo1(0).ListIndex = 0 Then ListaAnticipos
                        
                        CalcularDatosFacturaSinEntradas
                    Else
                        CargarAlbaranes vWhere
                
                        If ListView1.ListItems.Count <> 0 Then
                
'                            VisualizarAnticipos vSocio
                            If Combo1(0).ListIndex = 0 Then ListaAnticipos
                            
                            TerminaBloquear
                            
                            CalcularDatosFacturaNew
                            
                        End If
                    End If
                End If
            End If
        End If
    End If
    Set vSocio = Nothing
    
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub ListaAnticipos()
Dim Sql As String

    If vSocio.TipoProd = 1 Then
        Sql = "select distinct rcafter.numfactu, rcafter.fecfactu, variedades.nomvarie, rcafter.baseiva1 from (rcafter inner join rlifter on rcafter.codsocio = rlifter.codsocio and rcafter.numfactu = rlifter.numfactu and "
        Sql = Sql & " rcafter.fecfactu = rlifter.fecfactu) inner join variedades on rlifter.codvarie = variedades.codvarie  "
        Sql = Sql & " where rcafter.codsocio = " & vSocio.Codigo
        Sql = Sql & " and rcafter.esanticipo = 1 " ' sea un anticipo
        Sql = Sql & " and rlifter.descontado = 0 "
    

    Else
        Sql = "select rfactsoc.numfactu, rfactsoc.fecfactu,  variedades.nomvarie, rfactsoc.baseimpo from (rfactsoc inner join rfactsoc_variedad on rfactsoc.codtipom = rfactsoc_variedad.codtipom and "
        Sql = Sql & " rfactsoc.numfactu = rfactsoc_variedad.numfactu and rfactsoc.fecfactu = rfactsoc_variedad.fecfactu) inner join variedades on rfactsoc_variedad.codvarie = variedades.codvarie "
        Sql = Sql & " where rfactsoc.codsocio = " & DBSet(Text1(3).Text, "N")
        If Combo1(1).ListIndex = 1 Then
            'Sql = Sql & " and rfactsoc.codtipom = 'FAC' "
            Sql = Sql & " and rfactsoc.esanticipogasto = 1 and rfactsoc.esretirada = 1"
        Else
            'Sql = Sql & " and rfactsoc.codtipom = 'FAA' "
            Sql = Sql & " and rfactsoc.esanticipogasto = 1 and rfactsoc.esretirada = 0"
        End If
        Sql = Sql & " and rfactsoc_variedad.descontado = 0"

    End If

    If TotalRegistrosConsulta(Sql) = 0 Then Exit Sub


    Set frmMens = New frmMensajes
    
    frmMens.OpcionMensaje = 48
    frmMens.cadWHERE = Sql
    frmMens.Show vbModal
    
    Set frmMens = Nothing


End Sub

Private Sub VisualizarAnticipos(vSocio As cSocio)
Dim Sql As String
Dim Sql2 As String
Dim Variedades As String
Dim vVar As String
Dim Anticipo As Currency
Dim i As Integer


    On Error GoTo eVisualizarAnticipos
    


    TotalFactAnticipo = 0


    'si es venta campo cojo la variedad introducida ,text1(29)
    If Combo1(1).ListIndex = 1 Then
        Variedades = Text1(29).Text
    Else
        Variedades = ""
        
        For i = 1 To ListView1.ListItems.Count
            vVar = ""
            If ListView1.ListItems(i).Checked Then
                vVar = DevuelveValor("select codvarie from rhisfruta where numalbar = " & ListView1.ListItems(i).Text)
                Variedades = Variedades & vVar & ","
            End If
        Next i
        
        If Variedades <> "" Then Variedades = Mid(Variedades, 1, Len(Variedades) - 1)
    End If
    
    ' si es tercero
    If vSocio.TipoProd = 1 Then
        Sql = "select sum(rlifter.importel) impanticipo from rcafter inner join rlifter on rcafter.codsocio = rlifter.codsocio and rcafter.numfactu = rlifter.numfactu and "
        Sql = Sql & " rcafter.fecfactu = rlifter.fecfactu "
        Sql = Sql & " where rcafter.codsocio = " & vSocio.Codigo
        If Variedades <> "" Then Sql = Sql & " and rlifter.codvarie in (" & Variedades & ")"
        Sql = Sql & " and rcafter.esanticipo = 1 " ' sea un anticipo
        Sql = Sql & " and rlifter.descontado = 0 "

        If Anticipos <> "" Then
            Sql = Sql & " and (rcafter.numfactu, rcafter.fecfactu) in " & Anticipos
        Else
            Sql = Sql & " and (rcafter.numfactu, rcafter.fecfactu) = (null,null) "
        End If
    
        ' total anticipado
        Sql2 = "select sum(rcafter.totalfac) importe from rcafter inner join rlifter on rcafter.codsocio = rlifter.codsocio and rcafter.numfactu = rlifter.numfactu and "
        Sql2 = Sql2 & " rcafter.fecfactu = rlifter.fecfactu "
        Sql2 = Sql2 & " where rcafter.codsocio = " & vSocio.Codigo
        If Variedades <> "" Then Sql = Sql & " and rlifter.codvarie in (" & Variedades & ")"
        Sql2 = Sql2 & " and rcafter.esanticipo = 1 " ' sea un anticipo
        Sql2 = Sql2 & " and rlifter.descontado = 0 "
        
        If Anticipos <> "" Then
            Sql2 = Sql2 & " and (rcafter.numfactu, rcafter.fecfactu) in " & Anticipos
        Else
            Sql2 = Sql2 & " and (rcafter.numfactu, rcafter.fecfactu) = (null,null) "
        End If

        TotalFactAnticipo = DevuelveValor(Sql2)


        conn.Execute "delete from tmprfactsoc where codusu = " & vUsu.Codigo

        Sql2 = "insert into tmprfactsoc (codusu, codsocio,fecfactu, baseimpo) "
        Sql2 = Sql2 & "select distinct " & vUsu.Codigo & ", rcafter.numfactu, rcafter.fecfactu, rcafter.baseiva1 from rcafter inner join rlifter on rcafter.codsocio = rlifter.codsocio and rcafter.numfactu = rlifter.numfactu and "
        Sql2 = Sql2 & " rcafter.fecfactu = rlifter.fecfactu "
        Sql2 = Sql2 & " where rcafter.codsocio = " & vSocio.Codigo
        If Variedades <> "" Then Sql = Sql & " and rlifter.codvarie in (" & Variedades & ")"
        Sql2 = Sql2 & " and rcafter.esanticipo = 1 " ' sea un anticipo
        Sql2 = Sql2 & " and rlifter.descontado = 0 "
        
        If Anticipos <> "" Then
            Sql2 = Sql2 & " and (rcafter.numfactu, rcafter.fecfactu) in " & Anticipos
        Else
            Sql2 = Sql2 & " and (rcafter.numfactu, rcafter.fecfactu) = (null,null) "
        End If

        conn.Execute Sql2


    Else
        Sql = "select sum(rfactsoc_variedad.imporvar) impanticipo from rfactsoc inner join rfactsoc_variedad on rfactsoc.codtipom = rfactsoc_variedad.codtipom and "
        Sql = Sql & " rfactsoc.numfactu = rfactsoc_variedad.numfactu and rfactsoc.fecfactu = rfactsoc_variedad.fecfactu "
        Sql = Sql & " where rfactsoc.codsocio = " & DBSet(Text1(3).Text, "N")
        If Combo1(1).ListIndex = 1 Then
            'Sql = Sql & " and rfactsoc.codtipom = 'FAC' "
            Sql = Sql & " and rfactsoc.esanticipogasto = 1 and rfactsoc.esretirada = 1"
        Else
            'Sql = Sql & " and rfactsoc.codtipom = 'FAA' "
            Sql = Sql & " and rfactsoc.esanticipogasto = 1 and rfactsoc.esretirada = 0"
        End If
        Sql = Sql & " and rfactsoc_variedad.descontado = 0"
        If Variedades <> "" Then Sql = Sql & " and rfactsoc_variedad.codvarie in (" & Variedades & ")"
        
        If Anticipos <> "" Then
            Sql = Sql & " and (rfactsoc.numfactu, rfactsoc.fecfactu) in " & Anticipos
        Else
            Sql = Sql & " and (rfactsoc.numfactu, rfactsoc.fecfactu) = (null,null) "
        End If
        
        
        
        Sql2 = "select sum(rfactsoc.totalfac) importe from rfactsoc inner join rfactsoc_variedad on rfactsoc.codtipom = rfactsoc_variedad.codtipom and "
        Sql2 = Sql2 & " rfactsoc.numfactu = rfactsoc_variedad.numfactu and rfactsoc.fecfactu = rfactsoc_variedad.fecfactu "
        Sql2 = Sql2 & " where rfactsoc.codsocio = " & DBSet(Text1(3).Text, "N")
        If Combo1(1).ListIndex = 1 Then
'            Sql2 = Sql2 & " and rfactsoc.codtipom = 'FAC' "
            Sql2 = Sql2 & " and rfactsoc.esanticipogasto = 1 and rfactsoc.esretirada = 1"
        Else
'            Sql2 = Sql2 & " and rfactsoc.codtipom = 'FAA' "
            Sql2 = Sql2 & " and rfactsoc.esanticipogasto = 1 and rfactsoc.esretirada = 0"
        End If
        Sql2 = Sql2 & " and rfactsoc_variedad.descontado = 0"
        If Variedades <> "" Then Sql2 = Sql2 & " and rfactsoc_variedad.codvarie in (" & Variedades & ")"

        If Anticipos <> "" Then
            Sql2 = Sql2 & " and (rfactsoc.numfactu, rfactsoc.fecfactu) in " & Anticipos
        Else
            Sql2 = Sql2 & " and (rfactsoc.numfactu, rfactsoc.fecfactu) = (null,null) "
        End If



        TotalFactAnticipo = DevuelveValor(Sql2)
        
        
        ' metemos los anticipos
        conn.Execute "delete from tmprfactsoc where codusu = " & vUsu.Codigo

        Sql2 = "insert into tmprfactsoc (codusu,codtipom, numfactu ,fecfactu, baseimpo, basereten) "
        Sql2 = Sql2 & " select distinct " & vUsu.Codigo & ",rfactsoc.codtipom, rfactsoc.numfactu, rfactsoc.fecfactu, rfactsoc_variedad.imporvar, rfactsoc_variedad.codvarie from rfactsoc inner join rfactsoc_variedad on rfactsoc.codtipom = rfactsoc_variedad.codtipom and "
        Sql2 = Sql2 & " rfactsoc.numfactu = rfactsoc_variedad.numfactu and rfactsoc.fecfactu = rfactsoc_variedad.fecfactu "
        Sql2 = Sql2 & " where rfactsoc.codsocio = " & DBSet(Text1(3).Text, "N")
        If Combo1(1).ListIndex = 1 Then
'            Sql2 = Sql2 & " and rfactsoc.codtipom = 'FAC' "
            Sql2 = Sql2 & " and rfactsoc.esanticipogasto = 1 and rfactsoc.esretirada = 1"
        Else
'            Sql2 = Sql2 & " and rfactsoc.codtipom = 'FAA' "
            Sql2 = Sql2 & " and rfactsoc.esanticipogasto = 1 and rfactsoc.esretirada = 0"
        End If
        
        Sql2 = Sql2 & " and rfactsoc_variedad.descontado = 0"
        If Variedades <> "" Then Sql2 = Sql2 & " and rfactsoc_variedad.codvarie in (" & Variedades & ")"
        
        If Anticipos <> "" Then
            Sql2 = Sql2 & " and (rfactsoc.numfactu, rfactsoc.fecfactu) in " & Anticipos
        Else
            Sql2 = Sql2 & " and (rfactsoc.numfactu, rfactsoc.fecfactu) = (null,null) "
        End If
        
        
        conn.Execute Sql2
        
        
    End If

    Anticipo = DevuelveValor(Sql)
    Text1(8).Text = Format(Anticipo, "###,###,##0.00")
    
    Exit Sub
    
eVisualizarAnticipos:
    MuestraError Err.Number, "Visualizar Anticipos", Err.Description
    
End Sub
Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdCancelar_Click()
'    If Modo = 3 Then
'        PonerModo 0
'    Else
        LimpiarCampos
        PonerModo 0
'    End If
    Screen.MousePointer = vbDefault
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
Dim Cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    Cad = "numalbar = " & Adoaux(0).Recordset!NumAlbar
    ' ***************************************
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    If SituarData(Adoaux(0), Cad, Indicador) Then
        lblIndicador.Caption = Indicador
    End If
    ' ***********************************************************************************
End Sub


Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    If Index <> 3 Then DeseleccionaGrid DataGridAux(Index)
    ' ***************************************************
       
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 0 'albaranes
            txtAux(0).visible = False
            txtAux(1).visible = False
            txtAux(2).visible = False
            txtAux(3).visible = False
            For jj = 4 To 4
                txtAux(jj).visible = b
                txtAux(jj).Top = alto
            Next jj
            
            Text2(2).visible = False
            
            
    End Select
End Sub


Private Sub ModificarLinea()
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Long
Dim Cad As String
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
            If cadWHERE <> "" Then BloqueaRegistro "rhisfruta", cadWHERE

            If NumTabMto <> 3 Then
                V = Adoaux(NumTabMto).Recordset.Fields(0) 'el 2 es el nº de llinia
                CargaGrid NumTabMto, True
            End If

            ' *** si n'hi han tabs ***
'            SituarTab (NumTabMto)

            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            If NumTabMto <> 3 Then
                DataGridAux(NumTabMto).SetFocus
                Adoaux(NumTabMto).Recordset.Find (Adoaux(NumTabMto).Recordset.Fields(0).Name & " =" & V)
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
    If Not Adoaux(i).Recordset.EOF Then _
        PonerCamposForma2 Me, Adoaux(i), 2, "FrameAux" & i
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
End Sub


Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim b As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte
    
    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False

    b = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not b Then Exit Function
    
    DatosOkLlin = b

EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub PonerModoOpcionesMenu(Modo)
'Actives unes Opcions de Menú i Toolbar según el modo en que estem
Dim b As Boolean, bAux As Boolean
Dim i As Byte
    
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
    b = (Modo = 3 Or Modo = 4 Or Modo = 2)
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.Adoaux(i).Recordset.RecordCount > 0)
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

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index <> 0 And KeyCode <> 38 Then KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
    
    If KeyAscii = 13 Then
        txtAux_LostFocus (2)
    End If
    
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim TipoDto As Byte
Dim ImpDto As String
Dim Unidades As String
Dim cantidad As String

    If Index = 2 Then
        If PonerFormatoDecimal(txtAux(Index), 8) Then
'            'actualizarRegistro
'            Sql = "update rhisfruta set prestimado = " & DBSet(txtaux(2).Text, "N")
'            Sql = Sql & " where numalbar = " & DBSet(txtaux(2).ToolTipText, "N")
'
'            conn.Execute Sql
            
            ListView1.SelectedItem.SubItems(6) = txtAux(2).Text
        End If
        
        txtAux(2).visible = False
        txtAux(2).Enabled = False
        
'        ListView1.Refresh

        CalcularDatosFacturaNew
        
        Exit Sub
    
    End If
    

    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtAux(Index), ModoLineas) Then Exit Sub
    
    Select Case Index
        Case 4 ' Importe
            If txtAux(Index).Text <> "" Then PonerFormatoDecimal txtAux(Index), 1
            
            
        Case 2
            If PonerFormatoDecimal(txtAux(Index), 8) Then
                'actualizarRegistro
'                Sql = "update rhisfruta set prestimado = " & DBSet(txtaux(2).Text, "N")
'                Sql = Sql & " where numalbar = " & DBSet(txtaux(2).ToolTipText, "N")
'
'                conn.Execute Sql
                
                txtAux(2).visible = False
                txtAux(2).Enabled = False
            
            End If
            
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
                                

Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 0 To Combo1.Count - 1
        Combo1(i).Clear
    Next i
    
    'tipo de factura
    Combo1(0).AddItem "Liquidación"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Anticipo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
    
    'tipo de entrada
    Combo1(1).AddItem "Normal"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Venta Campo"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
   
    'tipo de Precio
    Combo1(2).AddItem "Normal"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 0
    Combo1(2).AddItem "Iva Inc. c/Ret."
    Combo1(2).ItemData(Combo1(2).NewIndex) = 1
    Combo1(2).AddItem "Iva Inc. s/Ret."
    Combo1(2).ItemData(Combo1(2).NewIndex) = 2
  
End Sub


Private Function InsertarGastos(tipoMov As String, numfactu As String, FecFac As String, vSocio As cSocio, MenError As String) As Boolean
'(rfactsoc_albaran)
'codcampo tiene valor cuando venimos de almazara que hemos tenido que buscarlo porque en el cursor Rs no lo tenemos
Dim Tipo As String

    Dim Sql As String
    Dim ImpLinea As Currency
    
    On Error GoTo eInsertLinea
    
    MensError = ""
    
    InsertarGastos = False
    
    'insertamos el albaran
    If vSocio.TipoProd = 0 Then ' socio
        Sql = "insert into rfactsoc_gastos (codtipom, numfactu, fecfactu, numlinea,codgasto,importe) "
        Sql = Sql & " values (" & DBSet(tipoMov, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ",1," & DBSet(Text1(30).Text, "N") & "," & DBSet(Text1(32).Text, "N") & ")"
    
        conn.Execute Sql
    
    Else 'tercero
        Sql = "update rcafter set impppago = " & DBSet(Text1(32).Text, "N") & ", brutofac = brutofac + " & DBSet(Text1(32).Text, "N")
        Sql = Sql & " where codsocio = " & DBSet(vSocio.Codigo, "N") & " and numfactu = " & DBSet(numfactu, "T") & " and fecfactu = " & DBSet(FecFac, "F")
    
        conn.Execute Sql
    
    End If
    
    InsertarGastos = True
    
    Exit Function
    
eInsertLinea:
    If Err.Number <> 0 Then
        MensError = "Se ha producido un error en la inserción de descuentos de factura "
        MuestraError Err.Number, MensError, Err.Description
    End If
End Function


Private Sub imgCheck_Click(Index As Integer)
Dim b As Boolean
Dim TotalArray As Integer
Dim item As ListItem
Dim Sql As String

    
    'En el listview1
    b = Index = 1
    For TotalArray = 1 To ListView1.ListItems.Count
        ListView1.ListItems(TotalArray).Checked = b
        
        Set item = ListView1.ListItems(TotalArray)
        
'        ListView1_ItemCheck (It.Text)
        Sql = "update rhisfruta_clasif set precio = " & DBSet(Text1(27).Text, "N")
        If vParamAplic.Cooperativa = 18 Then
            '[Monica]19/09/2013: hemos añadido el importe para que puedan modificarlo   [Monica]06/03/2014:faltaba el 2 del round
            Sql = Sql & " , importe = round(kilostra * " & DBSet(Text1(27).Text, "N") & ",2)"
        Else
            '[Monica]19/09/2013: hemos añadido el importe para que puedan modificarlo   [Monica]06/03/2014:faltaba el 2 del round
            Sql = Sql & " , importe = round(kilosnet * " & DBSet(Text1(27).Text, "N") & ",2)"
        End If
        Sql = Sql & " where numalbar = " & DBSet(item.Text, "N")
        conn.Execute Sql

        CalcularDatosFacturaNew

        
        If (TotalArray Mod 50) = 0 Then DoEvents
    Next TotalArray
End Sub
