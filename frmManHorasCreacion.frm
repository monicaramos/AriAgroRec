VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManHorasCreacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Creación Masiva de Horas por cuadrilla"
   ClientHeight    =   11085
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   19125
   Icon            =   "frmManHorasCreacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11085
   ScaleWidth      =   19125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   135
      TabIndex        =   45
      Top             =   90
      Width           =   1335
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   46
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
               Object.ToolTipText     =   "Generar Entradas Horas"
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   2
      Left            =   17055
      TabIndex        =   35
      Text            =   "Text3"
      Top             =   9810
      Width           =   1245
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   1
      Left            =   15660
      TabIndex        =   34
      Text            =   "Text3"
      Top             =   9810
      Width           =   1290
   End
   Begin VB.Frame FrameIntro 
      Height          =   2040
      Left            =   135
      TabIndex        =   13
      Top             =   810
      Width           =   18720
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
         Left            =   13455
         TabIndex        =   54
         Text            =   "12345678901234567890"
         Top             =   1530
         Width           =   4920
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
         Index           =   2
         Left            =   13455
         TabIndex        =   53
         Text            =   "12345678901234567890"
         Top             =   1125
         Width           =   4920
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
         Index           =   1
         Left            =   13455
         TabIndex        =   52
         Text            =   "12345678901234567890"
         Top             =   720
         Width           =   4920
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
         Index           =   0
         Left            =   13455
         TabIndex        =   51
         Text            =   "12345678901234567890"
         Top             =   315
         Width           =   4920
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
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
         Left            =   9240
         MaxLength       =   8
         TabIndex        =   7
         Tag             =   "Campo 2|N|S|||horasmasivo|codcampo|00000000||"
         Text            =   "12345678"
         Top             =   1125
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
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
         Left            =   11490
         MaxLength       =   7
         TabIndex        =   8
         Tag             =   "Horas Dia|N|S|||horasmasivo|horasdia|##0.00||"
         Text            =   "1234567"
         Top             =   1125
         Width           =   1000
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
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
         Left            =   11490
         MaxLength       =   7
         TabIndex        =   10
         Tag             =   "Horas Dia|N|S|||horasmasivo|horasdia|##0.00||"
         Text            =   "1234567"
         Top             =   1530
         Width           =   1000
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
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
         Left            =   9240
         MaxLength       =   8
         TabIndex        =   9
         Tag             =   "Campo 2|N|S|||horasmasivo|codcampo|00000000||"
         Text            =   "12345678"
         Top             =   1530
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
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
         Left            =   9240
         MaxLength       =   8
         TabIndex        =   5
         Tag             =   "Campo 2|N|S|||horasmasivo|codcampo|00000000||"
         Text            =   "12345678"
         Top             =   720
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
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
         Left            =   11490
         MaxLength       =   7
         TabIndex        =   6
         Tag             =   "Horas Dia|N|S|||horasmasivo|horasdia|##0.00||"
         Text            =   "1234567"
         Top             =   720
         Width           =   1000
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
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
         Left            =   11490
         MaxLength       =   7
         TabIndex        =   4
         Tag             =   "Horas Dia|N|S|||horasmasivo|horasdia|##0.00||"
         Text            =   "1234567"
         Top             =   315
         Width           =   1000
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
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
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Entrada|F|N|||horasmasivo|fecha|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   1350
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
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
         Left            =   9240
         MaxLength       =   8
         TabIndex        =   3
         Tag             =   "Campo 2|N|S|||horasmasivo|codcampo|00000000||"
         Text            =   "12345678"
         Top             =   315
         Width           =   1275
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
         Left            =   2490
         TabIndex        =   28
         Text            =   "12345678901234567890"
         Top             =   330
         Width           =   5190
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
         Left            =   1605
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "Capataz|N|N|0|999999|horasmasivo|codcapat|000000|S|"
         Text            =   "123456"
         Top             =   330
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
         Index           =   0
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Complemento|N|S|||horasmasivo|compleme|###,##0.00||"
         Text            =   "1234567"
         Top             =   870
         Width           =   1320
      End
      Begin VB.Label Label6 
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
         Height          =   255
         Index           =   5
         Left            =   12645
         TabIndex        =   59
         Top             =   1575
         Width           =   690
      End
      Begin VB.Label Label6 
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
         Height          =   255
         Index           =   4
         Left            =   12645
         TabIndex        =   58
         Top             =   1170
         Width           =   690
      End
      Begin VB.Label Label6 
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
         Height          =   255
         Index           =   3
         Left            =   12645
         TabIndex        =   57
         Top             =   765
         Width           =   690
      End
      Begin VB.Label Label6 
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
         Height          =   255
         Index           =   1
         Left            =   12645
         TabIndex        =   56
         Top             =   360
         Width           =   690
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   8955
         ToolTipText     =   "Buscar Campo"
         Top             =   1575
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   8955
         ToolTipText     =   "Buscar Campo"
         Top             =   1170
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   8955
         ToolTipText     =   "Buscar Campo"
         Top             =   765
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   8955
         ToolTipText     =   "Buscar Campo"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Campo 3"
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
         Left            =   8025
         TabIndex        =   44
         Top             =   1170
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Horas 3"
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
         Left            =   10635
         TabIndex        =   43
         Top             =   1155
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "Horas 4"
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
         Left            =   10635
         TabIndex        =   42
         Top             =   1575
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Campo 4"
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
         Left            =   8025
         TabIndex        =   41
         Top             =   1575
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Campo 2"
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
         Left            =   8025
         TabIndex        =   39
         Top             =   750
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Horas 2"
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
         Left            =   10635
         TabIndex        =   38
         Top             =   765
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Horas 1"
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
         Left            =   10635
         TabIndex        =   37
         Top             =   345
         Width           =   1095
      End
      Begin VB.Label Label14 
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
         Left            =   180
         TabIndex        =   32
         Top             =   1395
         Width           =   660
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1290
         Picture         =   "frmManHorasCreacion.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   1410
         Width           =   240
      End
      Begin VB.Label Label6 
         Caption         =   "Campo 1"
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
         Left            =   8025
         TabIndex        =   31
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label10 
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
         Height          =   255
         Left            =   180
         TabIndex        =   29
         Top             =   375
         Width           =   960
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1335
         ToolTipText     =   "Buscar Capataz"
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Complemento"
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
         Left            =   180
         TabIndex        =   14
         Top             =   900
         Width           =   1320
      End
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Index           =   0
      Left            =   14220
      TabIndex        =   30
      Text            =   "Text3"
      Top             =   9810
      Width           =   1335
   End
   Begin VB.Frame FrameAux0 
      Caption         =   "Horas Generadas"
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
      Height          =   7320
      Left            =   135
      TabIndex        =   16
      Top             =   3015
      Width           =   18740
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
         Height          =   290
         Index           =   4
         Left            =   4500
         MaxLength       =   40
         TabIndex        =   55
         Text            =   "precio"
         Top             =   2790
         Visible         =   0   'False
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
         Height          =   290
         Index           =   6
         Left            =   8010
         MaxLength       =   7
         TabIndex        =   50
         Tag             =   "Codusuario|N|S|||horasmasivo|codusu|###,##0.00||"
         Text            =   "usu"
         Top             =   2745
         Visible         =   0   'False
         Width           =   570
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
         Left            =   2115
         MaskColor       =   &H00000000&
         TabIndex        =   49
         ToolTipText     =   "Buscar trabajador"
         Top             =   2745
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
         Height          =   290
         Index           =   2
         Left            =   3780
         MaxLength       =   40
         TabIndex        =   48
         Text            =   "nomsalar"
         Top             =   2745
         Visible         =   0   'False
         Width           =   600
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
         Height          =   290
         Index           =   1
         Left            =   2295
         MaxLength       =   40
         TabIndex        =   47
         Text            =   "nomtraba"
         Top             =   2745
         Visible         =   0   'False
         Width           =   600
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
         Height          =   290
         Index           =   0
         Left            =   1035
         MaxLength       =   40
         TabIndex        =   33
         Text            =   "nomvarie"
         Top             =   2745
         Visible         =   0   'False
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
         Height          =   290
         Index           =   4
         Left            =   5985
         MaxLength       =   7
         TabIndex        =   20
         Tag             =   "Horas Dia|N|S|||horasmasivo|horasdia|##0.00||"
         Text            =   "Horas"
         Top             =   2745
         Visible         =   0   'False
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
         Height          =   290
         Index           =   3
         Left            =   7245
         MaxLength       =   7
         TabIndex        =   22
         Tag             =   "Complementos|N|S|||horasmasivo|compleme|###,##0.00||"
         Text            =   "complem"
         Top             =   2745
         Visible         =   0   'False
         Width           =   570
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
         Height          =   290
         Index           =   5
         Left            =   6615
         MaxLength       =   7
         TabIndex        =   21
         Tag             =   "Importe|N|N|||horasmasivo|importe|###,##0.00||"
         Text            =   "importe"
         Top             =   2745
         Visible         =   0   'False
         Width           =   600
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
         Height          =   290
         Index           =   2
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   19
         Tag             =   "Trabajador|N|N|||horasmasivo|codtraba|000000|S|"
         Text            =   "Trabaj"
         Top             =   2745
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtAux2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
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
         Height          =   315
         Index           =   3
         Left            =   3060
         MaxLength       =   10
         TabIndex        =   25
         Text            =   "categ"
         Top             =   2745
         Visible         =   0   'False
         Width           =   600
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
         Height          =   290
         Index           =   1
         Left            =   585
         MaxLength       =   6
         TabIndex        =   18
         Tag             =   "Variedad|N|N|||horasmasivo|codvarie|000000|S|"
         Text            =   "Var"
         Top             =   2745
         Visible         =   0   'False
         Width           =   405
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
         Height          =   290
         Index           =   0
         Left            =   135
         MaxLength       =   8
         TabIndex        =   17
         Tag             =   "Codcampo|N|N|||horasmasivo|codcampo|00000000|S|"
         Text            =   "campo"
         Top             =   2745
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   0
         Left            =   135
         TabIndex        =   26
         Top             =   225
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
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
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc AdoAux 
         Height          =   375
         Index           =   0
         Left            =   3720
         Top             =   225
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
         Bindings        =   "frmManHorasCreacion.frx":0097
         Height          =   6000
         Index           =   0
         Left            =   135
         TabIndex        =   27
         Top             =   630
         Width           =   18295
         _ExtentX        =   32279
         _ExtentY        =   10583
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
      Begin VB.Label Label11 
         Caption         =   "TOTALES: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   13005
         TabIndex        =   36
         Top             =   6795
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   10395
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
         TabIndex        =   12
         Top             =   180
         Width           =   2655
      End
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
      Left            =   17805
      TabIndex        =   24
      Top             =   10485
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
      Left            =   16545
      TabIndex        =   23
      Top             =   10485
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   1980
      Top             =   6120
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Left            =   17775
      TabIndex        =   15
      Top             =   10485
      Visible         =   0   'False
      Width           =   1065
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   18360
      TabIndex        =   40
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
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver Todos"
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
      Begin VB.Menu mnGastos 
         Caption         =   "&Cálculo Gastos"
         Enabled         =   0   'False
         Shortcut        =   ^G
         Visible         =   0   'False
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnExporImpor 
         Caption         =   "Exportar/Importar"
         Enabled         =   0   'False
         Visible         =   0   'False
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
Attribute VB_Name = "frmManHorasCreacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MANOLO                   -+-+
' +-+- Menú: CLIENTES                  -+-+
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

' +-+-+-+- DISSENY +-+-+-+-
' 1. Posar tots els controls al formulari
' 2. Posar els index correlativament
' 3. Si n'hi han botons de buscar repasar el ToolTipText
' 4. Alliniar els camps numérics a la dreta i el resto a l'esquerra
' 5. Posar els TAGs
' (si es INTEGER: si PK => mínim 1; si no PK => mínim 0; màxim => 99; format => 00)
' (si es DECIMAL; mínim => 0; màxim => 99.99; format => #,###,###,##0.00)
' (si es DATE; format => dd/mm/yyyy)
' 6. Posar els MAXLENGTHs
' 7. Posar els TABINDEXs

Option Explicit


Private Const IdPrograma = 4014

'Dim T1 As Single

Public DatosADevolverBusqueda As String    'Tindrà el nº de text que vol que torne, empipat
Public Event DatoSeleccionado(CadenaSeleccion As String)
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes ' campos del socio
Attribute frmMens.VB_VarHelpID = -1
'Private WithEvents frmExp As frmExpImpExcel ' Exportacion o importacion a pagina excel

'Private WithEvents frmArt As frmManArtic 'articulos
Private WithEvents frmVar As frmComVar 'variedades de comercial
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmCap As frmManCapataz 'capataces
Attribute frmCap.VB_VarHelpID = -1
Private WithEvents frmTra As frmManTraba 'trabajador
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmTar As frmManTarTra 'tarifas de transporte
Attribute frmTar.VB_VarHelpID = -1
Private WithEvents frmInc As frmManInciden 'incidencias
Attribute frmInc.VB_VarHelpID = -1
Private WithEvents frmCal As frmManCalidades 'calidades
Attribute frmCal.VB_VarHelpID = -1
Private WithEvents frmCam As frmManCampos 'campos
Attribute frmCam.VB_VarHelpID = -1
Private WithEvents frmCat As frmManCategorias 'categorias
Attribute frmCat.VB_VarHelpID = -1
'
'*****************************************************
Private Modo As Byte
'*************** MODOS ********************
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'   5.-  Manteniment Llinies

'+-+-Variables comuns a tots els formularis+-+-+

Dim ModoLineas As Byte
'1.- Afegir,  2.- Modificar,  3.- Borrar,  0.-Passar control a Llínies

Dim NumTabMto As Integer 'Indica quin nº de Tab està en modo Mantenimient
Dim TituloLinea As String 'Descripció de la llínia que està en Mantenimient
Dim PrimeraVez As Boolean

Private CadenaConsulta As String 'SQL de la taula principal del formulari
Private Ordenacion As String
Private NombreTabla As String  'Nom de la taula
Private NomTablaLineas As String 'Nom de la Taula de llínies del Mantenimient en que estem

Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

'Private VieneDeBuscar As Boolean
'Per a quan torna 2 poblacions en el mateix codi Postal. Si ve de pulsar prismatic
'de búsqueda posar el valor de població seleccionada i no tornar a recuperar de la Base de Datos

Dim Gastos As Boolean

Dim CodTipoMov As String
Dim NotaExistente As Boolean

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
'Dim CadAncho() As Boolean  'array, per a quan cridem al form de llínies
Dim Indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim CadB As String

Dim VarieAnt As String


Dim cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Dim cadParam As String  'Cadena con los parametros para Crystal Report
Dim numParam As Byte  'Numero de parametros que se pasan a Crystal Report
Dim cadselect As String  'Cadena para comprobar si hay datos antes de abrir Informe
Dim cadTitulo As String  'Titulo para la ventana frmImprimir
Dim cadNombreRPT As String  'Nombre del informe
Dim cadSelect1 As String 'Cadena para comprobar si hay datos antes de abrir Informe


Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos


Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 4 'Trabajador
            AbrirFrmTrabajador 7
        Case 5 'categoria
            AbrirFrmCategoria 11
            
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

Private Sub AbrirFrmTrabajador(Indice As Integer)
    indCodigo = 2
    Set frmTra = New frmManTraba
    frmTra.DatosADevolverBusqueda = "0|2|"
    frmTra.Show vbModal
    Set frmTra = Nothing
    
    PonerFoco txtAux(indCodigo)

End Sub

Private Sub AbrirFrmCategoria(Indice As Integer)
    indCodigo = 3
    Set frmCat = New frmManCategorias
    frmCat.DatosADevolverBusqueda = "0|1|"
    frmCat.Show vbModal
    Set frmCat = Nothing
    
    PonerFoco txtAux(indCodigo)
End Sub

Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
            
        Case 3 'INSERTAR
            If DatosOk Then
                InsertarRegistros
                CargaGrid 0, True
                CalcularTotales
                PonerModo 5
            Else
                ModoLineas = 0
            End If
            
        
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                    InsertarLinea
                Case 2 'modificar llínies
                    If Not ModificarLinea Then
                        PonerFoco txtAux(5)
                    End If
            End Select
            'nuevo calculamos los totales de lineas
            CalcularTotales
                    

        ' **************************
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
    If Modo = 5 And Not (ModoLineas = 1 Or ModoLineas = 2) Then
        PonerContRegIndicador lblIndicador, Adoaux(0), CadB
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
'    If PrimeraVez Then PrimeraVez = False
    If PrimeraVez Then
        BotonPedirDatos False
        PrimeraVez = False
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Modo = 4 Then TerminaBloquear
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim i As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
 
    ' ICONETS DE LA BARRA
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Pedir Datos
        .Buttons(2).Image = 15   'Generar FActura
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 12
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
    
    'cargar IMAGES .Image =de busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han llínies *******
    DataGridAux(0).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "horasmasivo"
    Ordenacion = " ORDER BY codcampo, codtraba"
    
    
    conn.Execute "delete from horasmasivo where codusu = " & vUsu.Codigo
    
    Data1.ConnectionString = conn
    '***** cambiar el nombre de la PK de la cabecera *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where codtraba is null"
    Data1.Refresh
       
    CargaGrid 0, False
       
    ModoLineas = 0
    
End Sub

Private Sub LimpiarCampos()
Dim SQL As String
    
    On Error Resume Next
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
    SQL = "delete from horasmasivo where codusu =" & DBSet(vUsu.Codigo, "N")
    conn.Execute SQL
    
    ' *** si n'hi han combos a la capçalera ***
    ' *****************************************

    Text3(0).Text = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub LimpiarCamposLin(FrameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, FrameAux  'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO s'habiliten, o no, els diversos camps del
'   formulari en funció del modo en que anem a treballar
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim B As Boolean
Dim i As Integer

    Modo = Kmodo
    
    If Modo = 5 Then
        If (ModoLineas = 1 Or ModoLineas = 2) Then
            PonerIndicador lblIndicador, Modo, ModoLineas
        Else
            PonerContRegIndicador lblIndicador, Me.Adoaux(0), CadB
        End If
    End If
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    B = (Modo = 3 Or (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2)))
        
    cmdAceptar.visible = B
    cmdAceptar.Enabled = B
    cmdCancelar.visible = B
    cmdCancelar.Enabled = B
    

    'b = (Modo = 3 Or Modo = 4 Or Modo = 1 Or Modo = 5)
    B = (Modo = 3 Or Modo = 4 Or Modo = 1) '06/09/2005, lleve el modo 5 per a que no es puga modificar la capçalera mentre treballe en les llínies
    
    ' no pone -1 pq no tenemos el text1(9)
    For i = 0 To Text1.Count
        BloquearTxt Text1(i), (Modo <> 3)
    Next i
    
    B = (Modo = 5)
    For i = 0 To 6
        BloquearTxt txtAux(i), Not B
        txtAux(i).visible = False
    Next i
    
    For i = 4 To 4
        BloquearBtn btnBuscar(i), Not B
        btnBuscar(i).visible = False
    Next i
        
    For i = 0 To 4
        txtAux2(i).visible = False
    Next i
    
    Me.FrameIntro.Enabled = (Modo = 3)
    Me.FrameAux0.Enabled = (Modo = 5)
       
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
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los TEXT1
    PonerLongCamposGnral Me, Modo, 1
End Sub

Private Sub PonerOpcionesMenu()
Dim J As Byte

    PonerOpcionesMenuGeneral Me
    
'    J = Val(Me.mnPedirDatos.HelpContextID)
'    If J < vUsu.Nivel Then Me.mnPedirDatos.Enabled = False
'
'    J = Val(Me.mnGenerarFac.HelpContextID)
'    If J < vUsu.Nivel Then Me.mnGenerarFac.Enabled = False
End Sub

Private Sub PonerModoOpcionesMenu(Modo)
'Actives unes Opcions de Menú i Toolbar según el modo en que estem
Dim B As Boolean, bAux As Boolean
Dim i As Byte
    
    'Barra de CAPÇALERA
    '------------------------------------------
    'b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    B = (Modo = 2 Or Modo = 0)
    'Buscar
    Toolbar1.Buttons(5).Enabled = B
    Me.mnBuscar.Enabled = B
    'Vore Tots
    Toolbar1.Buttons(6).Enabled = B
    Me.mnVerTodos.Enabled = B
    
    'Insertar
    Toolbar1.Buttons(1).Enabled = B And Not DeConsulta
    Me.mnNuevo.Enabled = B And Not DeConsulta
    
    B = (Modo = 2 And Data1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    Me.mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(3).Enabled = B
    Me.mnEliminar.Enabled = B
    
    'Imprimir
    'Toolbar1.Buttons(12).Enabled = (b Or Modo = 0)
    Toolbar1.Buttons(8).Enabled = True And Not DeConsulta
    Me.mnImprimir.Enabled = True And Not DeConsulta
       
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
'++monica: si insertamos lo he quitado
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    B = (Modo = 2) And Not DeConsulta
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = B
        If B Then bAux = (B And Me.Adoaux(i).Recordset.RecordCount > 0)
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i
    
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botons de Desplaçament; per a desplaçar-se pels registres de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index, 350
    PonerCampos
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
Dim KilosTot As Long

    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
               
        Case 0 'HORAS MASIVO
            SQL = "select horasmasivo.codusu, horasmasivo.codcampo, horasmasivo.codvarie, variedades.nomvarie, horasmasivo.codtraba, straba.nomtraba, "
            SQL = SQL & " straba.codcateg, salarios.nomcateg, salarios.impsalar,"
            SQL = SQL & " horasmasivo.horasdia, horasmasivo.importe, horasmasivo.compleme "
            SQL = SQL & " from ((horasmasivo inner join variedades on horasmasivo.codvarie = variedades.codvarie)  "
            SQL = SQL & " inner join straba on horasmasivo.codtraba = straba.codtraba) "
            SQL = SQL & " inner join salarios on straba.codcateg = salarios.codcateg "
            If enlaza Then
                SQL = SQL & ObtenerWhereCab(True)
            Else
                SQL = SQL & " WHERE horasmasivo.codtraba is null"
            End If
            
            SQL = SQL & " order by 6,2"
    End Select
    
    MontaSQLCarga = SQL
End Function

Private Sub frmC_Selec(vFecha As Date)
Dim Indice As Byte
    Indice = CByte(Me.imgFec(0).Tag)
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCam_DatoSeleccionado(CadenaSeleccion As String)
'Campos
    Text1(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codcampo
End Sub

Private Sub frmCap_DatoSeleccionado(CadenaSeleccion As String)
'Capataces
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codcapat
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmCat_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        txtAux(3).Text = RecuperaValor(CadenaSeleccion, 1)
        txtAux2(2).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
    
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo trabajador
    txtAux2(1).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre trabajador
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo trabajador
    txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre trabajador
End Sub

Private Sub imgFec_Click(Index As Integer)
   
   Screen.MousePointer = vbHourglass
   
   Dim esq As Long
   Dim dalt As Long
   Dim menu As Long
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
    
   menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

   frmC.Left = esq + imgFec(Index).Parent.Left + 30
   frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40

   
   frmC.NovaData = Now
   Select Case Index
        Case 0
            Indice = 1
   End Select
   
   Me.imgFec(0).Tag = Indice
   
   PonerFormatoFecha Text1(Indice)
   If Text1(Indice).Text <> "" Then frmC.NovaData = CDate(Text1(Indice).Text)

   Screen.MousePointer = vbDefault
   frmC.Show vbModal
   Set frmC = Nothing
   PonerFoco Text1(Indice)

End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        Indice = 20
        frmZ.pTitulo = "Observaciones de la Clasificación"
        frmZ.pValor = Text1(Indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco Text1(Indice)
    End If
End Sub




Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Pedir datos
            mnPedirDatos_Click
        Case 2 'Generar Factura
            mnGenerarEntHoras_Click
    End Select
End Sub

Private Sub mnPedirDatos_Click()
    BotonPedirDatos True
End Sub

Private Sub mnGenerarEntHoras_Click()
    BotonGenerarHoras
End Sub


Private Sub BotonPedirDatos(Preguntar As Boolean)
Dim Nombre As String
Dim i As Integer
Dim SQL As String

    TerminaBloquear

    SQL = "delete from horasmasivo where codusu = " & vUsu.Codigo
    conn.Execute SQL

    'Vaciamos todos los Text
    If Text1(3).Text <> "" And Preguntar Then
        If MsgBox("¿ Desea limpiar datos ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            LimpiarCampos
            'fecha
            Text1(1).Text = Format(Now, "dd/mm/yyyy")
        End If
    Else
        LimpiarCampos
        'fecha
        Text1(1).Text = Format(Now, "dd/mm/yyyy")
    End If
    
    CargaGrid 0, False
    
    PonerModo 3
    
    'desbloquear los registros de la rhisfruta
    TerminaBloquear
    
    'si vamos
    'desBloqueo Manual de las tablas
'    DesBloqueoManual ("scaalp")
    
    Me.cmdAceptar.visible = True
    Me.cmdAceptar.Enabled = True
    Me.cmdCancelar.visible = True
    Me.cmdCancelar.Enabled = True
    
    
    PonerFoco Text1(3)
End Sub


Private Sub BotonGenerarHoras()
Dim Nombre As String
Dim i As Integer
Dim SQL As String

    On Error GoTo eBotonGenerarHoras
    
    If Modo <> 5 Then Exit Sub
    
    TerminaBloquear

    If MsgBox("¿Son los datos correctos para insertar en horas?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If

    conn.BeginTrans

    SQL = "insert into horas (codtraba,fechahora,horasdia,compleme,codalmac,codvarie,importe,codcapat,codcateg,codcampo) "
    SQL = SQL & " select codtraba, " & DBSet(Text1(1), "F") & ", horasdia, compleme, 1, codvarie, importe, " & DBSet(Text1(3).Text, "N") & ","
    SQL = SQL & " codcateg, codcampo from horasmasivo where codusu = " & DBSet(vUsu.Codigo, "N")
    
    conn.Execute SQL
    
    SQL = "delete from horasmasivo where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute SQL
    
    limpiar Me
    
    PonerModo 3
    CargaGrid 0, False
    
    conn.CommitTrans
    Exit Sub
    
eBotonGenerarHoras:
    MuestraError Err.Number, "Generando Horas", Err.Description
    conn.RollbackTrans
End Sub








Private Sub cmdRegresar_Click()
Dim cad As String
Dim Aux As String
Dim i As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
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
            cad = cad & Text1(J).Text & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
        'Data1.Recordset.MoveLast
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub

EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim i As Integer
Dim CodPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    
    
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la capçalera
    
    ' *** si n'hi han llínies en datagrids ***
    'For i = 0 To DataGridAux.Count - 1
    For i = 0 To 0
        CargaGrid i, True
        If Not Adoaux(i).Recordset.EOF Then _
            PonerCamposForma2 Me, Adoaux(i), 2, "FrameAux" & i
    Next i

    
    ' ************* configurar els camps de les descripcions de la capçalera *************
    Text2(3).Text = PonerNombreDeCod(Text1(3), "variedades", "nomvarie")
    Text2(4).Text = PonerNombreDeCod(Text1(4), "rsocios", "nomsocio")
    ' ********************************************************************************
    
    CalcularTotales
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu
End Sub

Private Sub cmdCancelar_Click()
Dim i As Integer
Dim V

    Select Case Modo
        Case 3 'Búsqueda, Insertar
                LimpiarCampos
                If Data1.Recordset.EOF Then
                    PonerModo 0
                Else
                    PonerModo 2
                    PonerCampos
                End If
                ' *** foco al primer camp visible de la capçalera ***
                PonerFoco Text1(0)

        
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1
                    ModoLineas = 0
                    Me.DataGridAux(0).AllowAddNew = False
                    PonerModo 5
                    CargaGrid 0, True
            
                Case 2 'modificar llínies
                    ModoLineas = 0
                    
                    ' *** si n'hi han tabs ***
'                    SituarTab (NumTabMto + 1)
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                    PonerModo 5
                    If Not Adoaux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de llínies ***
                        V = Adoaux(NumTabMto).Recordset.Fields(4) 'el 2 es el nº de llinia
                        Adoaux(0).Recordset.Find (Adoaux(0).Recordset.Fields(4).Name & " =" & V)
                        ' ***************************************************************
                    End If
            End Select
            
    End Select
End Sub

Private Function DatosOk() As Boolean
Dim B As Boolean
Dim Nregs As Integer
Dim SQL As String

    On Error GoTo EDatosOK

    DatosOk = False
    
    B = CompForm2(Me, 1)
    If Not B Then Exit Function
    
    If B Then
        If Text1(3).Text = "" Then
            MsgBox "Debe introducir un valor en capataz. Revise", vbExclamation
            PonerFoco Text1(3)
            B = False
        End If
    End If
    
    If B Then
        If Text1(0).Text = "" Then
            MsgBox "Debe introducir un complemento para el capataz.", vbExclamation
            PonerFoco Text1(0)
            B = False
        End If
    End If
    
    If B Then
        If Text1(1).Text = "" Then
            MsgBox "Debe introducir un valor en fecha de entrada de horas. Revise", vbExclamation
            PonerFoco Text1(1)
            B = False
        End If
    End If
    
    If B Then
        If Text1(2).Text = "" Then
            MsgBox "Debe introducir un valor en campo1. Revise", vbExclamation
            PonerFoco Text1(2)
            B = False
        Else
            'el campo debe de existir
            SQL = DevuelveDesdeBDNew(cAgro, "rcampos", "codcampo", "codcampo", Text1(2).Text, "N")
            If SQL = "" Then
                MsgBox "Codigo de campo 1 no existe. Reintroduzca.", vbExclamation
                PonerFoco Text1(2)
                B = False
            End If
        End If
    End If
    
    If B Then
        If Text1(5).Text = "" Then
            MsgBox "Debe introducir un valor en horas1. Revise", vbExclamation
            PonerFoco Text1(2)
            B = False
        End If
    End If
    
    If B Then
        'el campo 2 debe de existir
        If Text1(8).Text <> "" Then
            SQL = DevuelveDesdeBDNew(cAgro, "rcampos", "codcampo", "codcampo", Text1(8).Text, "N")
            If SQL = "" Then
                MsgBox "Codigo de campo 2 no existe. Reintroduzca.", vbExclamation
                PonerFoco Text1(8)
                B = False
            Else
                If Text1(6).Text = "" Then
                    MsgBox "Debe introducir el nro de horas para el campo 2", vbExclamation
                    PonerFoco Text1(6)
                    B = False
                End If
            End If
        End If
    End If
    If B Then
        If Text1(11).Text <> "" Then
            'el campo 3 debe de existir
            SQL = DevuelveDesdeBDNew(cAgro, "rcampos", "codcampo", "codcampo", Text1(11).Text, "N")
            If SQL = "" Then
                MsgBox "Codigo de campo 3 no existe. Reintroduzca.", vbExclamation
                PonerFoco Text1(11)
                B = False
            Else
                If Text1(10).Text = "" Then
                    MsgBox "Debe introducir el nro de horas para el campo 3", vbExclamation
                    PonerFoco Text1(10)
                    B = False
                End If
            End If
        End If
    End If
    If B Then
        If Text1(7).Text <> "" Then
            'el campo 4 debe de existir
            SQL = DevuelveDesdeBDNew(cAgro, "rcampos", "codcampo", "codcampo", Text1(7).Text, "N")
            If SQL = "" Then
                MsgBox "Codigo de campo 4 no existe. Reintroduzca.", vbExclamation
                PonerFoco Text1(7)
                B = False
            Else
                If Text1(4).Text = "" Then
                    MsgBox "Debe introducir el nro de horas para el campo 4", vbExclamation
                    PonerFoco Text1(4)
                    B = False
                End If
            End If
        End If
    End If
    
    ' ************************************************************************************
    DatosOk = B
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub PosicionarData()
Dim cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    cad = "(nroclasif=" & DBSet(Text1(0).Text, "N")
    cad = cad & " and codvarie = " & DBSet(Text1(3).Text, "N")
    cad = cad & " and codsocio = " & DBSet(Text1(4).Text, "N")
    cad = cad & " and fechacla = " & DBSet(Text1(1).Text, "F")
    cad = cad & " and codcampo = " & DBSet(Text1(2).Text, "N")
    cad = cad & " and ordinal = " & DBSet(Text1(7).Text, "N") & ")"
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    If SituarDataMULTI(Data1, cad, Indicador) Then
    'If SituarData(Data1, cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub

Private Function Eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar

    conn.BeginTrans
    ' ***** canviar el nom de la PK de la capçalera, repasar codEmpre *******
    vWhere = " WHERE nroclasif=" & Data1.Recordset!nroclasif
    vWhere = vWhere & " and codvarie = " & Data1.Recordset!Codvarie
    vWhere = vWhere & " and codsocio = " & Data1.Recordset!Codsocio
    vWhere = vWhere & " and fechacla = " & DBSet(Data1.Recordset!fechacla, "F")
    vWhere = vWhere & " and codcampo = " & Data1.Recordset!codCampo
    vWhere = vWhere & " and ordinal  = " & Data1.Recordset!Ordinal
    
    
    ' ***** elimina les llínies ****
    conn.Execute "DELETE FROM rcontrol_plagas " & vWhere
        
    'Eliminar la CAPÇALERA
    conn.Execute "Delete from " & NombreTabla & vWhere
       
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim SQL As String
Dim Nregs As Integer
Dim Rs As ADODB.Recordset
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    
    ' ***************** configurar els LostFocus dels camps de la capçalera *****************
    Select Case Index
        Case 3 'Capataz
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rcapataz", "nomcapat")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Capataz: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCap = New frmManCapataz
                        frmCap.DatosADevolverBusqueda = "0|1|"
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmCap.Show vbModal
                        Set frmCap = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Adoaux(0), 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                Else
                    ' si tiene complemento me lo traigo
                    SQL = "select salarios.pluscapataz from (straba inner join rcapataz on straba.codtraba = rcapataz.codtraba)  "
                    SQL = SQL & " inner join salarios on straba.codcateg = salarios.codcateg "
                    SQL = SQL & " where codcapat = " & DBSet(Text1(Index).Text, "N")
                    
                    Set Rs = New ADODB.Recordset
                    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    If Not Rs.EOF Then
                        Text1(0).Text = Format(DBLet(Rs.Fields(0).Value, "N"), "###,##0.00")
                    End If
                    Set Rs = Nothing
                End If
            End If
        
        Case 4, 5, 6, 10 'horas
            PonerFormatoDecimal Text1(Index), 3
            
        Case 1
            '[Monica]28/08/2013: comprobamos que la fecha esté en la campaña
            PonerFormatoFecha Text1(Index)
        
        Case 2, 7, 8, 11 'campo
            PonerFormatoEntero Text1(Index)
            Select Case Index
                Case 2
                    Text2(0).Text = PartidaCampo(Text1(Index))
                Case 7
                    Text2(4).Text = PartidaCampo(Text1(Index))
                Case 8
                    Text2(1).Text = PartidaCampo(Text1(Index))
                Case 11
                    Text2(2).Text = PartidaCampo(Text1(Index))
            End Select
                    

        Case 0 'Complemento
            PonerFormatoDecimal Text1(Index), 3
            
        Case 9 'Categoria
            PonerFormatoEntero Text1(Index)
            
    End Select
    ' ***************************************************************************
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 3: KEYBusqueda KeyAscii, 0 'variedad
                Case 4: KEYBusqueda KeyAscii, 1 'socio
                Case 1: KEYFecha KeyAscii, 0 'fecha
            End Select
        End If
    Else
'        If Index <> 3 Or (Index = 3 And Text1(3).Text = "") Then KEYpress KeyAscii
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFec_Click (Indice)
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
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
    imgBuscar_Click (Indice)
End Sub

Private Sub KEYBusquedaBtn(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    btnBuscar_Click (Indice)
End Sub

'************* LLINIES: ****************************
Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
'-- pon el bloqueo aqui
    'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
    Select Case Button.Index
        Case 1
            BotonAnyadirLinea Index
        Case 2
            BotonModificarLinea Index
        Case 3
            BotonEliminarLinea Index
        Case Else
    End Select
    'End If
End Sub

Private Sub BotonEliminarLinea(Index As Integer)
Dim SQL As String

    On Error GoTo EEliminarLinea

    If Me.Adoaux(0).Recordset.EOF Then Exit Sub

    SQL = "delete from horasmasivo where codusu = " & DBSet(vUsu.Codigo, "N")
    SQL = SQL & " and codcampo = " & Me.Adoaux(0).Recordset!codCampo
    SQL = SQL & " and codvarie = " & Me.Adoaux(0).Recordset!Codvarie
    SQL = SQL & " and codtraba = " & Me.Adoaux(0).Recordset!CodTraba
    
    conn.Execute SQL
    
    CargaGrid 0, True
    Exit Sub

EEliminarLinea:
    MuestraError Err.Number, "Eliminar Linea", Err.Description
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
    PonerModo 5, Index
    ' *** bloqueje la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
  
    Select Case Index
        Case 0 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
            If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
                i = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
                DataGridAux(Index).Scroll 0, i
                DataGridAux(Index).Refresh
            End If
              
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 240
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
    End Select
    
    Select Case Index
        ' *** valor per defecte al modificar dels camps del grid ***
        Case 0 ' muestra
            txtAux(0).Text = DataGridAux(Index).Columns(1).Text 'codcampo
            txtAux(1).Text = DataGridAux(Index).Columns(2).Text
            txtAux2(0).Text = DataGridAux(Index).Columns(3).Text
            txtAux(2).Text = DataGridAux(Index).Columns(4).Text
            txtAux2(1).Text = DataGridAux(Index).Columns(5).Text
            txtAux2(3).Text = DataGridAux(Index).Columns(6).Text
            txtAux2(2).Text = DataGridAux(Index).Columns(7).Text
            txtAux2(4).Text = DataGridAux(Index).Columns(8).Text
            txtAux(4).Text = DataGridAux(Index).Columns(9).Text
            txtAux(5).Text = DataGridAux(Index).Columns(10).Text
            txtAux(3).Text = DataGridAux(Index).Columns(11).Text
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'muestras
            PonerFoco txtAux(4)
    End Select
    ' ***************************************************************************************
End Sub

Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim B As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    DeseleccionaGrid DataGridAux(Index)
       
    B = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 0 'muestras
            For jj = 3 To 5
                txtAux(jj).visible = B
                txtAux(jj).Top = alto
            Next jj
            
            If xModo = 1 Then
                For jj = 0 To 2
                    txtAux(jj).visible = B
                    txtAux(jj).Top = alto
                Next jj
            End If
            'insertar
            If xModo = 1 Then
                For jj = 4 To 4
                    btnBuscar(jj).visible = B
                    btnBuscar(jj).Top = alto
                Next jj
                For jj = 0 To 4
                    txtAux2(jj).visible = B
                    txtAux2(jj).Top = alto
                Next jj
            End If
            'modificar
            If xModo = 2 Then
                txtAux2(2).visible = B
                txtAux2(2).Top = alto
                txtAux2(3).visible = B
                txtAux2(3).Top = alto
                txtAux2(4).visible = B
                txtAux2(4).Top = alto
                
            End If
    End Select
End Sub


Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub



Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim PrecioHora As Currency
Dim SQL As String
Dim Rs As ADODB.Recordset

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    

    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 0 ' campo
            If PonerFormatoEntero(txtAux(Index)) Then
                If txtAux(0).Text <> "" Then
                    SQL = DevuelveDesdeBDNew(cAgro, "rcampos", "codvarie", "codcampo", txtAux(0), "N")
                    If SQL = "" Then
                        MsgBox "No existe el campo. Reintroduzca.", vbExclamation
                        PonerFoco txtAux(0)
                    Else
                        txtAux(1).Text = SQL
                        txtAux2(0).Text = PonerNombreDeCod(txtAux(1), "variedades", "nomvarie", "codvarie", "N")
                        If txtAux2(0).Text = "" Then
                            MsgBox "Variedad no existe. Revise.", vbExclamation
                            PonerFoco txtAux(1)
                        End If
                    End If
                End If
            End If
            
        Case 1 ' Variedad
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(0) = PonerNombreDeCod(txtAux(Index), "variedades", "nomvarie", "codvarie", "N")
                If txtAux2(0).Text = "" Then
                    MsgBox "Variedad no existe. Revise.", vbExclamation
                    PonerFoco txtAux(1)
                End If
            End If
            
        Case 2 'Trabajador
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(1).Text = PonerNombreDeCod(txtAux(Index), "straba", "nomtraba")
                If txtAux2(1).Text = "" Then
                    cadMen = "No existe el Trabajador: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmTra = New frmManTraba
                        frmTra.DatosADevolverBusqueda = "0|1|"
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmTra.Show vbModal
                        Set frmCap = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Adoaux(0), 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                Else
                    ' sacamos los datos del salario
                    SQL = "select salarios.codcateg, salarios.nomcateg, salarios.impsalar from salarios inner join straba on straba.codcateg = salarios.codcateg where codtraba = " & DBSet(txtAux(2).Text, "N")
                    Set Rs = New ADODB.Recordset
                    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    If Not Rs.EOF Then
                        txtAux2(3).Text = DBLet(Rs.Fields(0).Value, "N")
                        txtAux2(2).Text = DBLet(Rs.Fields(1).Value, "N")
                        txtAux2(4).Text = DBLet(Rs.Fields(2).Value, "N")
                    End If
                    Set Rs = Nothing
                
                End If
            End If
        
        Case 1 'Variedad
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(0).Text = PonerNombreDeCod(txtAux(Index), "variedades", "nomvarie")
                If txtAux2(0).Text = "" Then
                    cadMen = "No existe la Variedad: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmVar = New frmManVariedad
                        frmVar.DatosADevolverBusqueda = "0|1|"
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmVar.Show vbModal
                        Set frmVar = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Adoaux(0), 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
            End If
        
        Case 4, 5, 3 ' horas, importe y complemento
            PonerFormatoDecimal txtAux(Index), 3
        
            If Index = 4 Then
                PrecioHora = txtAux2(4).Text
                txtAux(5).Text = Round2(ComprobarCero(txtAux(4).Text) * PrecioHora, 2)
                PonerFormatoDecimal txtAux(5), 1
            End If
        
    End Select
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
   If Not txtAux(Index).MultiLine Then ConseguirFocoLin txtAux(Index)
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) Then
            Select Case Index
                Case 2: KEYBusquedaBtn KeyAscii, 4 'trabajador
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
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
    
    ' ******************************************************************************
    DatosOkLlin = B
    
EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function SepuedeBorrar(ByRef Index As Integer) As Boolean
    SepuedeBorrar = False
    
    ' *** si cal comprovar alguna cosa abans de borrar ***
'    Select Case Index
'        Case 0 'cuentas bancarias
'            If AdoAux(Index).Recordset!ctaprpal = 1 Then
'                MsgBox "No puede borrar una Cuenta Principal. Seleccione antes otra cuenta como Principal", vbExclamation
'                Exit Function
'            End If
'    End Select
    ' ****************************************************
    
    SepuedeBorrar = True
End Function

Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
    Indice = Index + 3
     Select Case Index
        Case 0 'Capataces
            Set frmCap = New frmManCapataz
            frmCap.DatosADevolverBusqueda = "0|1|"
            frmCap.CodigoActual = Text1(6).Text
            frmCap.Show vbModal
            Set frmCap = Nothing
            PonerFoco Text1(6)
    
        Case 1 'campos
            indCodigo = 2
            Set frmCam = New frmManCampos
            frmCam.DatosADevolverBusqueda = "0|"
            frmCam.Show vbModal
            Set frmCam = Nothing
            PonerFoco Text1(indCodigo)
        Case 2 'campos
            indCodigo = 8
            Set frmCam = New frmManCampos
            frmCam.DatosADevolverBusqueda = "0|"
            frmCam.Show vbModal
            Set frmCam = Nothing
            PonerFoco Text1(indCodigo)
        Case 3 'campos
            indCodigo = 11
            Set frmCam = New frmManCampos
            frmCam.DatosADevolverBusqueda = "0|"
            frmCam.Show vbModal
            Set frmCam = Nothing
            PonerFoco Text1(indCodigo)
        Case 4 'campos
            indCodigo = 7
            Set frmCam = New frmManCampos
            frmCam.DatosADevolverBusqueda = "0|"
            frmCam.Show vbModal
            Set frmCam = Nothing
            PonerFoco Text1(indCodigo)
    
    
        Case 99 'variedades
            Set frmVar = New frmComVar
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.CodigoActual = Text1(3).Text
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco Text1(3)
        Case 100 'socios
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
'            frmSoc.CodigoActual = Text1(4).Text
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco Text1(4)
    
    
    
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub


Private Sub CargaFrame(Index As Integer, enlaza As Boolean)
Dim tip As Integer
Dim i As Byte

    Adoaux(Index).ConnectionString = conn
    Adoaux(Index).RecordSource = MontaSQLCarga(Index, enlaza)
    Adoaux(Index).CursorType = adOpenDynamic
    Adoaux(Index).LockType = adLockPessimistic
    Adoaux(Index).Refresh
    
    If Not Adoaux(Index).Recordset.EOF Then
        PonerCamposForma2 Me, Adoaux(Index), 2, "FrameAux" & Index
    Else
        ' *** si n'hi han tabs sense datagrids, li pose els valors als camps ***
        NetejaFrameAux "FrameAux3" 'neteja només lo que te TAG
    End If
End Sub

' *** si n'hi han tabs sense datagrids ***
Private Sub NetejaFrameAux(nom_frame As String)
Dim Control As Object
    
    For Each Control In Me.Controls
        If (Control.Tag <> "") Then
            If (Control.Container.Name = nom_frame) Then
                If TypeOf Control Is TextBox Then
                    Control.Text = ""
                ElseIf TypeOf Control Is ComboBox Then
                    Control.ListIndex = -1
                End If
            End If
        End If
    Next Control

End Sub

Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
Dim B As Boolean
Dim i As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    CargaGridGnral Me.DataGridAux(Index), Me.Adoaux(Index), tots, PrimeraVez
    
    Select Case Index
        Case 0 'horas
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N|||||;" 'codigo de usuario
            tots = tots & "S|txtAux(0)|T|Campo|1100|;"
            tots = tots & "S|txtAux(1)|T|Código|900|;S|txtAux2(0)|T|Variedad|2500|;S|txtAux(2)|T|Codigo|900|;S|btnBuscar(4)|B||195|;"
            tots = tots & "S|txtAux2(1)|T|Trabajador|3500|;S|txtAux2(3)|T|Salar|900|;"
            tots = tots & "S|txtAux2(2)|T|Descripcion|2500|;S|txtAux2(4)|T|Precio|1500|;"
            tots = tots & "S|txtAux(4)|T|Horas|1200|;S|txtAux(5)|T|Importe|1430|;S|txtAux(3)|T|Compleme|1300|;"

            arregla tots, DataGridAux(Index), Me, 350
            
            DataGridAux(0).Columns(6).Alignment = dbgRight
       
            B = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
    End Select
    
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
    CalcularTotales
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub

Private Sub InsertarLinea()
'Inserta registre en les taules de Llínies
Dim nomframe As String
Dim B As Boolean

    On Error Resume Next

 
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        txtAux(7).Text = vUsu.Codigo
        If InsertarDesdeForm2(Me, 2, "FrameAux0") Then
            ' *** si n'hi ha que fer alguna cosa abas d'insertar
            ' *************************************************
            '++monica: en caso de estar insertando seccion y que no existan las
            'cuentas contables hacemos esto para que las inserte en contabilidad.
'            If NumTabMto = 1 Then
'               txtAux2(4).Text = PonerNombreCuenta(txtaux1(4), 3, Text1(0))
'               txtAux2(5).Text = PonerNombreCuenta(txtaux1(5), 3, Text1(0))
'            End If
            
            Select Case NumTabMto
                Case 0, 1 ' *** els index de les llinies en grid (en o sense tab) ***
                    CargaGrid NumTabMto, True
                    BotonAnyadirLinea NumTabMto
'                Case 3 ' *** els index dels tabs que NO tenen grid ***
'                    CargaFrame 3, True
'                    If b Then BotonModificar
'                    ModoLineas = 0
'                LLamaLineas NumTabMto, 0
            End Select
           
        End If
    End If
End Sub


Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vtabla As String
Dim anc As Single
Dim i As Integer

    ModoLineas = 1 'Posem Modo Afegir Llínia

    PonerModo 5

    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If



    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 0: vtabla = "horasmasivo"
    End Select
    ' ********************************************************

    vWhere = ObtenerWhereCab(False)

    Select Case Index
         Case 0

            AnyadirLinea DataGridAux(Index), Adoaux(Index)

            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 240
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If

            LLamaLineas Index, ModoLineas, anc

            For i = 0 To txtAux.Count - 1
                txtAux(i).Text = ""
            Next i
            
            txtAux(6).Text = vUsu.Codigo
            
            txtAux2(0).Text = "" 'calidad
            txtAux2(1).Text = ""
            txtAux2(2).Text = ""
            PonerFoco txtAux(0)


    End Select
End Sub



Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Integer
Dim SQL As String

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    If DatosOkLlin("FrameAux0") Then
        SQL = "update horasmasivo set horasdia = " & DBSet(txtAux(4), "N")
        SQL = SQL & " , importe = " & DBSet(txtAux(5).Text, "N")
        SQL = SQL & " , compleme = " & DBSet(txtAux(3).Text, "N")
        SQL = SQL & " where codusu = " & DBSet(vUsu.Codigo, "N")
        SQL = SQL & " and codcampo = " & DBSet(txtAux(0).Text, "N")
        SQL = SQL & " and codvarie = " & DBSet(txtAux(1).Text, "N")
        SQL = SQL & " and codtraba = " & DBSet(txtAux(2).Text, "N")
        
        conn.Execute SQL
        ModoLineas = 0
        V = Adoaux(NumTabMto).Recordset.Fields(4) 'el 2 es el nº de llinia
        CargaGrid NumTabMto, True
            
        ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
        PonerFocoGrid Me.DataGridAux(0)
        Adoaux(NumTabMto).Recordset.Find (Adoaux(NumTabMto).Recordset.Fields(4).Name & " =" & V)
            
        LLamaLineas NumTabMto, 0
        ModificarLinea = True
        PonerModo 5
    End If
End Function

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " codusu=" & DBSet(vUsu.Codigo, "N")
    
    ObtenerWhereCab = vWhere
End Function


'' *** neteja els camps dels tabs de grid que
''estan fora d'este, i els camps de descripció ***
Private Sub LimpiarCamposFrame(Index As Integer)
    On Error Resume Next
 
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CalcularTotales()
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim Sql2 As String
Dim TotalEnvases As String
Dim TotalCostes As String
Dim Total As Long
Dim Valor As Currency
Dim i As Integer

    On Error Resume Next

    SQL = "select sum(horasdia) horasdia, sum(importe) importe, sum(compleme) complemento "
    SQL = SQL & " from horasmasivo "
    SQL = SQL & " where codusu = " & DBSet(vUsu.Codigo, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Total = 0
    If Not Rs.EOF Then
        Text3(0).Text = DBLet(Rs.Fields(0).Value, "N")
        Text3(0).Text = Format(Text3(0).Text, "###,##0.00")
        Text3(1).Text = DBLet(Rs.Fields(1).Value, "N")
        Text3(1).Text = Format(Text3(1).Text, "###,###,##0.00")
        Text3(2).Text = DBLet(Rs.Fields(2).Value, "N")
        Text3(2).Text = Format(Text3(2).Text, "###,###,##0.00")
    End If

End Sub


Private Sub PasarSigReg()
'Nos situamos en el siguiente registro
    If Me.DataGridAux(0).Bookmark < Me.Adoaux(0).Recordset.RecordCount Then
'        DataGridAux(0).Row = DataGridAux(0).Row + 1
        DataGridAux(0).Bookmark = DataGridAux(0).Bookmark + 1
        BotonModificarLinea 0
    ElseIf DataGridAux(0).Bookmark = Adoaux(0).Recordset.RecordCount Then
'        PonerFocoBtn Me.cmdAceptar
        BotonModificarLinea 0
    End If
End Sub


Private Sub PasarAntReg()
'Nos situamos en el siguiente registro
    If Me.DataGridAux(0).Bookmark > 1 Then
'        DataGridAux(0).Row = DataGridAux(0).Row + 1
        DataGridAux(0).Bookmark = DataGridAux(0).Bookmark - 1
        BotonModificarLinea 0
    ElseIf DataGridAux(0).Bookmark = 1 Then
'        PonerFocoBtn Me.cmdAceptar
        BotonModificarLinea 0
    End If
End Sub



Private Sub InsertarRegistros()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim SQL As String
Dim SQLinsert As String
Dim SqlValues As String
Dim Variedad As String
Dim actualiza As Boolean
Dim NumF As Long
Dim Importe As Currency
Dim Trabajador As String
Dim Precio As String
Dim Rs As ADODB.Recordset
Dim i As Integer

    On Error GoTo eInsertarRegistros


    SQL = "delete from horasmasivo where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute SQL
    
    SQLinsert = "insert into horasmasivo (codusu, codcampo, codvarie, codtraba, horasdia, compleme, importe) values "
    
    SQL = " select rcuadrilla_trabajador.codtraba, salarios.impsalar from (rcuadrilla inner join rcuadrilla_trabajador on rcuadrilla.codcuadrilla = rcuadrilla_trabajador.codcuadrilla) inner join straba on rcuadrilla_trabajador.codtraba = straba.codtraba "
    SQL = SQL & " inner join salarios on straba.codcateg = salarios.codcateg "
    SQL = SQL & " where rcuadrilla.codcapat = " & DBSet(Text1(3).Text, "N")
    SQL = SQL & " and (straba.fechabaja is null or fechabaja = '')"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SqlValues = ""
    
    
    While Not Rs.EOF
    
        Trabajador = DBLet(Rs!CodTraba, "N")
    
        Precio = DBLet(Rs.Fields(1).Value, "N")
    
        For i = 1 To 4
            
            Select Case i
                Case 1
                    If Text1(2) <> "" Then
                        Variedad = DevuelveValor("select codvarie from rcampos where codcampo = " & DBSet(Text1(2), "N"))
                        
                        SqlValues = SqlValues & ",(" & DBSet(vUsu.Codigo, "N") & "," & DBSet(Text1(2), "N") & "," & DBSet(Variedad, "N") & ", " & DBSet(Trabajador, "N") & ","
                        SqlValues = SqlValues & DBSet(Text1(5).Text, "N") & ","
                        
'                        If Escapataz(Trabajador, Text1(3)) Then
'                            SqlValues = SqlValues & DBSet(Text1(0), "N") & ","
'                        Else
                            SqlValues = SqlValues & "0,"
'                        End If
                        Importe = Round2(Precio * ImporteSinFormato(Text1(5)), 2)
                        SqlValues = SqlValues & DBSet(Importe, "N") & ")"
                    End If
                Case 2
                    If Text1(8) <> "" Then
                        Variedad = DevuelveValor("select codvarie from rcampos where codcampo = " & DBSet(Text1(8), "N"))
                        
                        SqlValues = SqlValues & ",(" & DBSet(vUsu.Codigo, "N") & "," & DBSet(Text1(8), "N") & "," & DBSet(Variedad, "N") & ", " & DBSet(Trabajador, "N") & ","
                        SqlValues = SqlValues & DBSet(Text1(6).Text, "N") & ","
                        
'                        If Escapataz(Trabajador, Text1(3)) Then
'                            SqlValues = SqlValues & DBSet(Text1(0), "N") & ","
'                        Else
                            SqlValues = SqlValues & "0,"
'                        End If
                        Importe = Round2(Precio * ImporteSinFormato(Text1(6)), 2)
                        SqlValues = SqlValues & DBSet(Importe, "N") & ")"
                    End If
                
                Case 3
                    If Text1(11) <> "" Then
                        Variedad = DevuelveValor("select codvarie from rcampos where codcampo = " & DBSet(Text1(11), "N"))
                        
                        SqlValues = SqlValues & ",(" & DBSet(vUsu.Codigo, "N") & "," & DBSet(Text1(11), "N") & "," & DBSet(Variedad, "N") & ", " & DBSet(Trabajador, "N") & ","
                        SqlValues = SqlValues & DBSet(Text1(10).Text, "N") & ","
                        
'                        If Escapataz(Trabajador, Text1(3)) Then
'                            SqlValues = SqlValues & DBSet(Text1(0), "N") & ","
'                        Else
                            SqlValues = SqlValues & "0,"
'                        End If
                        Importe = Round2(Precio * ImporteSinFormato(Text1(10)), 2)
                        SqlValues = SqlValues & DBSet(Importe, "N") & ")"
                    End If
                
                Case 4
                    If Text1(7) <> "" Then
                        Variedad = DevuelveValor("select codvarie from rcampos where codcampo = " & DBSet(Text1(7), "N"))
                        
                        SqlValues = SqlValues & ",(" & DBSet(vUsu.Codigo, "N") & "," & DBSet(Text1(7), "N") & "," & DBSet(Variedad, "N") & ", " & DBSet(Trabajador, "N") & ","
                        SqlValues = SqlValues & DBSet(Text1(4).Text, "N") & ","
                        
'                        If Escapataz(Trabajador, Text1(3)) Then
'                            SqlValues = SqlValues & DBSet(Text1(0), "N") & ","
'                        Else
                            SqlValues = SqlValues & "0,"
'                        End If
                        Importe = Round2(Precio * ImporteSinFormato(Text1(4)), 2)
                        SqlValues = SqlValues & DBSet(Importe, "N") & ")"
                    End If
            
            End Select
        
        Next i
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    If SqlValues <> "" Then
        conn.Execute SQLinsert & Mid(SqlValues, 2)
    End If
    
    'si el complemento es cero no hacemos nada
    Dim Traba As String
    Traba = DevuelveValor("select codtraba from rcapataz where codcapat = " & DBSet(Text1(3).Text, "N"))
    If ComprobarCero(Text1(0).Text) <> 0 Then
        SQL = "update horasmasivo set compleme = " & DBSet(Text1(0).Text, "N")
        SQL = SQL & " where codusu = " & vUsu.Codigo
        SQL = SQL & " and codtraba = " & DBSet(Traba, "N")
        SQL = SQL & " and codcampo = " & DBSet(Text1(2).Text, "N")
        
        conn.Execute SQL
    End If
    
    Exit Sub
    
eInsertarRegistros:
    MuestraError Err.Number, "Insertando Registros", Err.Description
End Sub

Private Function Escapataz(codtra As String, codcapat As String) As Boolean
Dim SQL As String

    SQL = "select codtraba from rcapataz where codcapat = " & DBSet(codcapat, "N")
    
    Escapataz = (codtra = DevuelveValor(SQL))

End Function


Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim SQL As String

    On Error Resume Next
    
    SQL = " codusu= " & DBSet(vUsu.Codigo, "N")
    
    If conWhere Then SQL = " WHERE " & SQL
    ObtenerWhereCP = SQL
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function



