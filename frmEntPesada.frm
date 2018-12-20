VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEntPesada 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entrada de Pesada"
   ClientHeight    =   9825
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   13995
   Icon            =   "frmEntPesada.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   13995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   120
      TabIndex        =   108
      Top             =   30
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   109
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
      Left            =   3780
      TabIndex        =   106
      Top             =   30
      Width           =   885
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   210
         TabIndex        =   107
         Top             =   180
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Notas de Campo"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   4770
      TabIndex        =   104
      Top             =   30
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   105
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
               Object.ToolTipText     =   "Primero"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Anterior"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Siguiente"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Último"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.CheckBox chkVistaPrevia 
      Caption         =   "Vista previa"
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
      Left            =   10950
      TabIndex        =   103
      Top             =   270
      Width           =   1605
   End
   Begin VB.Frame Frame4 
      Height          =   705
      Left            =   6660
      TabIndex        =   41
      Top             =   4440
      Width           =   6885
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Left            =   1860
         MaxLength       =   7
         TabIndex        =   42
         Top             =   210
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Kilos por Cajon"
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
         Left            =   270
         TabIndex        =   43
         Top             =   240
         Width           =   1545
      End
   End
   Begin VB.Frame FrameDatosNota 
      Caption         =   "Datos Notas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   2835
      Left            =   6660
      TabIndex        =   44
      Top             =   1470
      Width           =   6885
      Begin VB.TextBox Text5 
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
         Left            =   1290
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   59
         Top             =   1140
         Width           =   5295
      End
      Begin VB.TextBox Text4 
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
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   58
         Top             =   1545
         Width           =   4305
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
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
         Left            =   1290
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   57
         Top             =   1545
         Width           =   945
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
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   56
         Top             =   750
         Width           =   4305
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         Left            =   1290
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   55
         Top             =   750
         Width           =   945
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
         Index           =   8
         Left            =   1290
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   49
         Text            =   "Text2"
         Top             =   2340
         Width           =   5295
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
         Left            =   1290
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   48
         Text            =   "Text2"
         Top             =   1950
         Width           =   5295
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
         Index           =   6
         Left            =   1290
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   45
         Text            =   "Text2"
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label Label1 
         Caption         =   "Zona"
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
         Index           =   26
         Left            =   240
         TabIndex        =   62
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label28 
         Caption         =   "Poblacion"
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
         Left            =   240
         TabIndex        =   61
         Top             =   1170
         Width           =   1065
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
         Index           =   0
         Left            =   240
         TabIndex        =   60
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label11 
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
         Left            =   210
         TabIndex        =   50
         Top             =   2010
         Width           =   1185
      End
      Begin VB.Label Label5 
         Caption         =   "Tarifa"
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
         Left            =   210
         TabIndex        =   47
         Top             =   2400
         Width           =   1185
      End
      Begin VB.Label Label4 
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
         Left            =   240
         TabIndex        =   46
         Top             =   420
         Width           =   1185
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4545
      Left            =   90
      TabIndex        =   17
      Top             =   780
      Width           =   13695
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   11
         Left            =   180
         MaxLength       =   7
         TabIndex        =   10
         Tag             =   "Tara Vehiculo|N|S|0|999999|rpesadas|taravehi|###,##0||"
         Top             =   4020
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         BeginProperty Font 
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
         Left            =   4920
         MaxLength       =   7
         TabIndex        =   3
         Tag             =   "Peso Bruto|N|N|||rpesadas|kilosbrut|###,##0||"
         Top             =   810
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   2
         Left            =   2070
         MaxLength       =   7
         TabIndex        =   11
         Tag             =   "Tara Vehiculo|N|S|0|999999|rpesadas|otratara|###,##0||"
         Top             =   4020
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   13
         Left            =   2070
         MaxLength       =   5
         TabIndex        =   5
         Tag             =   "Nro.Cajas 1|N|S|||rpesadas|numcajo1|#,##0||"
         Top             =   1860
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   18
         Left            =   4905
         MaxLength       =   7
         TabIndex        =   73
         Tag             =   "Tara 1|N|S|0|999999|rpesadas|taracaja1|###,##0||"
         Top             =   1860
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   19
         Left            =   4905
         MaxLength       =   7
         TabIndex        =   72
         Tag             =   "Tara 2|N|S|0|999999|rpesadas|taracaja2|###,##0||"
         Top             =   2220
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   20
         Left            =   4905
         MaxLength       =   7
         TabIndex        =   71
         Tag             =   "Tara 3|N|S|0|999999|rpesadas|taracaja3|###,##0||"
         Top             =   2580
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   9
         Left            =   4905
         MaxLength       =   7
         TabIndex        =   70
         Tag             =   "Tara 4|N|S|0|999999|rpesadas|taracaja4|###,##0||"
         Top             =   2940
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   10
         Left            =   4905
         MaxLength       =   7
         TabIndex        =   69
         Tag             =   "Tara 5|N|S|0|999999|rpesadas|taracaja5|###,##0||"
         Top             =   3300
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   14
         Left            =   2070
         MaxLength       =   5
         TabIndex        =   6
         Tag             =   "Nro.Cajas 2|N|S|||rpesadas|numcajo2|#,##0||"
         Top             =   2220
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   15
         Left            =   2070
         MaxLength       =   5
         TabIndex        =   7
         Tag             =   "Nro.Cajas 3|N|S|||rpesadas|numcajo3|#,##0||"
         Top             =   2580
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   16
         Left            =   2070
         MaxLength       =   5
         TabIndex        =   8
         Tag             =   "Nro.Cajas 4|N|S|||rpesadas|numcajo4|#,##0||"
         Top             =   2940
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   17
         Left            =   2070
         MaxLength       =   5
         TabIndex        =   9
         Tag             =   "Nro.Cajas 5|N|S|||rpesadas|numcajo5|#,##0||"
         Top             =   3300
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   23
         Left            =   4905
         MaxLength       =   7
         TabIndex        =   68
         Tag             =   "Tara 1|N|S|0|999999|rpesadas|taracaja0|###,##0||"
         Top             =   1500
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   22
         Left            =   2070
         MaxLength       =   5
         TabIndex        =   4
         Tag             =   "Nro.Cajas Reales|N|S|||rpesadas|cajonesrea|#,##0||"
         Top             =   1500
         Width           =   1215
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
         Left            =   6660
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Código Transporte|T1|S|||rpesadas|codtrans|||"
         Top             =   210
         Width           =   1155
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
         Left            =   3510
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Pesada|F|N|||rpesadas|fecpesada|dd/mm/yyyy|N|"
         Top             =   210
         Width           =   1240
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
         Left            =   7860
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   20
         Text            =   "Text2"
         Top             =   210
         Width           =   5265
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         BeginProperty Font 
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
         Left            =   1230
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "NºPesada|N|S|||rpesadas|nropesada|0000000|S|"
         Text            =   "Text1 7"
         Top             =   210
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         BeginProperty Font 
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
         Left            =   4890
         MaxLength       =   7
         TabIndex        =   95
         Tag             =   "Peso Neto|N|N|0|999999|rpesadas|kilosnet|###,##0||"
         Top             =   3840
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
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
         Index           =   24
         Left            =   4890
         MaxLength       =   7
         TabIndex        =   102
         Tag             =   "Peso Neto|N|N|0|999999|rpesadas|kilostra|###,##0||"
         Top             =   3840
         Width           =   1155
      End
      Begin VB.Label Label9 
         Caption         =   "Pesos y Taras:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   225
         Left            =   210
         TabIndex        =   100
         Top             =   900
         Width           =   1755
      End
      Begin VB.Label Label17 
         Caption         =   "Peso Bruto"
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
         Left            =   3690
         TabIndex        =   99
         Top             =   840
         Width           =   1185
      End
      Begin VB.Label Label8 
         Caption         =   "Tara Vehiculo"
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
         TabIndex        =   98
         Top             =   3780
         Width           =   1515
      End
      Begin VB.Label Label7 
         Caption         =   "Peso Neto"
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
         Left            =   3690
         TabIndex        =   97
         Top             =   3855
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "Otras Taras"
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
         Left            =   2070
         TabIndex        =   96
         Top             =   3780
         Width           =   1185
      End
      Begin VB.Label Label13 
         Caption         =   "Cajas"
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
         Left            =   2100
         TabIndex        =   94
         Top             =   1200
         Width           =   1185
      End
      Begin VB.Line Line3 
         X1              =   180
         X2              =   6030
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line2 
         X1              =   195
         X2              =   6090
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label16 
         Caption         =   "Tara"
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
         Left            =   4920
         TabIndex        =   93
         Top             =   1200
         Width           =   1185
      End
      Begin VB.Label Label14 
         Caption         =   "Peso Caja"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3690
         TabIndex        =   92
         Top             =   1200
         Width           =   1185
      End
      Begin VB.Label Label19 
         Caption         =   "x  Peso 1  "
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
         Left            =   3540
         TabIndex        =   91
         Top             =   1890
         Width           =   840
      End
      Begin VB.Label Label15 
         Caption         =   "Tarifa"
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
         Left            =   195
         TabIndex        =   90
         Top             =   1890
         Width           =   1830
      End
      Begin VB.Label Label15 
         Caption         =   "Tarifa"
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
         Left            =   195
         TabIndex        =   89
         Top             =   2250
         Width           =   1830
      End
      Begin VB.Label Label15 
         Caption         =   "Tarifa"
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
         Left            =   195
         TabIndex        =   88
         Top             =   2610
         Width           =   1830
      End
      Begin VB.Label Label15 
         Caption         =   "Tarifa"
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
         Left            =   195
         TabIndex        =   87
         Top             =   2970
         Width           =   1830
      End
      Begin VB.Label Label15 
         Caption         =   "Tarifa"
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
         Left            =   195
         TabIndex        =   86
         Top             =   3330
         Width           =   1830
      End
      Begin VB.Label Label19 
         Caption         =   "x  Peso 1"
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
         Left            =   3540
         TabIndex        =   85
         Top             =   2250
         Width           =   705
      End
      Begin VB.Label Label19 
         Caption         =   "x  Peso 1"
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
         Left            =   3540
         TabIndex        =   84
         Top             =   2610
         Width           =   705
      End
      Begin VB.Label Label19 
         Caption         =   "x  Peso 1"
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
         Left            =   3540
         TabIndex        =   83
         Top             =   2970
         Width           =   705
      End
      Begin VB.Label Label19 
         Caption         =   "x  Peso 1"
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
         Left            =   3540
         TabIndex        =   82
         Top             =   3330
         Width           =   705
      End
      Begin VB.Label Label10 
         Caption         =   "= "
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
         Left            =   4710
         TabIndex        =   81
         Top             =   1890
         Width           =   150
      End
      Begin VB.Label Label10 
         Caption         =   "= "
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
         Left            =   4710
         TabIndex        =   80
         Top             =   2250
         Width           =   150
      End
      Begin VB.Label Label10 
         Caption         =   "= "
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
         Left            =   4710
         TabIndex        =   79
         Top             =   2610
         Width           =   150
      End
      Begin VB.Label Label10 
         Caption         =   "= "
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
         Left            =   4710
         TabIndex        =   78
         Top             =   2970
         Width           =   150
      End
      Begin VB.Label Label10 
         Caption         =   "= "
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
         Left            =   4710
         TabIndex        =   77
         Top             =   3330
         Width           =   150
      End
      Begin VB.Label Label10 
         Caption         =   "= "
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
         Left            =   4710
         TabIndex        =   76
         Top             =   1530
         Width           =   150
      End
      Begin VB.Label Label15 
         Caption         =   "CAJONES REALES"
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
         TabIndex        =   75
         Top             =   1530
         Width           =   1830
      End
      Begin VB.Label Label19 
         Caption         =   "x  Peso 1  "
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
         Left            =   3540
         TabIndex        =   74
         Top             =   1530
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Iva"
         Height          =   255
         Index           =   14
         Left            =   10410
         TabIndex        =   40
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha "
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
         Left            =   2550
         TabIndex        =   21
         Top             =   240
         Width           =   585
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   3210
         Picture         =   "frmEntPesada.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Transportista"
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
         Left            =   4860
         TabIndex        =   19
         Top             =   240
         Width           =   1350
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   6300
         ToolTipText     =   "Buscar Transportista"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "NºPesada"
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
         Left            =   210
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   8
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   66
      Top             =   900
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   65
      Top             =   960
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   6
      Left            =   4290
      MaxLength       =   10
      TabIndex        =   64
      Top             =   900
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   4290
      MaxLength       =   10
      TabIndex        =   63
      Top             =   900
      Width           =   1065
   End
   Begin VB.Frame FrameAux1 
      Caption         =   "Notas de Campo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   3840
      Left            =   90
      TabIndex        =   22
      Top             =   5370
      Width           =   13695
      Begin VB.TextBox txtAux 
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
         Height          =   360
         Index           =   11
         Left            =   12780
         MaxLength       =   7
         TabIndex        =   101
         Tag             =   "Peso Transp|N|N|0|999999|rentradas|kilostra|###,##0||"
         Text            =   "kilostr"
         Top             =   2250
         Visible         =   0   'False
         Width           =   615
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
         Left            =   10470
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Tag             =   "Transportado por|N|N|0|1|rentradas|transportadopor||N|"
         Top             =   2250
         Width           =   750
      End
      Begin VB.TextBox txtAux 
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
         Height          =   360
         Index           =   2
         Left            =   12330
         MaxLength       =   7
         TabIndex        =   67
         Tag             =   "Peso Bruto|N|N|0|999999|rentradas|kilosbru|###,##0||"
         Text            =   "pesobru"
         Top             =   2250
         Visible         =   0   'False
         Width           =   615
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
         Height          =   300
         Index           =   4
         Left            =   8700
         MaskColor       =   &H00000000&
         TabIndex        =   54
         ToolTipText     =   "Buscar Tarifa"
         Top             =   2250
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
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
         Height          =   360
         Index           =   0
         Left            =   7080
         MaxLength       =   4
         TabIndex        =   27
         Tag             =   "Código Capataz|N|S|0|9999|rentradas|codcapat|0000||"
         Text            =   "capa"
         Top             =   2250
         Visible         =   0   'False
         Width           =   735
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
         Left            =   9690
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Tag             =   "Recolectado|N|N|0|1|rentradas|recolect||N|"
         Top             =   2250
         Width           =   750
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
         Left            =   8910
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Tag             =   "Tipo Entrada|N|N|0|3|rentradas|tipoentr||N|"
         Top             =   2250
         Width           =   750
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
         Height          =   300
         Index           =   3
         Left            =   7800
         MaskColor       =   &H00000000&
         TabIndex        =   53
         ToolTipText     =   "Buscar Capataz"
         Top             =   2250
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
         Height          =   300
         Index           =   2
         Left            =   6870
         MaskColor       =   &H00000000&
         TabIndex        =   52
         ToolTipText     =   "Buscar Campo"
         Top             =   2250
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
         Height          =   300
         Index           =   1
         Left            =   5910
         MaskColor       =   &H00000000&
         TabIndex        =   51
         ToolTipText     =   "Buscar Socio"
         Top             =   2250
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
         Height          =   300
         Index           =   0
         Left            =   3690
         MaskColor       =   &H00000000&
         TabIndex        =   37
         ToolTipText     =   "Buscar Variedad"
         Top             =   2250
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox Text2 
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
         Height          =   360
         Index           =   5
         Left            =   3930
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   36
         Text            =   "Nombre variedad"
         Top             =   2250
         Width           =   1200
      End
      Begin VB.TextBox txtAux 
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
         Height          =   360
         Index           =   1
         Left            =   300
         MaxLength       =   12
         TabIndex        =   35
         Tag             =   "Num.Pesada|N|N|||rentradas|nropesada|0000000|S|"
         Text            =   "NumPes"
         Top             =   2250
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox txtAux 
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
         Height          =   360
         Index           =   5
         Left            =   2760
         MaxLength       =   6
         TabIndex        =   24
         Tag             =   "Variedad|N|N|||rentradas|codvarie|000000|N|"
         Text            =   "variedad"
         Top             =   2250
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtAux 
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
         Height          =   360
         Index           =   6
         Left            =   5160
         MaxLength       =   6
         TabIndex        =   25
         Tag             =   "Socio|N|N|||rentradas|codsocio|000000||"
         Text            =   "socio"
         Top             =   2250
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
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
         Height          =   360
         Index           =   3
         Left            =   1050
         MaxLength       =   12
         TabIndex        =   34
         Tag             =   "Num.Linea|N|N|||rentradas|numlinea|000|S|"
         Text            =   "Linea"
         Top             =   2250
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtAux 
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
         Height          =   360
         Index           =   7
         Left            =   6120
         MaxLength       =   8
         TabIndex        =   26
         Tag             =   "Campo|N|N|||rentradas|codcampo|00000000||"
         Text            =   "campo"
         Top             =   2250
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
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
         Height          =   360
         Index           =   8
         Left            =   7920
         MaxLength       =   3
         TabIndex        =   28
         Tag             =   "Código Tarifa|N|S|0|999|rentradas|codtarif|000||"
         Text            =   "ta"
         Top             =   2250
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux 
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
         Height          =   360
         Index           =   9
         Left            =   11190
         MaxLength       =   12
         TabIndex        =   32
         Tag             =   "Num.Cajon1|N|N|||rentradas|numcajo1|###,##0||"
         Text            =   "caja"
         Top             =   2250
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.TextBox txtAux 
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
         Height          =   360
         Index           =   10
         Left            =   11760
         MaxLength       =   7
         TabIndex        =   33
         Tag             =   "Peso Neto|N|N|0|999999|rentradas|kilosnet|###,##0||"
         Text            =   "pesonet"
         Top             =   2250
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux 
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
         Height          =   360
         Index           =   4
         Left            =   1740
         MaxLength       =   7
         TabIndex        =   23
         Tag             =   "Nro Nota|N|N|||rentradas|numnotac|0000000|N|"
         Text            =   "Nota"
         Top             =   2250
         Visible         =   0   'False
         Width           =   900
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   1
         Left            =   270
         TabIndex        =   38
         Top             =   300
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
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
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Salir"
               Object.Tag             =   "2"
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "frmEntPesada.frx":0097
         Height          =   2760
         Left            =   270
         TabIndex        =   39
         Top             =   780
         Width           =   13280
         _ExtentX        =   23416
         _ExtentY        =   4868
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adoaux 
         Height          =   330
         Index           =   1
         Left            =   1680
         Top             =   300
         Visible         =   0   'False
         Width           =   1425
         _ExtentX        =   2514
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
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   0
         Left            =   8700
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   480
         Width           =   240
      End
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   9210
      Width           =   2175
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
         Left            =   150
         TabIndex        =   16
         Top             =   180
         Width           =   1785
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
      Left            =   12615
      TabIndex        =   13
      Top             =   9330
      Width           =   1095
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
      Left            =   11400
      TabIndex        =   12
      Top             =   9330
      Width           =   1095
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
      Left            =   12630
      TabIndex        =   14
      Top             =   9330
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   3000
      Top             =   1080
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   240
      Top             =   8040
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
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   240
      Top             =   8070
      Visible         =   0   'False
      Width           =   1425
      _ExtentX        =   2514
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
      Left            =   13320
      TabIndex        =   110
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
         Caption         =   "&Ver Todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnNotaCampo 
         Caption         =   "&Notas Campo"
         HelpContextID   =   2
         Shortcut        =   ^C
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
Attribute VB_Name = "frmEntPesada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 4004



'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public Albaran As String  ' venimos de albaranes para ver las facturas donde aparece el albaran

'========== VBLES PRIVADAS ====================
Private WithEvents frmEntPrev As frmEntPesadaPrev
Attribute frmEntPrev.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1

Private WithEvents frmVar As frmComVar 'variedades comercial
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmCap As frmManCapataz 'capataces
Attribute frmCap.VB_VarHelpID = -1
Private WithEvents frmTrans As frmManTranspor 'transportista
Attribute frmTrans.VB_VarHelpID = -1
Private WithEvents frmTar As frmManTarTra 'tarifa de transportista
Attribute frmTar.VB_VarHelpID = -1
Private WithEvents frmCamp As frmManCampos 'campos
Attribute frmCamp.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'mensajes
Attribute frmMens.VB_VarHelpID = -1

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

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

Dim EsCabecera As Boolean


Dim EsDeVarios As Boolean
'Si el cliente mostrado es de Varios o No

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1


Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal
Dim indice As Byte

Dim TipoFactura As Byte
Private BuscaChekc As String

Dim FechaAnt As String
Dim TransporAnt As String
Dim CajonreaAnt As String
Dim NetoAnt As String

Dim v_cadena As String

Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 0 'Variedades
            Set frmVar = New frmComVar
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco txtAux(5)
        Case 1
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco txtAux(6)
        Case 2
            PonerCamposSocioVariedad
            PonerFoco txtAux(7)
        Case 3
            Set frmCap = New frmManCapataz
            frmCap.DatosADevolverBusqueda = "0|1|"
            frmCap.CodigoActual = Text1(12).Text
            frmCap.Show vbModal
            Set frmCap = Nothing
            PonerFoco Text1(12)
        Case 4
            Set frmTar = New frmManTarTra
            frmTar.DeConsulta = True
            frmTar.DatosADevolverBusqueda = "0|1|"
            frmTar.CodigoActual = txtAux(8).Text
            frmTar.Show vbModal
            Set frmTar = Nothing
            PonerFoco txtAux(8)
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1

End Sub

Private Sub Check1_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "check1(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "check1(" & Index & ")|"
    End If
End Sub

Private Sub chkAbonos_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cmdAceptar_Click()
Dim i As Integer
Dim V As Integer


    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda

        Case 3  'AÑADIR
            If DatosOk Then InsertarCabecera

        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaCabecera Then
                
                    If CajonreaAnt <> Text1(22).Text Or NetoAnt <> Text1(12).Text Then
                        Data1.Refresh
                        Data1.Recordset.Find (Data1.Recordset.Fields(0).Name & " =" & Text1(0).Text)
                        mnNotaCampo_Click
                    Else
                        espera 0.2
                        TerminaBloquear
                        PosicionarData
                        PonerCampos
                        PonerCamposLineas
                    End If
                
                    
'                    If CajonreaAnt <> Text1(22).Text Or NetoAnt <> Text1(12).Text Then
'                        mnNotaCampo_Click
'                    End If
                    
                End If
            End If
            
         Case 5 'InsertarModificar LINEAS
            Select Case ModificaLineas
                Case 1 'afegir llínia
                    InsertarLinea
                Case 2 'modificar llínies
                    If ModificarLinea Then
                        V = Adoaux(1).Recordset.Fields(1) 'el 2 es el nº de llinia
                        CargaGrid DataGrid3, Adoaux(1), True
                
                        DataGrid3.SetFocus
                        Adoaux(1).Recordset.Find (Adoaux(1).Recordset.Fields(1).Name & " =" & V)
                
                        LLamaLineas ModificaLineas, 0, "DataGrid3"
                        
                        'PosicionarData
                        PonerModo 5
                    End If
            End Select
    End Select
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 0, 1, 3 'Busqueda, Insertar
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
            LLamaLineas Modo, 0, "DataGrid2"
            PonerFoco Text1(3)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(3)
            
        Case 5 'Lineas Detalle
            TerminaBloquear
            If ModificaLineas = 1 Then 'INSERTAR
                ModificaLineas = 0
                DataGrid3.AllowAddNew = False
                If Not Adoaux(1).Recordset.EOF Then Adoaux(1).Recordset.MoveFirst
            End If
            ModificaLineas = 0
            LLamaLineas Modo, 0, "DataGrid3"
            PonerModo 5
            DataGrid3.Enabled = True
            If Not Data1.Recordset.EOF Then _
                Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
            'Habilitar las opciones correctas del menu segun Modo
            PonerModoOpcionesMenu (Modo)
            PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
            DataGrid3.Enabled = True
            PonerFocoGrid DataGrid3
    
    End Select
End Sub
Private Sub BotonAnyadir()

    LimpiarCampos 'Huida els TextBox
    
'    TipoFactura = 1
    
    PonerModo 3
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la capçalera *******
'    Text1(0).Text = SugerirCodigoSiguienteStr("palets", "numpalet")
'    FormateaCampo Text1(0)
'    cmbAux(0).ListIndex = -1
    
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
'    Check1(0).Value = 1
        
    LimpiarDataGrids
    
    PonerFoco Text1(1) '*** 1r camp visible que siga PK ***
    
    ' *** si n'hi han camps de descripció a la capçalera ***
    'PosarDescripcions
End Sub


Private Sub BotonBuscar()
Dim anc As Single

    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PonerModo 1
        
'        'poner los txtaux para buscar por lineas de albaran
'        anc = DataGrid2.Top
'        If DataGrid2.Row < 0 Then
'            anc = anc + 440
'        Else
'            anc = anc + DataGrid2.RowTop(DataGrid2.Row) + 20
'        End If
'        LLamaLineas Modo, anc, "DataGrid2"
        
        
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbLightBlue
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbLightBlue
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
    If chkVistaPrevia.Value = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia ""
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select rpesadas.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla
'        CadenaConsulta = CadenaConsulta & " WHERE scafac.codtipom='" & CodTipoMov & "'"
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index, True
    PonerCampos
End Sub


Private Sub BotonModificar()
Dim DeVarios As Boolean

'    'solo se puede modificar la factura si no esta contabilizada
'    If Check1(0).Value = 1 Or Check1(1).Value = 1 Or Check1(2).Value = 1 Then
'        MsgBox "Esta factura no podemos modificarla", vbExclamation
'        TerminaBloquear
'        Exit Sub
'    End If
    
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    FechaAnt = Text1(1).Text
    TransporAnt = Text1(3).Text
    
    CajonreaAnt = Text1(22).Text
    NetoAnt = Text1(12).Text
    
    PonerFoco Text1(1) '*** 1r camp visible que siga PK ***
        
End Sub

Private Sub BotonNotaCampo()
Dim DeVarios As Boolean

    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 5
    
End Sub




Private Sub BotonModificarLinea(Index As Integer)
'Modificar una linea
Dim vWhere As String
Dim anc As Single
Dim J As Byte

    On Error GoTo eModificarLinea


    ModificaLineas = 2 'Modificar

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then  '1= Insertar
        TerminaBloquear
        Exit Sub
    End If
    
'--monica
'    If Data2.Recordset.EOF Then
'        TerminaBloquear
'        Exit Sub
'    End If
    
    
    PonerModo 5, Index
 

    
    vWhere = ObtenerWhereCP(False)
    vWhere = vWhere & " and numlinea=" & Adoaux(1).Recordset!NumLinea
    If Not BloqueaRegistro("rentradas", vWhere) Then
        TerminaBloquear
        Exit Sub
    End If
    If DataGrid3.Bookmark < DataGrid3.FirstRow Or DataGrid3.Bookmark > (DataGrid3.FirstRow + DataGrid3.VisibleRows - 1) Then
        J = DataGrid3.Bookmark - DataGrid3.FirstRow
        DataGrid3.Scroll 0, J
        DataGrid3.Refresh
    End If
    
'    anc = ObtenerAlto(Me.DataGrid1)
    anc = DataGrid3.Top
    If DataGrid3.Row < 0 Then
        anc = anc + 240
    Else
        anc = anc + DataGrid3.RowTop(DataGrid3.Row) + 5
    End If

    txtAux(1).Text = DataGrid3.Columns(0).Text ' nro de pesada
    txtAux(3).Text = DataGrid3.Columns(1).Text ' linea
    txtAux(4).Text = DataGrid3.Columns(2).Text ' nro de nota
    txtAux(5).Text = DataGrid3.Columns(3).Text ' variedad
    Text2(5).Text = DataGrid3.Columns(4).Text ' nombre de la variedad
    txtAux(6).Text = DataGrid3.Columns(5).Text ' socio
    txtAux(7).Text = DataGrid3.Columns(6).Text ' campo
    txtAux(0).Text = DataGrid3.Columns(7).Text ' capataz
    txtAux(8).Text = DataGrid3.Columns(8).Text ' tarifa
    txtAux(9).Text = DataGrid3.Columns(15).Text ' cajones
    txtAux(10).Text = DataGrid3.Columns(16).Text ' peso neto
    txtAux(2).Text = DataGrid3.Columns(17).Text ' peso bruto
    txtAux(11).Text = DataGrid3.Columns(18).Text ' peso bruto
    
    
    PosicionarCombo Combo1(0), DataGrid3.Columns(9).Text  ' tipo de entrada
    PosicionarCombo Combo1(1), DataGrid3.Columns(11).Text 'recolectado por
    PosicionarCombo Combo1(2), DataGrid3.Columns(13).Text 'transportado por
    

    BloquearTxt txtAux(4), True 'el nro de nota no se puede modificar
    'el peso neto no se inserta ni se modifica, siempre es calculado con el
    'nro de cajas
    BloquearTxt txtAux(10), True 'el peso neto tampoco
    
'    BloquearTxt txtAux(5), True
'    BloquearTxt txtAux(7), True
'    BloquearTxt txtAux(9), True
'    txtAux(4).Enabled = False
'    txtAux(5).Enabled = False
'    txtAux(7).Enabled = False
'    txtAux(9).Enabled = False
'
'    BloquearTxt txtAux(6), False
'    BloquearTxt txtAux(8), False
'
'    BloquearBtn Me.btnBuscar(0), True
    
    LLamaLineas ModificaLineas, anc, "DataGrid3"
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
    DataGrid3.Enabled = True
    
    PonerFoco txtAux(5)
    Me.DataGrid3.Enabled = False


eModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub BotonSalirNotas()
'Modificar una linea
Dim vWhere As String
Dim anc As Single
Dim J As Byte
Dim b As Boolean
Dim Mens As String

    On Error GoTo eSalirNotas

    conn.BeginTrans

    Mens = "Cuadrar pesada: "
    b = CuadrarPesada(Mens)
    If b Then
        Mens = "Actualizar chivato: "
        b = ActualizarChivato(Mens)
    End If

eSalirNotas:
    If Err.Number <> 0 Or Not b Then
        MsgBox Mens & vbCrLf & Err.Description, vbExclamation
    End If
    If b Then
        conn.CommitTrans
        ModificaLineas = 0
        CargaGrid DataGrid3, Adoaux(1), True
        PonerModo 2
    Else
        conn.RollbackTrans
        ModificaLineas = 0
        PonerModo 5
    End If
End Sub


Private Sub LLamaLineas(xModo As Byte, Optional alto As Single, Optional grid As String)
Dim jj As Integer
Dim b As Boolean
    
    Select Case grid
        Case "DataGrid3"
            DeseleccionaGrid Me.DataGrid3
            b = (xModo = 1 Or xModo = 2)
            
            txtAux(0).Height = DataGrid3.RowHeight - 10
            txtAux(0).Top = alto + 5
            txtAux(0).visible = b
            txtAux(0).Enabled = b
            
            
            For jj = 4 To 10
                txtAux(jj).Height = DataGrid3.RowHeight - 10
                txtAux(jj).Top = alto + 5
                txtAux(jj).visible = b
                txtAux(jj).Enabled = b
            Next jj
            
'            txtAux(4).Enabled = False
            
            Text2(5).Height = DataGrid3.RowHeight - 10
            Text2(5).Top = alto + 5
            Text2(5).visible = b
           
            For jj = 0 To btnBuscar.Count - 1
                btnBuscar(jj).Height = DataGrid3.RowHeight - 10
                btnBuscar(jj).Top = alto + 5
                btnBuscar(jj).visible = b
            Next jj
            
            For jj = 0 To 2
'                Combo1(jj).Height = DataGrid3.RowHeight - 10
                Combo1(jj).Top = alto + 5
                Combo1(jj).visible = b
                Combo1(jj).Enabled = b
            Next jj
    End Select
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Facturas (scafac)
' y los registros correspondientes de las tablas cab. albaranes (scafac1)
' y las lineas de la factura (slifac)
Dim Cad As String
'Dim vTipoMov As CTiposMov

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
'    'solo se puede modificar la factura si no esta contabilizada
'    If FactContabilizada Then Exit Sub
    
    Cad = "Cabecera de Pesada en Báscula." & vbCrLf
    Cad = Cad & "-------------------------------------      " & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar la Pesada:            "
    Cad = Cad & vbCrLf & "Nº Pesada:  " & Format(Text1(0).Text, "0000000")
    Cad = Cad & vbCrLf & "Fecha:  " & Format(Text1(1).Text, "dd/mm/yyyy")

    Cad = Cad & vbCrLf & vbCrLf & " ¿Desea Eliminarla? "

    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
'        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
'        NumPedElim = Data1.Recordset.Fields(1).Value
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
        End If
        
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEliminar:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminar Pesada", Err.Description
End Sub

Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim Cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        DataGrid3.Enabled = True
        If Not Data1.Recordset.EOF Then _
            Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        Cad = Data1.Recordset.Fields(0) & "|"
        Cad = Cad & Data1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(Cad)
        Unload Me
    End If
End Sub

Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbLightBlue
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Combo1(Index).BackColor = vbLightBlue Then Combo1(Index).BackColor = vbWhite
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid3_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Adoaux(1).Recordset.EOF And ModificaLineas <> 1 Then
        If Not IsNull(Adoaux(1).Recordset.Fields(0).Value) Then
            Text2(6).Text = DevuelveDesdeBDNew(cAgro, "rsocios", "nomsocio", "codsocio", Adoaux(1).Recordset!Codsocio, "N")
            Text2(0).Text = DevuelveDesdeBDNew(cAgro, "rcapataz", "nomcapat", "codcapat", Adoaux(1).Recordset!codcapat, "N")
            Text2(8).Text = DevuelveDesdeBDNew(cAgro, "rtarifatra", "nomtarif", "codtarif", Adoaux(1).Recordset!Codtarif, "N")
            PonerDatosCampo CStr(Adoaux(1).Recordset!codCampo)
        End If
    Else
        Text2(6).Text = ""
        Text2(0).Text = ""
        Text2(8).Text = ""
        
        Text2(4).Text = ""
        Text2(2).Text = ""
        Text3(3).Text = ""
        Text4(3).Text = ""
        Text5(3).Text = ""
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim Sql As String

    'Icono del formulario
    Me.Icon = frmPpal.Icon

     'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next kCampo

    ' ICONITOS DE LA BARRA
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
        'el 10 i el 11 son separadors
'        .Buttons(11).Image = 26    'tarar tractor
'        .Buttons(12).Image = 24  'paletizacion
        .Buttons(8).Image = 10  'Imprimir
    End With
    
    With Me.Toolbar2
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26 'Notas de Campo
    End With
    
    ' desplazamiento
    With Me.ToolbarDes
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 6
        .Buttons(2).Image = 7
        .Buttons(3).Image = 8
        .Buttons(4).Image = 9
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 12
    End With
    
    ' ******* si n'hi han llínies *******
    'ICONETS DE LES BARRES ALS TABS DE LLÍNIA
    For kCampo = 1 To 1
        With Me.ToolAux(kCampo)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
            .Buttons(4).Image = 11  'Salir de las notas de campo
        End With
    Next kCampo
   ' ***********************************
    
    For i = 0 To imgAyuda.Count - 1
        imgAyuda(i).Picture = frmPpal.ImageListB.ListImages(10).Picture
    Next i
    
    For i = 0 To 4
        Me.Label15(i).Caption = ""
        Me.Label19(i).Caption = ""
    Next i
    
    ' cargamos los labels de parametros
    If vParamAplic.TipoCaja1 <> "" Then
        Me.Label15(0).Caption = vParamAplic.TipoCaja1
        Me.Label19(0).Caption = "x  " & vParamAplic.PesoCaja1
    End If
    If vParamAplic.TipoCaja2 <> "" Then
        Me.Label15(1).Caption = vParamAplic.TipoCaja2
        Me.Label19(1).Caption = "x  " & vParamAplic.PesoCaja2
    End If
    If vParamAplic.TipoCaja3 <> "" Then
        Me.Label15(2).Caption = vParamAplic.TipoCaja3
        Me.Label19(2).Caption = "x  " & vParamAplic.PesoCaja3
    End If
    If vParamAplic.TipoCaja4 <> "" Then
        Me.Label15(3).Caption = vParamAplic.TipoCaja4
        Me.Label19(3).Caption = "x  " & vParamAplic.PesoCaja4
    End If
    If vParamAplic.TipoCaja5 <> "" Then
        Me.Label15(4).Caption = vParamAplic.TipoCaja5
        Me.Label19(4).Caption = "x  " & vParamAplic.PesoCaja5
    End If
    
    Me.Label19(5).Caption = "x  " & vParamAplic.PesoCajaLLena
    
    LimpiarCampos   'Limpia los campos TextBox
    CargaCombo

    CodTipoMov = "PES"
    VieneDeBuscar = False
    
    '## A mano
    NombreTabla = "rpesadas"
    NomTablaLineas = "rentradas" 'Tabla notas de entrada
    Ordenacion = " ORDER BY rpesadas.nropesada"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    CadenaConsulta = "select * from rpesadas "
    If Albaran <> "" Then
        CadenaConsulta = CadenaConsulta & " where nropesada = " & Albaran
    Else
        CadenaConsulta = CadenaConsulta & " where nropesada = -1"
    End If
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    ' borramos de la tabla temporal
    Sql = "delete from tmppesada where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute Sql

    'cargamos la primera parte de la cadena xml
    v_cadena = "<?xml version=" & """1.0""" & " standalone=" & """yes""" & " ?><DATAPACKET "
    v_cadena = v_cadena & "Version=" & """1.0""" & "><METADATA><FIELDS>"
    v_cadena = v_cadena & "<FIELD attrname=" & """notacamp""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """fechaent""" & " fieldtype=" & """dateTime""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """codprodu""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """codvarie""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """codsocio""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """codcampo""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """kilosbru""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """kilosnet""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """numcajo1""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """numcajo2""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """numcajo3""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """numcajo4""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """numcajo5""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """matricul""" & " fieldtype=" & """string""" & " WIDTH=" & """10""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """codcapat""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """identifi""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """altura""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "<FIELD attrname=" & """zona""" & " fieldtype=" & """i4""" & " />"
    v_cadena = v_cadena & "</FIELDS></METADATA><ROWDATA>"

'        CargaGrid DataGrid1, Data2, False
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
'    SSTab1.Tab = 0
    
'    If DatosADevolverBusqueda <> "" Then
'        Text1(0).Text = DatosADevolverBusqueda
'        HacerBusqueda
'    Else
'        PonerModo 0
'    End If
    
    If DatosADevolverBusqueda = "" Then
        If Albaran = "" Then
            PonerModo 0
        Else
            HacerBusqueda
'            SSTab1.Tab = 0
        End If
    Else
        BotonBuscar
    End If
    
    PrimeraVez = False
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
'    Me.Combo1(0).ListIndex = -1
'    Me.Check1(0).Value = 0
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Modo = 5 Then
        Cancel = 1
        MsgBox "Debe salir previamente de las Notas de Campo.", vbExclamation
        Exit Sub
    Else
        Cancel = 0
    End If
    
    If Modo = 4 Then TerminaBloquear
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = Aux
        
        
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    If imgFec(0).Tag < 2 Then
        Text1(CByte(imgFec(0).Tag) + 1).Text = Format(vFecha, "dd/mm/yyyy") '<===
    Else
        Text1(CByte(imgFec(0).Tag) + 8).Text = Format(vFecha, "dd/mm/yyyy") '<===
    End If
    ' ********************************************
End Sub

Private Sub frmCamp_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(7).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo de campo
    FormateaCampo txtAux(7)
    If EstaCampoDeAlta(txtAux(7).Text) Then
        PonerDatosCampo txtAux(7).Text
    Else
        MsgBox "El campo no está dado de alta. Reintroduzca.", vbExclamation
        txtAux(7).Text = ""
        PonerFoco txtAux(7)
    End If

End Sub

Private Sub frmCap_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(0).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo de capataz
    FormateaCampo txtAux(0)
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre de capataz
End Sub

Private Sub frmEntPrev_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
    
    If CadenaSeleccion <> "" Then
        CadB = "nropesada = " & DBSet(RecuperaValor(CadenaSeleccion, 1), "N")
        
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(7).Text = RecuperaValor(CadenaSeleccion, 1)
    PonerDatosCampo txtAux(7)
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(6).Text = RecuperaValor(CadenaSeleccion, 1) 'codsocio
    FormateaCampo txtAux(6)
    Text2(6).Text = RecuperaValor(CadenaSeleccion, 2) 'nomsocio
End Sub

Private Sub frmTar_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(8).Text = RecuperaValor(CadenaSeleccion, 1) ' codigo tarifa
    FormateaCampo txtAux(8)
    Text2(8).Text = RecuperaValor(CadenaSeleccion, 2) ' nombre tarifa
End Sub

Private Sub frmTrans_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de transportistas
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Codigo de trnsportista
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(5).Text = RecuperaValor(CadenaSeleccion, 1)  'Codigo de variedad
    Text2(5).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
End Sub

Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           CargaCadenaAyuda vCadena, Index
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
    
End Sub

Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. de Transportista
            indice = 3
            Set frmTrans = New frmManTranspor
            frmTrans.DeConsulta = True
            frmTrans.DatosADevolverBusqueda = "0|1|"
            frmTrans.CodigoActual = Text1(3).Text
            frmTrans.Show vbModal
            Set frmTrans = Nothing
            PonerFoco Text1(3)
    End Select
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgFec_Click(Index As Integer)
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

    If Index < 2 Then
        imgFec(0).Tag = Index '<===
        ' *** repasar si el camp es txtAux o Text1 ***
        If Text1(Index + 1).Text <> "" Then frmC.NovaData = Text1(Index + 1).Text
    Else
        imgFec(0).Tag = Index '<===
        ' *** repasar si el camp es txtAux o Text1 ***
        If Text1(Index + 8).Text <> "" Then frmC.NovaData = Text1(Index + 8).Text
    End If
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    If Index < 2 Then
        PonerFoco Text1(CByte(imgFec(0).Tag) + 1) '<===
    Else
        PonerFoco Text1(CByte(imgFec(0).Tag) + 8) '<===
    End If
    ' ********************************************
End Sub


Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        indice = 2
        frmZ.pTitulo = "Observaciones del Albarán"
        frmZ.pValor = Text1(indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco Text1(indice)
    End If
End Sub


Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
'         BotonEliminarLinea
    Else   'Eliminar Pedido
         BotonEliminar
    End If
End Sub

Private Sub mnNotaCampo_Click()
    If BLOQUEADesdeFormulario(Me) Then
        BotonNotaCampo
    End If
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
        'bloquea la tabla cabecera de albaranes: scaalb
        If BLOQUEADesdeFormulario(Me) Then
            If BloqueaLineasAlb Then BotonModificarLinea (1)
        End If
         
    Else   'Modificar albaran
        'bloquea la tabla cabecera de albaranes: scaalb
        If BLOQUEADesdeFormulario(Me) Then
            BotonModificar
        End If
    End If
End Sub


Private Function BloqueaLineasAlb() As Boolean
'bloquea todas las lineas de la factura
Dim Sql As String

    On Error GoTo EBloqueaLin

    BloqueaLineasAlb = False
    'bloquear cabecera albaranes
    Sql = "select * FROM slialb "
    Sql = Sql & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute Sql, , adCmdText
    BloqueaLineasAlb = True

EBloqueaLin:
    If Err.Number <> 0 Then BloqueaLineasAlb = False
End Function

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    If (Modo = 5) Then 'Modo 5: Mto Lineas
        '1:Insertar linea, 2: Modificar
        If ModificaLineas = 1 Or ModificaLineas = 2 Then cmdCancelar_Click
        cmdRegresar_Click
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub



Private Sub Text1_Change(Index As Integer)
    If Index = 9 Then HaCambiadoCP = True 'Cod. Postal
End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
'    If Index = 9 Then HaCambiadoCP = False 'CPostal
'    If Index = 1 And Modo = 1 Then
'        SendKeys "{tab}"
'        Exit Sub
'    End If
    ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 1: KEYFecha KeyAscii, 0 ' fecha
                Case 3: KEYBusqueda KeyAscii, 0 'transportista
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
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
Dim devuelve As String
Dim cadMen As String
Dim Sql As String
Dim Nregs As Long
Dim Tara As String
        
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1 'Fecha pesada
            '[Monica]28/08/2013: comprobamos que la fecha esté en la campaña
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index), True
            
        Case 3 ' Transportista
            If Text1(Index).Text <> "" Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "rtransporte", "nomtrans")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Transportista: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmTrans = New frmManTranspor
                        frmTrans.DatosADevolverBusqueda = "0|1|"
                        frmTrans.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmTrans.Show vbModal
                        Set frmTrans = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                Else
                    If Modo = 3 Then
                        Tara = DevuelveDesdeBDNew(cAgro, "rtransporte", "taravehi", "codtrans", Text1(3), "T")
                        If Tara <> "" Then
                            Text1(11).Text = Format(Tara, "###,##0")
                            CalcularTaras
                        End If
                    End If
                
                
                End If
            Else
            
                Text2(Index).Text = ""
            End If
        
            
        Case 22, 13, 14, 15, 16, 17, 21, 2, 11 'pesos
            If Modo = 1 Then Exit Sub
            If PonerFormatoEntero(Text1(Index)) Then
                CalcularTaras
            End If

    End Select
End Sub


Private Sub HacerBusqueda()
Dim CadB As String
Dim cadAux As String
    
'    '--- Laura 12/01/2007
'    cadAux = Text1(5).Text
'    If Text1(4).Text <> "" Then Text1(5).Text = ""
'    '---
    
'    '--- Laura 12/01/2007
'    Text1(5).Text = cadAux
'    '---
    
'--monica
'    CadB = ObtenerBusqueda(Me)
'++monica
    If Albaran = "" Then
        CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    Else
        CadB = "numalbar = " & Albaran & " "
    End If

    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select rpesadas.* from " & NombreTabla
        CadenaConsulta = CadenaConsulta & " WHERE " & CadB & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)

    Set frmEntPrev = New frmEntPesadaPrev
    frmEntPrev.cWhere = CadB
    frmEntPrev.DatosADevolverBusqueda = "0|1|2|"
    frmEntPrev.Show vbModal
    
    Set frmEntPrev = Nothing

End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass

    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then
            PonerFoco Text1(kCampo)
'            Text1(0).BackColor = vblightblue
        End If
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        '--monica
        'LLamaLineas Modo, 0, "DataGrid2"
        PonerCampos
    End If


    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCamposLineas()
'Carga el grid de los AlbaranesxFactura, es decir, la tabla scafac1 de la factura seleccionada
Dim b As Boolean
Dim b2 As Boolean
Dim i As Integer

    On Error GoTo EPonerLineas

    If Data1.Recordset.EOF Then Exit Sub

    Screen.MousePointer = vbHourglass
    
    If Data1.Recordset.RecordCount > 0 Then
        CargaGrid DataGrid3, Adoaux(1), True
    Else
        CargaGrid DataGrid3, Adoaux(1), False
    End If
    If Not Adoaux(1).Recordset.EOF Then
        Text2(6).Text = DevuelveDesdeBDNew(cAgro, "rsocios", "nomsocio", "codsocio", Adoaux(1).Recordset!Codsocio, "N")
        Text2(0).Text = DevuelveDesdeBDNew(cAgro, "rcapataz", "nomcapat", "codcapat", Adoaux(1).Recordset!codcapat, "N")
        Text2(8).Text = DevuelveDesdeBDNew(cAgro, "rtarifatra", "nomtarif", "codtarif", Adoaux(1).Recordset!Codtarif, "N")
        PonerDatosCampo CStr(Adoaux(1).Recordset!codCampo)
    Else
        Text2(6).Text = ""
        Text2(0).Text = ""
        Text2(8).Text = ""
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim BrutoFac As Single
Dim b As Boolean

    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    
'    b = PonerCamposForma2(Me, Data1, 2, "FrameDatosPesosTaras")
    b = PonerCamposForma(Me, Data1)
    'poner descripcion campos
    Modo = 4
    
    PosarDescripcions
    
    Modo = 2
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario

    If Err.Number <> 0 Then Err.Clear
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim i As Byte, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar
    '## No tiene el boton modificar y no utiliza la funcion general
'    ActualizarToolbar Modo, Kmodo
    BuscaChekc = ""
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    If Modo = 5 And ModificaLineas = 0 Then
        lblIndicador.Caption = ""
    End If

    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Or Albaran <> "" Then
        cmdRegresar.visible = b
    Else
        cmdRegresar.visible = False
    End If
        
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible b And Data1.Recordset.RecordCount > 1
          
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    
'    FrameDatosNota.Enabled = False
    Frame4.Enabled = False
    
    b = (Modo <> 1)
    'Campos Nº Albarán bloqueado y en azul
    BloquearTxt Text1(0), b, True
    
    b = (Modo <> 1) And (Modo <> 3) And (Modo <> 4)
    BloquearTxt Text1(1), b 'fechapesada
    BloquearTxt Text1(3), b 'transportista
'    BloquearCmb Me.Combo1(0), (Modo <> 1)
    
'    BloquearChk Me.Check1(0), (Modo = 0 Or Modo = 2)
    
'    Me.imgZoom(0).Enabled = Not (Modo = 0)
    
    'Si no es modo lineas Boquear los TxtAux
    For i = 1 To 1
        txtAux(i).visible = False
        BloquearTxt txtAux(i), True
    Next i
    For i = 3 To 10
        txtAux(i).visible = False
        BloquearTxt txtAux(i), True
    Next i
    
    For i = 5 To 5
        Text2(i).visible = ((Modo = 5) And (indFrame = 1))
        Text2(i).Enabled = False
    Next i
    
    For i = 0 To 4
        BloquearBtn Me.btnBuscar(i), (ModificaLineas = 0)
    Next i
    
    For i = 0 To 2
        Combo1(i).visible = False '(Modo = 5)
        Combo1(i).Enabled = False
    Next i
    
    '---------------------------------------------
'    b = (Modo <> 0 And Modo <> 2) Or (Modo = 5 And ModificaLineas <> 0)
    b = (Modo = 1) Or Modo = 3 Or Modo = 4 Or (Modo = 5 And ModificaLineas <> 0)
    cmdCancelar.visible = b
    CmdAceptar.visible = b
    
    BloquearImgBuscar Me, Modo, ModificaLineas
    BloquearImgFec Me, 0, Modo
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    ' *** si n'hi han llínies i algún tab que no te datagrid ***
    BloquearFrameAux Me, "FrameAux1", Modo, 1
    
'    'Campos Nº entrada bloqueado y en azul
'    BloquearTxt Text1(0), b, True
    
    'taras desbloqueadas unicamente para buscar
    For i = 18 To 20
        BloquearTxt Text1(i), Not (Modo = 1)
    Next i
    For i = 9 To 10
        BloquearTxt Text1(i), Not (Modo = 1)
    Next i
    For i = 23 To 23
        BloquearTxt Text1(i), Not (Modo = 1)
    Next i
    
    Me.ToolAux(1).Enabled = (Modo = 5)
    Frame2.Enabled = Not (Modo = 5)
'    FrameDatosPesosTaras.Enabled = Not (Modo = 5)
    Toolbar1.Enabled = Not (Modo = 5)
    
    PonerTarasVisibles
    
        
    ' ***************************
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim b As Boolean
Dim Serie As String
Dim Sql As String

    On Error GoTo EDatosOK

    DatosOk = False
    
'    ComprobarDatosTotales

    'comprobamos datos OK de la tabla scaalb
    b = CompForm2(Me, 2, "Frame2") ' , 1) 'Comprobar formato datos ok de la cabecera: opcion=1
    If Not b Then Exit Function
    
        
    '[Monica]29/11/2017: comprobamos recolectado por y transportado por
    '                    de momento solo para picassent, deberia generalizarlo
    If b Then
        If vParamAplic.Cooperativa = 2 Then
            Sql = "select count(*) from rentradas where nropesada = " & DBSet(Text1(0).Text, "N")
            If TotalRegistros(Sql) <> 0 Then
                If ExistenNotasSinTransportista(Text1(0).Text) And CLng(ComprobarCero(Text1(3).Text)) = 0 Then
                    If MsgBox("Si la pesada está transportada por el socio, no deben existir notas transportadas por la cooperativa. " & vbCrLf & vbCrLf & "¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then b = False
                End If
                If ExistenNotasSinTransportista2(Text1(0).Text) And CLng(ComprobarCero(Text1(3).Text)) <> 0 Then
                    If MsgBox("Si la pesada está transportada por la cooperativa, no deben existir notas transportadas por el socio. " & vbCrLf & vbCrLf & "¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then b = False
                    b = False
                End If
            End If
        End If
    End If
    
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function ExistenNotasSinTransportista(Pesada As String) As Boolean
Dim Sql As String

    Sql = "select count(*) from rentradas where nropesada = " & DBSet(Pesada, "N")
    Sql = Sql & " and (codtrans ='0' or codtrans is null)"
    
    ExistenNotasSinTransportista = (TotalRegistros(Sql) <> 0)

End Function


Private Function ExistenNotasSinTransportista2(Pesada As String) As Boolean
Dim Sql As String

    Sql = "select count(*) from rentradas where nropesada = " & DBSet(Pesada, "N")
    Sql = Sql & " and (transportadopor=1)"
    
    ExistenNotasSinTransportista2 = (TotalRegistros(Sql) <> 0)

End Function


Private Function DatosOkLinea() As Boolean
Dim b As Boolean
Dim i As Byte

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True

    For i = 4 To 7
        If i = 4 Or i = 6 Or i = 7 Then
            If txtAux(i).Text = "" Then
                MsgBox "El campo " & txtAux(i).Tag & " no puede ser nulo", vbExclamation
                b = False
                PonerFoco txtAux(i)
                Exit Function
            End If
        End If
    Next i
            
    DatosOkLinea = b
    
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)

    If BloqueaRegistro(NombreTabla, "nropesada = " & Data1.Recordset!nropesada) Then
'    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
        Select Case Button.Index
            Case 1
                BotonAnyadirLinea Index
            Case 2
                BotonModificarLinea Index
            Case 3
                BotonEliminarLinea Index
            Case 4 'salimos
                BotonSalirNotas
            Case Else
        End Select
    End If

End Sub


Private Sub BotonEliminarLinea(Index As Integer)
Dim Cad As String
Dim Sql As String
Dim Mens As String
Dim b As Boolean

    On Error GoTo EEliminarLinea

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    ' ***************************************************************************

    b = True

    ' *************** canviar la pregunta ****************
    Cad = "¿Seguro que desea eliminar la Nota de la Pesada?"
    Cad = Cad & vbCrLf & "Pesada: " & Adoaux(1).Recordset.Fields(0)
    Cad = Cad & vbCrLf & "Nota: " & Adoaux(1).Recordset.Fields(2)
    
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminarLinea
        Screen.MousePointer = vbHourglass
        NumRegElim = Adoaux(1).Recordset.AbsolutePosition
        
        If Not EliminarLinea Then
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
'            CalcularDatosAlbaran
            If SituarDataTrasEliminar(Adoaux(1), NumRegElim) Then
                PonerCampos
            Else
                PonerCampos
'                        LimpiarCampos
'                        PonerModo 0
            End If
            PonerModo 5
        End If
    End If
    Screen.MousePointer = vbDefault
       
EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Or Not b Then MuestraError Err.Number, "Eliminar Linea de Pesada", Err.Description

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 5  'Buscar
            mnBuscar_Click
        Case 6  'Todos
            BotonVerTodos
        
        Case 1  'Añadir
            mnNuevo_Click

        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
    End Select
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub ActualizarToolbar(Modo As Byte, Kmodo As Byte)
'Modo: Modo antiguo
'Kmodo: Modo que se va a poner

    If (Modo = 5) And (Kmodo <> 5) Then
        'El modo antigu era modificando las lineas
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
'        Toolbar1.Buttons(5).Image = 3
'        Toolbar1.Buttons(5).ToolTipText = "Nuevo Albaran"
        '-- Modificar
        Toolbar1.Buttons(5).Image = 4
        Toolbar1.Buttons(5).ToolTipText = "Modificar Factura"
        '-- eliminar
        Toolbar1.Buttons(6).Image = 5
        Toolbar1.Buttons(6).ToolTipText = "Eliminar Factura"
    End If
    If Kmodo = 5 Then
        'Ponemos nuevos dibujitos y tal y tal
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
'        Toolbar1.Buttons(5).Image = 12
'        Toolbar1.Buttons(5).ToolTipText = "Nueva linea"
        '-- Modificar
        Toolbar1.Buttons(5).Image = 13
        Toolbar1.Buttons(5).ToolTipText = "Modificar linea factura"
        '-- eliminar
        Toolbar1.Buttons(6).Image = 14
        Toolbar1.Buttons(6).ToolTipText = "Eliminar linea factura"
    End If
End Sub
    


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim b As Boolean
Dim Opcion As Byte
Dim Sql As String

    On Error GoTo ECargaGrid

    b = DataGrid3.Enabled
    Select Case vDataGrid.Name
        Case "DataGrid3" 'notas de entrada
            Opcion = 1
    End Select
    
    Sql = MontaSQLCarga(enlaza, Opcion)
    CargaGridGnral vDataGrid, vData, Sql, PrimeraVez
    
    vDataGrid.RowHeight = 270
    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
    
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
    vDataGrid.Enabled = Not b
    
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim tots As String
Dim i As Integer

    On Error GoTo ECargaGrid

    Select Case vDataGrid.Name
         Case "DataGrid3" 'rentradas
            tots = "N||||0|;N||||0|;S|txtAux(4)|T|Nota|900|;"
            tots = tots & "S|txtAux(5)|T|Código|1000|;S|btnBuscar(0)|B|||;"
            tots = tots & "S|Text2(5)|T|Variedad|2400|;S|txtAux(6)|T|Socio|900|;S|btnBuscar(1)|B|||;"
            tots = tots & "S|txtAux(7)|T|Campo|1200|;S|btnBuscar(2)|B|||;S|txtAux(0)|T|Capataz|900|;S|btnBuscar(3)|B|||;S|txtAux(8)|T|Tarifa|700|;"
            tots = tots & "S|btnBuscar(4)|B|||;N||||0|;S|Combo1(0)|C|Tipo Entr.|900|;N||||0|;S|Combo1(1)|C|Recolect.|900|;N||||0|;S|Combo1(2)|C|Transpor.|900|;"
            tots = tots & "S|txtAux(9)|T|Cajas|900|;S|txtAux(10)|T|Peso Neto|1100|;N||||0|;N||||0|;"
            
            arregla tots, DataGrid3, Me, 350
            
    End Select
    
    For i = 2 To 8
        DataGrid3.Columns(i).Alignment = dbgLeft
    Next i
    
    vDataGrid.HoldFields
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Buscar
            mnNotaCampo_Click
    End Select
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub


Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 5 Then
            Select Case Index
                Case 5: KEYImage KeyAscii, 0 'variedad
                Case 6: KEYImage KeyAscii, 1 'socio
                Case 7: KEYImage KeyAscii, 2 'campo
                Case 0: KEYImage KeyAscii, 3 'capataz
                Case 8: KEYImage KeyAscii, 4 'tarifa
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub


Private Sub KEYImage(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    btnBuscar_Click (indice)
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Sql As String
Dim devuelve As String
Dim b As Boolean
Dim TipoDto As Byte


    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 6
            If PonerFormatoEntero(txtAux(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(txtAux(Index), "rsocios", "nomsocio")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Socio: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmSoc = New frmManSocios
                        frmSoc.DatosADevolverBusqueda = "0|1|"
'                        frmVar.NuevoCodigo = Text1(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmSoc.Show vbModal
                        Set frmSoc = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                Else
                    If EstaSocioDeAlta(txtAux(Index)) Then
                        PonerCamposSocioVariedad
                    Else
                        MsgBox "El socio está dado de baja. Reintroduzca.", vbExclamation
                        txtAux(Index).Text = ""
                        PonerFoco txtAux(Index)
                    End If
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        Case 5 'VARIEDAD
            If PonerFormatoEntero(txtAux(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(txtAux(Index), "variedades", "nomvarie")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Variedad: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmVar = New frmComVar
                        frmVar.DatosADevolverBusqueda = "0|1|"
'                        frmVar.NuevoCodigo = Text1(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmVar.Show vbModal
                        Set frmVar = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                Else
                    If (Modo = 5) And EsVariedadGrupo6(txtAux(Index).Text) Then
                        MsgBox "Esta variedad es del Grupo de Bodega. Revise.", vbExclamation
                        PonerFoco txtAux(Index)
                    Else
                        PonerCamposSocioVariedad
                    End If
                End If
            Else
                Text2(Index).Text = ""
            End If
        
                
        Case 7 'codigo de campo
            If Modo = 1 Then Exit Sub
            If PonerFormatoEntero(txtAux(Index)) Then
                Sql = ""
                Sql = DevuelveDesdeBDNew(cAgro, "rcampos", "codcampo", "codcampo", txtAux(Index).Text, "N")
                If Sql = "" Then
                    cadMen = "No existe el Campo: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCamp = New frmManCampos
                        frmCamp.DatosADevolverBusqueda = "0|1|"
'                        frmCamp.NuevoCodigo = Text1(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmCamp.Show vbModal
                        Set frmCamp = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                Else
                    If Not EstaCampoDeAlta(txtAux(Index).Text) Then
                        MsgBox "El campo no está dado de alta. Reintroduzca.", vbExclamation
                        txtAux(Index).Text = ""
                        PonerFoco txtAux(Index)
                    Else
                        '[Monica]13/08/2018: no se permiten entradas de campos de tratamiento
                        If EsCampoDeTratamiento(txtAux(Index).Text) Then
                            MsgBox "El campo es de tratamiento. Reintroduzca.", vbExclamation
                            txtAux(Index).Text = ""
                            PonerFoco txtAux(Index)
                        Else
                            If Not EsCampoSocioVariedad(txtAux(Index).Text, txtAux(6).Text, txtAux(5).Text) Then
                                MsgBox "El campo no es del Socio Variedad. Reintroduzca.", vbExclamation
                                PonerFoco txtAux(Index)
                            Else
                                PonerDatosCampo (txtAux(Index))
                                If ModificaLineas = 1 Then
                                    Combo1(1).ListIndex = DevuelveValor("select recolect from rcampos where codcampo = " & DBSet(txtAux(7).Text, "N"))
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            
        Case 8 'tarifa de transporte
            If PonerFormatoEntero(txtAux(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(txtAux(Index), "rtarifatra", "nomtarif")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Tarifa de Transporte: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmTar = New frmManTarTra
                        frmTar.DatosADevolverBusqueda = "0|1|"
                        frmTar.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmTar.Show vbModal
                        Set frmTar = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
                
        Case 0 'capataz
            If PonerFormatoEntero(txtAux(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(txtAux(Index), "rcapataz", "nomcapat")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Capataz: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCap = New frmManCapataz
                        frmCap.DatosADevolverBusqueda = "0|1|"
                        frmCap.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmCap.Show vbModal
                        Set frmCap = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 9 'Cajas
            If PonerFormatoEntero(txtAux(Index)) Then
                CalcularPesoNeto
                cmdAceptar_Click
            End If
    End Select
    
End Sub




Private Function Eliminar() As Boolean
Dim Sql As String, LEtra As String, Sql2 As String
Dim Rs As ADODB.Recordset
Dim SQL1 As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String
Dim NumF As Long
    On Error GoTo FinEliminar

    If Data1.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    b = True

    If b Then
        'Eliminar en tablas de cabecera de albaran
        '------------------------------------------
        
        
        Sql = " " & ObtenerWhereCP(True)
        ' cada una de las lineas se insertan en la vtempo
        SQL1 = "select * from rentradas  " & Sql
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs.EOF And b
            ' insertamos en el chivato
            NumF = SugerirCodigoSiguienteStr("chivato", "numorden")
            
            SQL1 = "insert into chivato (numorden, basedato, nomtabla, operacio, fechadia, separado,"
            SQL1 = SQL1 & "claveant, clavenue, nombmemo, nombmem1, nombmem2, horaproc, nombmem3, nombmem4) values ("
            SQL1 = SQL1 & DBSet(NumF, "N") & ","
            SQL1 = SQL1 & "'agro','sentba','D'," & DBSet(Now, "F") & ",'&'," & DBLet(Adoaux(1).Recordset.Fields(2).Value, "N") & ","
            SQL1 = SQL1 & ValorNulo & ","
            SQL1 = SQL1 & ValorNulo & ","
            SQL1 = SQL1 & ValorNulo & ","
            SQL1 = SQL1 & ValorNulo & ","
            SQL1 = SQL1 & "'" & Format(Now, "hh:mm:ss") & "',"
            SQL1 = SQL1 & ValorNulo & ","
            SQL1 = SQL1 & ValorNulo & ")"
            
            
            
'[Monica]03/12/2009 anterior tabla
'
'            SQL1 = " insert into chivato (numlinea, basedato, tabla, operacio, fechadia, separado,"
'            SQL1 = SQL1 & "claveant, clavenue, xml) Values ("
'            SQL1 = SQL1 & DBSet(NumF, "N") & ","
'            SQL1 = SQL1 & "'agro','sentba','D'," & DBSet(Now, "F") & ",'&'," & DBLet(Adoaux(1).Recordset.Fields(2), "N") & ","
'            SQL1 = SQL1 & ValorNulo & ","
'            SQL1 = SQL1 & ValorNulo & ")"
        
            conn.Execute SQL1
        
            Mens = "Insertando en temporal "
            b = InsertarTemporal("Z", Mens)
        
            Rs.MoveNext
        Wend
        
        Set Rs = Nothing
        
        If b Then
            'Lineas de envases (slialb)
            conn.Execute "Delete from rentradas " & Sql
            
            'Cabecera de factura
            conn.Execute "Delete from " & NombreTabla & Sql
            
            'Decrementar contador si borramos el ult. palet
            Set vTipoMov = New CTiposMov
            vTipoMov.DevolverContador CodTipoMov, Val(Text1(0).Text)
            Set vTipoMov = Nothing
        End If
    End If
FinEliminar:
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, "Eliminar Pesada Báscula", Err.Description & " " & Mens
        b = False
    End If
    If Not b Then
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function

Private Function EliminarLinea() As Boolean
Dim Sql As String, LEtra As String, SQL1 As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String
Dim Linea As Long

    On Error GoTo FinEliminar

    b = False
    If Adoaux(1).Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    Mens = ""
    
    b = False
    
    'Eliminar en tablas de slialb
    '------------------------------------------
    Sql = " where nropesada = " & Adoaux(1).Recordset.Fields(0)
    Sql = Sql & " and numlinea = " & Adoaux(1).Recordset.Fields(1)

    ' Insertamos en el chivato
    Linea = SugerirCodigoSiguienteStr("chivato", "numorden")
    
    SQL1 = "insert into chivato (numorden, basedato, nomtabla, operacio, fechadia, separado,"
    SQL1 = SQL1 & "claveant, clavenue, nombmemo, nombmem1, nombmem2, horaproc, nombmem3, nombmem4) values ("
    SQL1 = SQL1 & DBSet(Linea, "N") & ","
    SQL1 = SQL1 & "'agro','sentba','D',"
    SQL1 = SQL1 & DBSet(Now, "F") & ",'&',"
    SQL1 = SQL1 & DBLet(Adoaux(1).Recordset.Fields(2), "N") & ","
    SQL1 = SQL1 & ValorNulo & ","
    SQL1 = SQL1 & ValorNulo & ","
    SQL1 = SQL1 & ValorNulo & ","
    SQL1 = SQL1 & ValorNulo & ","
    SQL1 = SQL1 & "'" & Format(Now, "hh:mm:ss") & "',"
    SQL1 = SQL1 & ValorNulo & ","
    SQL1 = SQL1 & ValorNulo & ")"
    
    
'[Monica]03/12/2009
'    SQL1 = "insert into chivato (numlinea, basedato, tabla, operacio, fechadia, separado,"
'    SQL1 = SQL1 & "claveant, clavenue, xml) values ("
'    SQL1 = SQL1 & DBSet(Linea, "N") & ","
'    SQL1 = SQL1 & "'agro','sentba','D',"
'    SQL1 = SQL1 & DBSet(Now, "FH") & ",'&',"
'    SQL1 = SQL1 & DBLet(Adoaux(1).Recordset.Fields(2), "N") & ","
'    SQL1 = SQL1 & ValorNulo & ","
'    SQL1 = SQL1 & ValorNulo & ")"
    
    conn.Execute SQL1
    ' Insertamos en la tabla temporal
    Mens = "Insertando en temporal "
    b = InsertarTemporal("Z", Mens)

    If b Then
    'Lineas de variedades
        conn.Execute "Delete from rentradas " & Sql
    End If
    
    
FinEliminar:
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, "Eliminar Entrada de la Pesada ", Err.Description & " " & Mens
        b = False
    End If
    If Not b Then
        conn.RollbackTrans
        EliminarLinea = False
    Else
        conn.CommitTrans
        EliminarLinea = True
    End If
End Function

Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next

    CargaGrid DataGrid3, Me.Adoaux(1), False 'nro de notas
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         If SituarDataMULTI(Data1, vWhere, Indicador) Then
             PonerModo 2
             lblIndicador.Caption = Indicador
        Else
             LimpiarCampos
             'Poner los grid sin apuntar a nada
             LimpiarDataGrids
             PonerModo 0
         End If
    Else
        'El Data esta vacio, desde el modo de inicio se pulsa Insertar
        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim Sql As String

    On Error Resume Next
    
    Sql = "nropesada= " & DBSet(Text1(0).Text, "N")
    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function


Private Function MontaSQLCarga(enlaza As Boolean, Opcion As Byte) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim Sql As String
    
    Select Case Opcion
        Case 1  'rentradas
            Sql = "SELECT nropesada, numlinea, numnotac, rentradas.codvarie, variedades.nomvarie, codsocio, codcampo,"
            Sql = Sql & "codcapat, codtarif, tipoentr, CASE tipoentr WHEN 0 THEN ""Normal"" WHEN 1 THEN ""V.Campo"" WHEN 2 THEN ""P.Integ."" WHEN 3 THEN ""Ind.Directo"" END, recolect, CASE recolect WHEN 0 THEN ""Coop"" WHEN 1 THEN ""Socio"" END, "
            Sql = Sql & "transportadopor, CASE transportadopor WHEN 0 THEN ""Coop"" WHEN 1 THEN ""Socio"" END,"
            Sql = Sql & "numcajo1, kilosnet, kilosbru, kilostra "
            Sql = Sql & " FROM rentradas, variedades "
            Sql = Sql & " WHERE rentradas.codvarie = variedades.codvarie "
    End Select
    
    If enlaza Then
        Sql = Sql & " and " & ObtenerWhereCP(False)
    Else
        Sql = Sql & " and nropesada = -1"
    End If
    Sql = Sql & " ORDER BY nropesada, numlinea"
    MontaSQLCarga = Sql
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean, bAux As Boolean
Dim i As Integer

        b = ((Modo = 2) Or (Modo = 0)) And (Albaran = "") 'Or (Modo = 5 And ModificaLineas = 0)
        'Buscar
        Toolbar1.Buttons(5).Enabled = b
        Me.mnBuscar.Enabled = b
        'Vore Tots
        Toolbar1.Buttons(6).Enabled = b
        Me.mnVerTodos.Enabled = b
        'Añadir
        Toolbar1.Buttons(1).Enabled = b
        Me.mnModificar.Enabled = b
        
        
        b = (Modo = 2 And Data1.Recordset.RecordCount > 0) And (Albaran = "")
        'Modificar
        Toolbar1.Buttons(2).Enabled = b
        Me.mnModificar.Enabled = b
        'eliminar
        Toolbar1.Buttons(3).Enabled = b
        Me.mnEliminar.Enabled = b
        'Notas de Campo
        Toolbar2.Buttons(1).Enabled = (Modo = 2) Or (Albaran <> "")
        Me.mnNotaCampo.Enabled = (Modo = 2) Or (Albaran <> "")

    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
'++monica: si insertamos lo he quitado
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    b = (Modo = 5) And (Albaran = "")
    For i = 1 To 1
        ToolAux(i).Buttons(1).Enabled = b ' añadir y salir siempre activos
        ToolAux(i).Buttons(4).Enabled = b
        
        If b Then
            bAux = (b And Me.Adoaux(1).Recordset.RecordCount > 0)
        End If
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i


End Sub




Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim i As Byte
    
    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 0 To Combo1.Count - 1
        Combo1(i).Clear
    Next i
    
    'tipo de entrada
    Combo1(0).AddItem "Normal"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "V.Campo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "P.Integrado"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    Combo1(0).AddItem "Ind.Directo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 3
    Combo1(0).AddItem "Retirada"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 4
    Combo1(0).AddItem "Venta Comercio"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 5
    
    'tipo de recoleccion
    Combo1(1).AddItem "Cooperativa"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Socio"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1

    'transportado por
    Combo1(2).AddItem "Cooperativa"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 0
    Combo1(2).AddItem "Socio"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 1
End Sub

Private Function ModificaCabecera() As Boolean
Dim b As Boolean
Dim MenError As String
Dim Sql As String

    On Error GoTo EModificarCab

    conn.BeginTrans
    
    b = ModificaDesdeFormulario2(Me, 2, "Frame2")
    If b Then
        If FechaAnt <> Text1(1).Text Or TransporAnt <> Text1(3).Text Then
            MenError = "Modificando datos de lineas"
            If b Then
                Sql = "update rentradas set fechaent = " & DBSet(Text1(1).Text, "F")
                Sql = Sql & ", horaentr = concat(concat(" & DBSet(Text1(1).Text, "F") & ",' '), time(horaentr)) " 'Format(Now, "hh:mm:ss"), "FH")
                Sql = Sql & ", codtrans = " & DBSet(Text1(3).Text, "T")
                Sql = Sql & " where nropesada = " & DBSet(Text1(0).Text, "N")
                
                conn.Execute Sql
                
                b = SeHaModificadoCabecera
            End If
        End If
        
    End If
EModificarCab:
    If Err.Number <> 0 Then
        MenError = "Modificando Pesada." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        b = False
    End If
    If b Then
        ModificaCabecera = True
        conn.CommitTrans
    Else
        ModificaCabecera = False
        conn.RollbackTrans
    End If
End Function


Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String

    On Error GoTo EInsertarCab
    
'    CodTipoMov = Text1(6).Text
    
'    If TipoFactura = 0 Then
        Set vTipoMov = New CTiposMov
        If vTipoMov.Leer(CodTipoMov) Then
            Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
            Sql = CadenaInsertarDesdeForm(Me)
            If Sql <> "" Then
                If InsertarOferta(Sql, vTipoMov) Then
                    CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                    PonerCadenaBusqueda
                    PonerModo 2
                    'Ponerse en Modo Insertar Lineas
    '                BotonMtoLineas 0, "Variedades"
                    BotonAnyadirLinea 0
                End If
            End If
            Text1(0).Text = Format(Text1(0).Text, "0000000")
        End If
        Set vTipoMov = Nothing
'    Else
'            Sql = CadenaInsertarDesdeForm(Me)
'            Conn.Execute Sql
'
'            CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
'            PonerCadenaBusqueda
'            PonerModo 2
'            'Ponerse en Modo Insertar Lineas
''                BotonMtoLineas 0, "Variedades"
'            BotonAnyadirLinea 0
'            Text1(0).Text = Format(Text1(0).Text, "0000000")
'
'    End If
    
EInsertarCab:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Function InsertarOferta(vSQL As String, vTipoMov As CTiposMov) As Boolean
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim cambiaSQL As Boolean
Dim devuelve As String

    On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Albaranes
    'para ello vemos si existe una factura con ese contador y si existe la incrementamos
    Do
        devuelve = DevuelveDesdeBDNew(cAgro, NombreTabla, "nropesada", "nropesada", Text1(0), "N")
        If devuelve <> "" Then
            'Ya existe el contador incrementarlo
            Existe = True
            vTipoMov.IncrementarContador (CodTipoMov)
            Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
            cambiaSQL = True
        Else
            Existe = False
        End If
    Loop Until Not Existe
    If cambiaSQL Then vSQL = CadenaInsertarDesdeForm(Me)
    
    
    'Aqui empieza transaccion
    conn.BeginTrans
    MenError = "Error al insertar en la tabla Cabecera de Pesadas (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    MenError = "Error al actualizar el contador de la Pesada."
    vTipoMov.IncrementarContador (CodTipoMov)
    
EInsertarOferta:
        If Err.Number <> 0 Then
            MenError = "Insertando Pesada." & vbCrLf & "----------------------------" & vbCrLf & MenError
            MuestraError Err.Number, MenError, Err.Description
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            InsertarOferta = True
        Else
            conn.RollbackTrans
            InsertarOferta = False
        End If
End Function


'Private Sub CargaForaGrid()
'    If DataGrid2.Columns.Count <= 2 Then Exit Sub
'    ' *** posar als camps de fora del grid el valor de la columna corresponent ***
'    Text3(0) = DataGrid2.Columns(12).Text    'Fecha
'    Text3(1) = DataGrid2.Columns(13).Text    'Matricula
'    Text3(2) = DataGrid2.Columns(14).Text    'Destino
'    Text3(3) = DataGrid2.Columns(15).Text   'Variedad
'    Text3(4) = DataGrid2.Columns(16).Text   'Confeccion
'    ' *** Si fora del grid n'hi han camps de descripció, posar-los valor ***
'    ' **********************************************************************
'End Sub

Private Sub InsertarLinea()
'Inserta registre en les taules de Llínies
Dim nomframe As String
Dim b As Boolean
Dim Mens As String

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
'        Case 0: nomFrame = "FrameAux0" 'variedades
    nomframe = "FrameAux1" 'notas de entrada
    ' ***************************************************************
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        
        If InsertarLineaEnv(txtAux(3).Text) Then
'            CalcularDatosAlbaran
            b = BloqueaRegistro("rpesadas", "nropesada = " & Data1.Recordset!nropesada)
            CargaGrid DataGrid3, Adoaux(1), True
            If b Then BotonAnyadirLinea 1
        End If
    End If
End Sub

Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vtabla As String
Dim anc As Single
Dim i As Integer
    
    ModificaLineas = 1 'Posem Modo Afegir Llínia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModificaLineas = 0 Then Exit Sub
    End If
       
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
'    BloquearTxt Text1(6), True
'    BloquearTxt Text1(1), True
'
    
    ' **************************************************

    ' *** posar el nom del les distintes taules de llínies ***
    vtabla = "rentradas"
    ' ********************************************************
    
    vWhere = ObtenerWhereCab(False)
    
    NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)
    ' ***************************************************************

    AnyadirLinea DataGrid3, Adoaux(1)

    anc = DataGrid3.Top
    If DataGrid3.Row < 0 Then
        anc = anc + 240 '210
    Else
        anc = anc + DataGrid3.RowTop(DataGrid3.Row) + 5
    End If
  
    LLamaLineas ModificaLineas, anc, "DataGrid3"

    LimpiarCamposLin "FrameAux1"
    txtAux(1).Text = Text1(0).Text 'nro de pesada
    txtAux(3).Text = NumF
    PonerFoco txtAux(4)
    For i = 5 To 5
        Text2(i).Text = ""
    Next i
    txtAux(10).Enabled = False
'    txtAux(10).visible = False
    BloquearTxt txtAux(10), True
'    BloquearTxt Text2(16), False
    For i = 0 To btnBuscar.Count - 1
        BloquearBtn Me.btnBuscar(i), False
    Next i
    
    '[Monica]20/09/2010: el tipo de entrada lo ponemos por defecto
    Combo1(0).ListIndex = 0
    
    Combo1(2).ListIndex = 0
    
' ******************************************
End Sub

Private Sub LimpiarCamposLin(FrameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, FrameAux  'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Integer
Dim Cad As String
Dim Sql As String
Dim b As Boolean
Dim Mens As String
    
    On Error GoTo eModificarLinea

    ModificarLinea = 0
    Sql = ""

    ' *** posa els noms del frames, tant si son de grid com si no ***
    nomframe = "FrameAux1" 'notas de entrada
    ' **************************************************************


    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        
        '#### LAURA 15/11/2006
        conn.BeginTrans
        
'        Set vCStock = New CStock
        'iniciamos la clase con los valores anteriores para deshacer lo q insertamos antes
        b = ModificaDesdeFormulario2(Me, 2, "FrameAux1")
        If b Then
            Mens = "Insertando en temporal"
            b = InsertarTemporal("U", Mens)
        End If
            
        ModificaLineas = 0
    Else
        Exit Function
    End If
'
        
eModificarLinea:
    If Err.Number <> 0 Or Not b Then
        MuestraError Err.Number, "Modificar Lineas Factura" & vbCrLf & Err.Description & vbCrLf & Mens
        b = False
    End If
    
    If b Then
        conn.CommitTrans
        ModificarLinea = True
    Else
        conn.RollbackTrans
        ModificarLinea = False
    End If
        
End Function
        

Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim b As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte
Dim Cliente As String

    DatosOkLlin = True
    
    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False

    b = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not b Then Exit Function
    
    'comprobamos que no se haya introducido ya el nro de nota
    If Modo = 5 And ModificaLineas = 1 Then
        Sql = "select count(*) from rentradas where numnotac = " & DBSet(txtAux(4).Text, "N")
        If TotalRegistros(Sql) <> 0 Then
            MsgBox "Este nro de nota ya existe. Revise.", vbExclamation
            b = False
            PonerFoco txtAux(4)
        End If
    End If
    
    '[Monica]29/11/2017: comprobamos recolectado por y transportado por
    '                    de momento solo para picassent, deberia generalizarlo
    If b Then
        If vParamAplic.Cooperativa = 2 Then
            If Combo1(1).ListIndex = 0 And (ComprobarCero(txtAux(0).Text) = 0) Then
                MsgBox "Si la entrada está recolectada por la cooperativa, debe introducir capataz. Revise.", vbExclamation
                b = False
                PonerFoco txtAux(0)
            End If
            
            If Combo1(1).ListIndex = 1 And (ComprobarCero(txtAux(0).Text) <> 0) Then
                MsgBox "Si la entrada está recolectada por el socio, no debe introducir capataz. Revise.", vbExclamation
                b = False
                PonerFoco Text1(12)
            End If
        End If
    End If
    If b Then
        If vParamAplic.Cooperativa = 2 Then
            If Combo1(2).ListIndex = 0 And (ComprobarCero(Text1(3).Text) = 0) Then
                MsgBox "Si la entrada está transportada por la cooperativa, debe introducir transportista. Revise.", vbExclamation
                b = False
                PonerFocoCmb Combo1(2)
            End If
            If Combo1(2).ListIndex = 1 And (ComprobarCero(Text1(3).Text) <> 0) Then
                MsgBox "Si la entrada está transportada por el socio, no debe introducir transportista. Revise.", vbExclamation
                b = False
                PonerFocoCmb Combo1(2)
            End If
        End If
    End If
        
    
    DatosOkLlin = b

EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " nropesada= " & Val(Text1(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function

'' ***** si n'hi han varios nivells de tabs *****
'Private Sub SituarTab(numTab As Integer)
'    On Error Resume Next
'
'    SSTab1.Tab = numTab
'
'    If Err.Number <> 0 Then Err.Clear
'End Sub

' **********************************************
    

Private Function InsertarLineaEnv(NumLinea As String) As Boolean
'Inserta un registro en la tabla de lineas de Albaranes: slialb
Dim Sql As String
Dim vWhere As String
Dim b As Boolean
Dim DentroTRANS As Boolean
Dim Mens As String

    On Error GoTo EInsertarLineaEnv
    
    
    
    InsertarLineaEnv = False
    Sql = ""
    
    'Conseguir el siguiente numero de linea
    vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
'    numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
    
    conn.BeginTrans
    
    
    b = InsertarLineaEntrada
    
    If b Then
        Mens = "Insertando en temporal "
        b = InsertarTemporal("I", Mens)
    End If
    
    
    If b Then
        conn.CommitTrans
        InsertarLineaEnv = True
    Else
        conn.RollbackTrans
        InsertarLineaEnv = False
    End If
    Exit Function
    
EInsertarLineaEnv:
    MuestraError Err.Number, "Insertar Notas de Entrada" & vbCrLf & Err.Description
End Function


Private Sub PonerCamposSocioVariedad()
Dim Cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim Rs As ADODB.Recordset


    If txtAux(6).Text = "" Or txtAux(5).Text = "" Then Exit Sub
    

    Cad = "rcampos.codsocio = " & DBSet(txtAux(6).Text, "N") & " and rcampos.fecbajas is null"
    '[Monica]13/08/2018: no se permiten entradas de campos de tratamiento
    Cad = Cad & " and rcampos.tipocampo <> 3 "
    Cad = Cad & " and rcampos.codvarie = " & DBSet(txtAux(5), "N")
     
    Cad1 = "select count(*) from rcampos where " & Cad
     
    NumRegis = TotalRegistros(Cad1)
    
    If NumRegis = 0 Then Exit Sub
    If NumRegis = 1 Then
        Cad1 = "select codcampo from rcampos where " & Cad
        Set Rs = New ADODB.Recordset
        Rs.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            txtAux(7).Text = DBLet(Rs.Fields(0).Value)
            PonerDatosCampo txtAux(7).Text
        End If
    Else
        Set frmMens = New frmMensajes
        frmMens.cadWHERE = " and " & Cad '"rcampos.codsocio = " & NumCod & " and rcampos.fecbajas is null"
        frmMens.campo = txtAux(7).Text
        frmMens.OpcionMensaje = 6
        frmMens.Show vbModal
        Set frmMens = Nothing
    End If
    
End Sub

Private Sub PonerDatosCampo(campo As String)
Dim Cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim Rs As ADODB.Recordset


    If campo = "" Then Exit Sub
    
'    If Not (Modo = 3 Or Modo = 4) Then Exit Sub

    '[Monica]13/08/2018: no se permiten entradas de campos de tratamiento
    Cad = "rcampos.codcampo = " & DBSet(campo, "N") & " and rcampos.fecbajas is null and rcampos.tipocampo <> 3"
     
    Cad1 = "select rcampos.codparti, rpartida.nomparti, rpartida.codzonas, rzonas.nomzonas, "
    Cad1 = Cad1 & " rpueblos.despobla from rcampos, rpartida, rzonas, rpueblos "
    Cad1 = Cad1 & " where " & Cad
    Cad1 = Cad1 & " and rcampos.codparti = rpartida.codparti "
    Cad1 = Cad1 & " and rpartida.codzonas = rzonas.codzonas "
    Cad1 = Cad1 & " and rpartida.codpobla = rpueblos.codpobla "
     
    Set Rs = New ADODB.Recordset
    Rs.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Text1(5).Text = ""
    Text2(4).Text = ""
    Text2(2).Text = ""
    Text3(3).Text = ""
    Text4(3).Text = ""
    Text5(3).Text = ""
    If Not Rs.EOF Then
        Text1(5).Text = campo
        PonerFormatoEntero Text1(5)
        Text2(4).Text = DBLet(Rs.Fields(0).Value, "N") ' codigo de partida
        If Text2(4).Text <> "" Then Text2(4).Text = Format(Text2(4).Text, "0000")
        Text2(2).Text = DBLet(Rs.Fields(1).Value, "T") ' nombre de partida
        Text3(3).Text = DBLet(Rs.Fields(2).Value, "N") ' codigo de zona
        If Text3(3).Text <> "" Then Text3(3).Text = Format(Text3(3).Text, "0000")
        Text4(3).Text = DBLet(Rs.Fields(3).Value, "T") ' nombre de zona
        Text5(3).Text = DBLet(Rs.Fields(4).Value, "T") ' descripcion de poblacion
    End If
    
    Set Rs = Nothing
    
End Sub

Private Sub PonerTarasVisibles()
    'tara1
    Text1(13).Enabled = (vParamAplic.TipoCaja1 <> "")
    Text1(13).visible = (vParamAplic.TipoCaja1 <> "")
    Text1(18).Enabled = (vParamAplic.TipoCaja1 <> "")
    Text1(18).visible = (vParamAplic.TipoCaja1 <> "")

    'tara2
    Text1(14).Enabled = (vParamAplic.TipoCaja2 <> "")
    Text1(14).visible = (vParamAplic.TipoCaja2 <> "")
    Text1(19).Enabled = (vParamAplic.TipoCaja2 <> "")
    Text1(19).visible = (vParamAplic.TipoCaja2 <> "")
    
    'tara3
    Text1(15).Enabled = (vParamAplic.TipoCaja3 <> "")
    Text1(15).visible = (vParamAplic.TipoCaja3 <> "")
    Text1(20).Enabled = (vParamAplic.TipoCaja3 <> "")
    Text1(20).visible = (vParamAplic.TipoCaja3 <> "")
    
    'tara4
    Text1(16).Enabled = (vParamAplic.TipoCaja4 <> "")
    Text1(16).visible = (vParamAplic.TipoCaja4 <> "")
    Text1(9).Enabled = (vParamAplic.TipoCaja4 <> "")
    Text1(9).visible = (vParamAplic.TipoCaja4 <> "")
    
    'tara5
    Text1(17).Enabled = (vParamAplic.TipoCaja5 <> "")
    Text1(17).visible = (vParamAplic.TipoCaja5 <> "")
    Text1(10).Enabled = (vParamAplic.TipoCaja5 <> "")
    Text1(10).visible = (vParamAplic.TipoCaja5 <> "")
End Sub

Private Sub PosarDescripcions()
Dim NomEmple As String

    On Error GoTo EPosarDescripcions

    Text2(3).Text = PonerNombreDeCod(Text1(3), "rtransporte", "nomtrans", "codtrans", "T")
    
    ComprobarKilosCaja Text1(22).Text, CCur(Text1(12).Text)
    
EPosarDescripcions:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo descripciones", Err.Description
End Sub


Private Sub CalcularTaras()
Dim Tara0 As Currency
Dim Tara1 As Currency
Dim Tara2 As Currency
Dim Tara3 As Currency
Dim Tara4 As Currency
Dim Tara5 As Currency

Dim Tara00 As Currency
Dim Tara11 As Currency
Dim Tara12 As Currency
Dim Tara13 As Currency
Dim Tara14 As Currency
Dim Tara15 As Currency

Dim PesoBruto As Currency
Dim PesoNeto As Currency
Dim PesoTrans As Currency
Dim TaraVehi As Currency
Dim OtrasTaras As Currency

Dim KiloCaja As Currency

    Tara0 = 0
    Tara1 = 0
    Tara2 = 0
    Tara3 = 0
    Tara4 = 0
    Tara5 = 0
    
    Tara00 = 0
    Tara11 = 0
    Tara12 = 0
    Tara13 = 0
    Tara14 = 0
    Tara15 = 0
    
    
    Text1(23).Text = ""
    Text1(18).Text = ""
    Text1(19).Text = ""
    Text1(20).Text = ""
    Text1(9).Text = ""
    Text1(10).Text = ""
    
    'tara 0
    If Text1(22).Text <> "" Then
        Tara0 = Round2(CCur(ImporteSinFormato(Text1(22).Text)) * vParamAplic.PesoCajaLLena, 0)
        Tara00 = Tara0
        Text1(23).Text = Tara0
        PonerFormatoEntero Text1(23)
    End If
    
    'tara 1
    If Text1(13).Text <> "" Then
        Tara1 = Round2(CCur(ImporteSinFormato(Text1(13).Text)) * vParamAplic.PesoCaja1, 0)
        Tara11 = Round2(CCur(ImporteSinFormato(Text1(13).Text)) * vParamAplic.PesoCaja11, 0)
        Text1(18).Text = Tara1
        PonerFormatoEntero Text1(18)
    End If
    'tara 2
    If Text1(14).Text <> "" Then
        Tara2 = Round2(CCur(ImporteSinFormato(Text1(14).Text)) * vParamAplic.PesoCaja2, 0)
        Tara12 = Round2(CCur(ImporteSinFormato(Text1(14).Text)) * vParamAplic.PesoCaja12, 0)
        Text1(19).Text = Tara2
        PonerFormatoEntero Text1(19)
    End If
    'tara 3
    If Text1(15).Text <> "" Then
        Tara3 = Round2(CCur(ImporteSinFormato(Text1(15).Text)) * vParamAplic.PesoCaja3, 0)
        Tara13 = Round2(CCur(ImporteSinFormato(Text1(15).Text)) * vParamAplic.PesoCaja13, 0)
        Text1(20).Text = Tara3
        PonerFormatoEntero Text1(20)
    End If
    'tara 4
    If Text1(16).Text <> "" Then
        Tara4 = Round2(CCur(ImporteSinFormato(Text1(16).Text)) * vParamAplic.PesoCaja4, 0)
        Tara14 = Round2(CCur(ImporteSinFormato(Text1(16).Text)) * vParamAplic.PesoCaja14, 0)
        Text1(9).Text = Tara4
        PonerFormatoEntero Text1(9)
    End If
    'tara 5
    If Text1(17).Text <> "" Then
        Tara5 = Round2(CCur(ImporteSinFormato(Text1(17).Text)) * vParamAplic.PesoCaja5, 0)
        Tara15 = Round2(CCur(ImporteSinFormato(Text1(17).Text)) * vParamAplic.PesoCaja15, 0)
        Text1(10).Text = Tara5
        PonerFormatoEntero Text1(10)
    End If

    'peso neto
    PesoBruto = 0
    TaraVehi = 0
    OtrasTaras = 0
    If Text1(21).Text <> "" Then PesoBruto = CCur(Text1(21).Text)
    If Text1(11).Text <> "" Then TaraVehi = CCur(Text1(11).Text)
    If Text1(2).Text <> "" Then OtrasTaras = CCur(Text1(2).Text)
    PesoNeto = PesoBruto - Tara0 - Tara1 - Tara2 - Tara3 - Tara4 - Tara5 - TaraVehi - OtrasTaras
    PesoTrans = PesoBruto - Tara00 - Tara11 - Tara12 - Tara13 - Tara14 - Tara15 - TaraVehi - OtrasTaras
    Text1(12).Text = CStr(PesoNeto)
    Text1(24).Text = CStr(PesoTrans)
    PonerFormatoEntero Text1(12)
    
    ComprobarKilosCaja Text1(22).Text, PesoNeto
    
End Sub


Private Sub ComprobarKilosCaja(CajasReal As String, Neto As Currency)
Dim KiloCaja As Currency

    If ComprobarCero(CajasReal) <> 0 Then
        KiloCaja = Round2(Neto / CCur(CajasReal), 2)
        Text1(4).Text = CStr(KiloCaja)
        PonerFormatoEntero Text1(4)
        
        If Not EstaKilosCajaEntreLimites(KiloCaja) Then
            MsgBox "El valor de Kilos/Cajón ha de estar entre " & vParamAplic.KilosCajaMin & " y " & vParamAplic.KilosCajaMax & ".", vbExclamation
        End If
    End If

End Sub

Private Function EstaKilosCajaEntreLimites(KilCaj As Currency) As Boolean
    EstaKilosCajaEntreLimites = (KilCaj >= vParamAplic.KilosCajaMin) And (KilCaj <= vParamAplic.KilosCajaMax)
End Function


Private Sub CalcularPesoNeto()
Dim KgsCajon As String

    On Error GoTo eCalcularPesoNeto

    KgsCajon = DevuelveDesdeBDNew(cAgro, "variedades", "kgscajon", "codvarie", txtAux(5).Text, "N")
    If ComprobarCero(KgsCajon) <> 0 Then
        txtAux(10).Text = Round2(CCur(txtAux(9).Text) * CCur(KgsCajon), 0)
        PonerFormatoEntero txtAux(10)
        txtAux(2).Text = Round2(Data1.Recordset!KilosBrut * txtAux(10).Text / Data1.Recordset!KilosNet, 0)
        txtAux(11).Text = "0"
    Else
        txtAux(10).Text = "0"
        txtAux(2).Text = "0"
        txtAux(11).Text = "0"
    End If
    
    Exit Sub
    
eCalcularPesoNeto:
    MuestraError Err.Number, "Calculando Peso Neto"
End Sub


Private Sub CalcularPesoNetoPicassent()
Dim KgsCajon As String
Dim TotalCaj As Currency
Dim cajas As Currency
Dim TotalNeto As Currency

    On Error GoTo eCalcularPesoNeto


    TotalCaj = ComprobarCero(Text1(22).Text)
    TotalNeto = ComprobarCero(Text1(12).Text)
    cajas = ComprobarCero(txtAux(9).Text)
    
    If TotalCaj <> 0 Then
        txtAux(10).Text = Round2(cajas * Data1.Recordset!KilosNet / TotalCaj, 0)
        PonerFormatoEntero txtAux(10)
        txtAux(2).Text = Round2(cajas * Data1.Recordset!KilosBrut / TotalCaj, 0)
        
        '[Monica]19/10/2016: hacemos que los kilostrans sean los mismos que los netos
        txtAux(11).Text = Round2(cajas * Data1.Recordset!KilosNet / TotalCaj, 0)
    Else
        txtAux(10).Text = "0"
        txtAux(2).Text = "0"
        txtAux(11).Text = "0"
    End If
    
    Exit Sub
    
eCalcularPesoNeto:
    MuestraError Err.Number, "Calculando Peso Neto"
End Sub




Private Function InsertarLineaEntrada() As Boolean
Dim Sql As String
    
    On Error GoTo EInsertarLineaEntrada

    InsertarLineaEntrada = False
    
    'Inserta en tabla "facturas_envases"
    Sql = "INSERT INTO rentradas "
    Sql = Sql & "(nropesada, numlinea, numnotac, codvarie, codsocio, codcampo, "
    Sql = Sql & "codcapat, codtarif, tipoentr, recolect, transportadopor, numcajo1, kilosnet, codtrans,"
    Sql = Sql & "fechaent, horaentr, kilosbru )"
    Sql = Sql & "VALUES (" & DBSet(txtAux(1).Text, "N") & ", " & DBSet(txtAux(3).Text, "N") & ", " & DBSet(txtAux(4).Text, "N") & ","
    Sql = Sql & DBSet(txtAux(5).Text, "N") & ", "
    Sql = Sql & DBSet(txtAux(6).Text, "N") & ", "
    Sql = Sql & DBSet(txtAux(7).Text, "N") & ", " & DBSet(txtAux(0).Text, "N") & ", "
    Sql = Sql & DBSet(txtAux(8).Text, "N") & ","
    Sql = Sql & DBSet(Combo1(0).ListIndex, "N") & ","
    Sql = Sql & DBSet(Combo1(1).ListIndex, "N") & ","
    Sql = Sql & DBSet(Combo1(2).ListIndex, "N") & ","
    Sql = Sql & DBSet(txtAux(9).Text, "N") & ","
    Sql = Sql & DBSet(txtAux(10).Text, "N") & ","
    Sql = Sql & DBSet(Text1(3).Text, "T") & ","
    Sql = Sql & DBSet(Text1(1).Text, "F") & ","
    Sql = Sql & DBSet(Text1(1).Text & " " & Format(Now, "hh:mm:ss"), "FH") & ","
    Sql = Sql & DBSet(txtAux(2).Text, "N")
    Sql = Sql & ")"
    
    'insertar la linea
    conn.Execute Sql

    InsertarLineaEntrada = True
    Exit Function

EInsertarLineaEntrada:
    MuestraError Err.Number, "Insertar Linea Entrada", Err.Description
End Function



Private Function InsertarTemporal(Operacion As String, Mens As String) As Boolean
Dim Linea As String
Dim SQL1 As String
    
    On Error GoTo eInsertarTemporal
    
    InsertarTemporal = False
    
    Linea = SugerirCodigoSiguienteStr("tmppesada", "contador", "codusu = " & vUsu.Codigo)
    
    SQL1 = "insert into tmppesada (codusu, operacion, nropesada, numnotac, contador) "
    SQL1 = SQL1 & " values ("
    SQL1 = SQL1 & vUsu.Codigo & ","
    SQL1 = SQL1 & DBSet(Operacion, "T") & ","
    
    If Operacion = "I" Then
        SQL1 = SQL1 & DBSet(txtAux(1).Text, "N") & ","
        SQL1 = SQL1 & DBSet(txtAux(4).Text, "N") & ","
    Else
        SQL1 = SQL1 & DBLet(Adoaux(1).Recordset.Fields(0), "N") & ","
        SQL1 = SQL1 & DBLet(Adoaux(1).Recordset.Fields(2), "N") & ","
    End If
    SQL1 = SQL1 & DBSet(Linea, "N") & ")"

    conn.Execute SQL1

    InsertarTemporal = True
    Exit Function
    
eInsertarTemporal:
    Mens = Mens & " " & Err.Description
    InsertarTemporal = False
End Function


Private Function SeHaModificadoCabecera() As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cadena As String
Dim NumF As String
Dim Producto As String
' Se ha modificado en cabecera la matricula(transportista) o la fecha y se modifican en todas las lineas de pesada
    
    On Error GoTo eSeHaModificadoCabecera
    
    SeHaModificadoCabecera = False
    
    Sql = "select codvarie, numcajo1, numcajo2, numcajo3, numcajo4, numcajo5, nropesada, numnotac, codsocio,"
    Sql = Sql & "codcampo, codcapat, codtarif, kilosbru, kilosnet, fechaent "
    Sql = Sql & " from rentradas where nropesada = " & DBSet(Text1(0).Text, "N")
    Sql = Sql & " order by nropesada, numnotac "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Producto = DevuelveValor("select codprodu from variedades where codvarie = " & DBSet(Rs!codvarie, "N"))
    
        cadena = v_cadena & "<ROW notacamp=" & """" & Format(DBLet(Rs!NumNotac, "N"), "######0") & """"
        cadena = cadena & " fechaent=" & """" & Format(Text1(1).Text, "yyyymmdd") & """"
        cadena = cadena & " codprodu=" & """" & Format(DBLet(Producto, "N"), "#####0") & """"
        cadena = cadena & " codvarie=" & """" & Format(DBLet(Rs!codvarie, "N"), "#####0") & """"
        cadena = cadena & " codsocio=" & """" & Format(DBLet(Rs!Codsocio, "N"), "#####0") & """"
        cadena = cadena & " codcampo=" & """" & Format(DBLet(Rs!codCampo, "N"), "#######0") & """"
        cadena = cadena & " kilosbru=" & """" & Format(DBLet(Rs!KilosBru, "N"), "###0") & """"
        cadena = cadena & " kilosnet=" & """" & Format(DBLet(Rs!KilosNet, "N"), "###0") & """"
        cadena = cadena & " numcajo1=" & """" & Format(DBLet(Rs!numcajo1, "N"), "##0") & """"
        cadena = cadena & " numcajo2=" & """" & Format(0, "##0") & """"
        cadena = cadena & " numcajo3=" & """" & Format(0, "##0") & """"
        cadena = cadena & " numcajo4=" & """" & Format(0, "##0") & """"
        cadena = cadena & " numcajo5=" & """" & Format(0, "##0") & """"
        
        cadena = cadena & " matricul=" & """" & Text1(3).Text & """" ' tranportista
        cadena = cadena & " codcapat=" & """" & Format(DBLet(Rs!codcapat, "N"), "###0") & """"
        cadena = cadena & " identifi=" & """" & Format(0, "#####0") & """"
        cadena = cadena & " altura=" & """" & Format(vParamAplic.CajasporPalet, "##0") & """"
        cadena = cadena & " zona=" & """" & Format(0, "#########0") & """"
        cadena = cadena & " /></ROWDATA></DATAPACKET>"
    
        
        NumF = SugerirCodigoSiguienteStr("chivato", "numorden")
        
        Sql = "insert into chivato (numorden, basedato, nomtabla, operacio, fechadia, separado,"
        Sql = Sql & "claveant, clavenue, nombmemo, nombmem1, nombmem2, horaproc, nombmem3, nombmem4) values ("
        Sql = Sql & DBSet(NumF, "N") & ","
        Sql = Sql & "'agro',"
        Sql = Sql & "'sentba',"
        Sql = Sql & "'U',"
        Sql = Sql & DBSet(Now, "F") & ","
        Sql = Sql & DBSet("&", "T") & ","
        Sql = Sql & DBSet(Rs!NumNotac, "N") & ","
        Sql = Sql & DBSet(Rs!NumNotac, "N") & ","
        Sql = Sql & DBSet(cadena, "T") & ","
        Sql = Sql & ValorNulo & ","
        Sql = Sql & ValorNulo & ","
        Sql = Sql & "'" & Format(Now, "hh:mm:ss") & "',"
        Sql = Sql & ValorNulo & ","
        Sql = Sql & ValorNulo & ")"
        
        
'[Monica] 03/12/2009
'        Sql = "insert into chivato (numlinea, basedato, tabla, operacio, fechadia, separado,"
'        Sql = Sql & "claveant, clavenue, xml) values ("
'        Sql = Sql & DBSet(NumF, "N") & ","
'        Sql = Sql & "'agro',"
'        Sql = Sql & "'sentba',"
'        Sql = Sql & "'U',"
'        Sql = Sql & DBSet(Now, "FH") & ","
'        Sql = Sql & DBSet("&", "T") & ","
'        Sql = Sql & DBSet(Rs!numnotac, "N") & ","
'        Sql = Sql & DBSet(Rs!numnotac, "N") & ","
'        Sql = Sql & DBSet(Cadena, "T") & ")"
         
        conn.Execute Sql
        
        Rs.MoveNext
    
    Wend
    
    Set Rs = Nothing
    
        
    SeHaModificadoCabecera = True
    Exit Function
    
eSeHaModificadoCabecera:
    MuestraError Err.Number, "Modificado Cabecera", Err.Description
End Function


Private Function ActualizarChivato(Mens As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Sql2 As String
Dim Rs1 As ADODB.Recordset
Dim cadena As String
Dim Producto As String
Dim NumF As String

    On Error GoTo eActualizarChivato

    ActualizarChivato = False
    
    Sql = "select codvarie, numcajo1, numcajo2, numcajo3, numcajo4, numcajo5, numnotac, codsocio, codcampo, codcapat, codtarif, "
    Sql = Sql & "kilosbru, kilosnet, tipoentr, fechaent, codtrans, nropesada "
    Sql = Sql & "from rentradas"
    Sql = Sql & " where nropesada = " & Data1.Recordset.Fields!nropesada
    Sql = Sql & " order by nropesada, numnotac "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    While Not Rs.EOF
        Producto = DevuelveValor("select codprodu from variedades where codvarie = " & DBSet(Rs!codvarie, "N"))
        
        cadena = v_cadena & "<ROW notacamp=" & """" & Format(DBLet(Rs!NumNotac, "N"), "######0") & """"
        cadena = cadena & " fechaent=" & """" & Format(Rs!FechaEnt, "yyyymmdd") & """"
        cadena = cadena & " codprodu=" & """" & Format(DBLet(Producto, "N"), "#####0") & """"
        cadena = cadena & " codvarie=" & """" & Format(DBLet(Rs!codvarie, "N"), "#####0") & """"
        cadena = cadena & " codsocio=" & """" & Format(DBLet(Rs!Codsocio, "N"), "#####0") & """"
        cadena = cadena & " codcampo=" & """" & Format(DBLet(Rs!codCampo, "N"), "#######0") & """"
        cadena = cadena & " kilosbru=" & """" & Format(DBLet(Rs!KilosBru, "N"), "###0") & """"
        cadena = cadena & " kilosnet=" & """" & Format(DBLet(Rs!KilosNet, "N"), "###0") & """"
        cadena = cadena & " numcajo1=" & """" & Format(DBLet(Rs!numcajo1, "N"), "##0") & """"
        cadena = cadena & " numcajo2=" & """" & Format(0, "##0") & """"
        cadena = cadena & " numcajo3=" & """" & Format(0, "##0") & """"
        cadena = cadena & " numcajo4=" & """" & Format(0, "##0") & """"
        cadena = cadena & " numcajo5=" & """" & Format(0, "##0") & """"
        
        cadena = cadena & " matricul=" & """" & DBLet(Rs!codTrans, "T") & """"
        cadena = cadena & " codcapat=" & """" & Format(DBLet(Rs!codcapat, "N"), "###0") & """"
        cadena = cadena & " identifi=" & """" & Format(0, "#####0") & """"
        cadena = cadena & " altura=" & """" & Format(vParamAplic.CajasporPalet, "##0") & """"
        cadena = cadena & " zona=" & """" & Format(0, "#########0") & """"
        cadena = cadena & " /></ROWDATA></DATAPACKET>"
    
        Sql2 = "select * from tmppesada where codusu= " & vUsu.Codigo
        Sql2 = Sql2 & " and nropesada = " & Data1.Recordset.Fields!nropesada
        Sql2 = Sql2 & " and numnotac = " & DBSet(Rs!NumNotac, "N")
        Sql2 = Sql2 & " order by contador "
        
        Set Rs1 = New ADODB.Recordset
        Rs1.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs1.EOF
'            NumF = SugerirCodigoSiguienteStr("chivato", "numorden")
            
            NumF = DevuelveValor("select max(numorden) + 1 from chivato")
            
            
            Sql = "insert into chivato (numorden, basedato, nomtabla, operacio, fechadia, separado,"
            Sql = Sql & "claveant, clavenue, nombmemo, nombmem1, nombmem2, horaproc, nombmem3, nombmem4) values ("
            Sql = Sql & DBSet(NumF, "N") & ","
            Sql = Sql & "'agro',"
            Sql = Sql & "'sentba',"
            
            Select Case Rs1!Operacion
                Case "I" ' insertada
                    Sql = Sql & "'I',"
                Case "U" ' actualizada
                    Sql = Sql & "'U',"
                Case "Z" ' borrada
                    Sql = Sql & "'D',"
            End Select
            
            Sql = Sql & DBSet(Now, "F") & ","
            Sql = Sql & DBSet("&", "T") & ","
            Sql = Sql & DBSet(Rs!NumNotac, "N") & ","
            Sql = Sql & DBSet(Rs!NumNotac, "N") & ","
            Sql = Sql & DBSet(cadena, "T") & ","
            Sql = Sql & ValorNulo & ","
            Sql = Sql & ValorNulo & ","
            Sql = Sql & "'" & Format(Now, "hh:mm:ss") & "',"
            Sql = Sql & ValorNulo & ","
            Sql = Sql & ValorNulo & ")"
            
            conn.Execute Sql
            
            ' borramos de la temporal tras introducir en el chivato
            Sql = "delete from tmppesada where codusu = " & vUsu.Codigo
            Sql = Sql & " and nropesada = " & Rs!nropesada
            Sql = Sql & " and numnotac = " & Rs!NumNotac
            Sql = Sql & " and contador = " & Rs1!Contador
            
            conn.Execute Sql
            
            Rs1.MoveNext
        Wend
        
        Set Rs1 = Nothing
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    ActualizarChivato = True
    Exit Function
    
eActualizarChivato:
    Mens = Mens & Err.Description
End Function


Private Function CuadrarPesada(Mens As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim KilosNetos As Currency
Dim Kilos As Currency
Dim KilosTot As Currency
Dim KilosBruTot As Currency
Dim MaxNota As String

    On Error GoTo eCuadrarPesada
    
    CuadrarPesada = False
    
    Sql = DevuelveValor("select sum(numcajo1) from rentradas where nropesada= " & Data1.Recordset!nropesada)
    
    If CCur(Sql) <> Data1.Recordset!cajonesrea Then
        Mens = Mens & vbCrLf & vbCrLf & "La suma de los cajones no coincide con los cajones reales."
        Exit Function
    Else
        'como el numero de cajas es correcto tengo que repartir kilos
        'primero recalculamos kilosnetos de entradas
        Sql = "select * from rentradas where nropesada = " & Data1.Recordset!nropesada
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        While Not Rs.EOF
            txtAux(5).Text = DBLet(Rs!codvarie, "N")
            txtAux(9).Text = DBLet(Rs!numcajo1, "N")
            If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
                CalcularPesoNetoPicassent
            Else
                CalcularPesoNeto
            End If
            
            Sql2 = "update rentradas set kilosnet = " & DBSet(txtAux(10).Text, "N")
            Sql2 = Sql2 & ", kilosbru= " & DBSet(txtAux(2).Text, "N")
            Sql2 = Sql2 & ", kilostra= " & DBSet(txtAux(11).Text, "N")
            Sql2 = Sql2 & " where nropesada = " & Data1.Recordset!nropesada
            Sql2 = Sql2 & " and numnotac = " & DBSet(Rs!NumNotac, "N")
            
            conn.Execute Sql2
            
            Rs.MoveNext
        Wend
        Set Rs = Nothing
        
        ' sacamos el total de kilosnetos de las notas de la pesada
        Sql = "select sum(kilosnet) from rentradas where nropesada = " & Data1.Recordset!nropesada
        KilosNetos = DevuelveValor(Sql)
        
        'repartimos los kilos segun cajas en las entradas de la pesada
        Sql = "select * from rentradas where nropesada = " & Data1.Recordset!nropesada
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        KilosTot = 0
        KilosBruTot = 0
        While Not Rs.EOF
            Kilos = 0
            If KilosNetos <> 0 Then
                Kilos = Round2(Data1.Recordset!KilosNet * DBLet(Rs!KilosNet, "N") / KilosNetos, 0)
            End If
            
            KilosTot = KilosTot + Kilos
            KilosBruTot = KilosBruTot + DBLet(Rs!KilosBru, "N")
            
            Sql = "update rentradas set kilosnet = " & DBSet(Kilos, "N")
            '[Monica]19/10/2016: no tocabamos los kilostra
            If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
                Sql = Sql & ", kilostra = " & DBSet(Kilos, "N")
            End If
            Sql = Sql & " where nropesada = " & Data1.Recordset!nropesada
            Sql = Sql & " and numnotac = " & DBSet(Rs!NumNotac, "N")
            
            conn.Execute Sql
            
            Rs.MoveNext
        Wend
        Set Rs = Nothing
                
        ' si hay diferencia de kilos los metemos en la nota más alta
        If KilosTot <> Data1.Recordset!KilosNet Or KilosBruTot <> Data1.Recordset!KilosBrut Then
            Sql = "select max(numnotac) from rentradas where nropesada = " & Data1.Recordset!nropesada
            MaxNota = DevuelveValor(Sql)
            
            Sql = "update rentradas set kilosnet = kilosnet + " & (Data1.Recordset!KilosNet - KilosTot)
            Sql = Sql & ", kilosbru = kilosbru + " & (Data1.Recordset!KilosBrut - KilosBruTot)
            '[Monica]19/10/2016: no tocabamos los kilostra
            If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
                Sql = Sql & ", kilostra = kilostra + " & (Data1.Recordset!KilosNet - KilosTot)
            End If
            Sql = Sql & " where nropesada = " & Data1.Recordset!nropesada
            Sql = Sql & " and numnotac = " & DBSet(MaxNota, "N")
            
            conn.Execute Sql
        End If
    End If
    
    CuadrarPesada = True
    Exit Function
    
eCuadrarPesada:
    Mens = Mens & Err.Description
End Function


Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

