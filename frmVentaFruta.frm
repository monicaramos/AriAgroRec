VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmVentaFruta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Venta de Fruta Báscula"
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   13065
   Icon            =   "frmVentaFruta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   13065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   135
      TabIndex        =   82
      Top             =   90
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   83
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
      Left            =   3795
      TabIndex        =   80
      Top             =   90
      Width           =   885
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   81
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
               Object.ToolTipText     =   "Listado de Comprobación"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   4755
      TabIndex        =   78
      Top             =   90
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   79
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
      Left            =   10350
      TabIndex        =   77
      Top             =   225
      Width           =   1605
   End
   Begin VB.Frame Frame2 
      Height          =   3840
      Left            =   150
      TabIndex        =   27
      Top             =   840
      Width           =   12715
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
         Left            =   7515
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Tag             =   "Tipo Albarán|N|N|0|1|vtafrutacab|tipoalbaran|0|N|"
         Top             =   210
         Width           =   2010
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
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   9
         Tag             =   "Bultos 03|N|S|||vtafrutacab|bultos03|###,##0||"
         Top             =   2535
         Width           =   735
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
         Index           =   25
         Left            =   1935
         MaxLength       =   40
         TabIndex        =   10
         Tag             =   "Textos 03|T|S|||vtafrutacab|textos03|||"
         Top             =   2535
         Width           =   4035
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
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
         Left            =   9840
         MaxLength       =   6
         TabIndex        =   18
         Tag             =   "Peso Palot|N|S|||vtafrutacab|pesopalot|##0.00||"
         Top             =   2700
         Width           =   795
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
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
         Left            =   8130
         MaxLength       =   5
         TabIndex        =   17
         Tag             =   "Nro.Palot|N|S|||vtafrutacab|numpalot|#,##0||"
         Top             =   2700
         Width           =   1275
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
         Left            =   10860
         MaxLength       =   7
         TabIndex        =   19
         Tag             =   "Tara Palot|N|S|||vtafrutacab|tarapalot|#,##0||"
         Top             =   2700
         Width           =   1425
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
         Index           =   15
         Left            =   1935
         MaxLength       =   40
         TabIndex        =   8
         Tag             =   "Textos 02|T|S|||vtafrutacab|textos02|||"
         Top             =   2175
         Width           =   4035
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
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   7
         Tag             =   "Bultos 02|N|S|||vtafrutacab|bultos02|###,##0||"
         Top             =   2175
         Width           =   735
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
         Index           =   9
         Left            =   1935
         MaxLength       =   40
         TabIndex        =   6
         Tag             =   "Textos 01|T|S|||vtafrutacab|textos01|||"
         Top             =   1815
         Width           =   4035
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
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   5
         Tag             =   "Bultos 01|N|S|||vtafrutacab|bultos01|###,##0||"
         Top             =   1815
         Width           =   735
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
         Height          =   510
         Index           =   16
         Left            =   165
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Tag             =   "Observaciones|T|S|||vtafrutacab|observac|||"
         Top             =   3180
         Width           =   5805
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
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
         Left            =   1200
         MaxLength       =   12
         TabIndex        =   4
         Tag             =   "Matricula|T|S|||vtafrutacab|matriveh|||"
         Top             =   1425
         Width           =   1500
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
         Index           =   17
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   62
         Text            =   "Text2"
         Top             =   1020
         Width           =   3840
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
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Código Cliente|N|S|||vtafrutacab|codclien|000000||"
         Top             =   1020
         Width           =   930
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
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
         Left            =   7560
         MaxLength       =   7
         TabIndex        =   20
         Tag             =   "Tara Vehiculo|N|S|0|999999|vtafrutacab|taravehi|###,##0||"
         Top             =   3270
         Width           =   1425
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
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
         Left            =   10875
         MaxLength       =   7
         TabIndex        =   12
         Tag             =   "Peso Bruto|N|S|||vtafrutacab|pesobruto|###,##0||"
         Top             =   780
         Width           =   1425
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
         Left            =   8130
         MaxLength       =   5
         TabIndex        =   13
         Tag             =   "Nro.Cajas 1|N|S|||vtafrutacab|numcajon|#,##0||"
         Top             =   1830
         Width           =   1275
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
         Left            =   10860
         MaxLength       =   7
         TabIndex        =   14
         Tag             =   "Tara Cajas 1|N|S|||vtafrutacab|taracajon|#,##0||"
         Top             =   1830
         Width           =   1425
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
         Left            =   10860
         MaxLength       =   7
         TabIndex        =   16
         Tag             =   "Nro.Cajas 2|N|S|||vtafrutacab|tarapalet|#,##0||"
         Top             =   2280
         Width           =   1425
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
         Left            =   8130
         MaxLength       =   5
         TabIndex        =   15
         Tag             =   "Nro.Cajas 2|N|S|||vtafrutacab|numpalet|#,##0||"
         Top             =   2280
         Width           =   1275
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
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Código Socio|N|S|||vtafrutacab|codsocio|||"
         Top             =   660
         Width           =   930
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
         Left            =   4665
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Albarán|F|N|||vtafrutacab|fecalbar|dd/mm/yyyy|S|"
         Top             =   210
         Width           =   1350
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
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   30
         Text            =   "Text2"
         Top             =   660
         Width           =   3840
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
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
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Albaran|N|S|||vtafrutacab|numalbar|0000000|S|"
         Text            =   "Text1 7"
         Top             =   210
         Width           =   1020
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
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
         Left            =   10845
         MaxLength       =   7
         TabIndex        =   21
         Tag             =   "Peso Neto|N|S|||vtafrutacab|pesoneto|###,##0||"
         Top             =   3300
         Width           =   1425
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   21
         Left            =   2190
         MaxLength       =   10
         TabIndex        =   68
         Tag             =   "Tipo Movimiento|T|N|||vtafrutacab|codtipom||S|"
         Top             =   660
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Albarán"
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
         Left            =   6135
         TabIndex        =   75
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label1 
         Caption         =   "Bulto 3"
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
         Left            =   165
         TabIndex        =   74
         Top             =   2565
         Width           =   705
      End
      Begin VB.Label Label10 
         Caption         =   "x"
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
         Index           =   3
         Left            =   9690
         TabIndex        =   73
         Top             =   2730
         Width           =   285
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
         Height          =   225
         Index           =   2
         Left            =   10650
         TabIndex        =   72
         Top             =   2730
         Width           =   150
      End
      Begin VB.Label Label15 
         Caption         =   "PALOT"
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
         Index           =   2
         Left            =   6135
         TabIndex        =   71
         Top             =   2730
         Width           =   1830
      End
      Begin VB.Label Label1 
         Caption         =   "Bulto 2"
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
         Left            =   165
         TabIndex        =   67
         Top             =   2205
         Width           =   705
      End
      Begin VB.Label Label1 
         Caption         =   "Bulto 1"
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
         Left            =   165
         TabIndex        =   66
         Top             =   1845
         Width           =   705
      End
      Begin VB.Label Label29 
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
         Height          =   255
         Left            =   165
         TabIndex        =   65
         Top             =   2940
         Width           =   1530
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   1770
         ToolTipText     =   "Zoom descripción"
         Top             =   2940
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Matrícula"
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
         Left            =   165
         TabIndex        =   64
         Top             =   1455
         Width           =   1185
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   915
         ToolTipText     =   "Buscar Cliente"
         Top             =   1065
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
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
         Left            =   165
         TabIndex        =   63
         Top             =   1050
         Width           =   705
      End
      Begin VB.Label Label9 
         Caption         =   "Pesos y Taras:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   315
         Left            =   6165
         TabIndex        =   61
         Top             =   720
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
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   9705
         TabIndex        =   60
         Top             =   810
         Width           =   1125
      End
      Begin VB.Label Label8 
         Caption         =   "Tara Vehículo"
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
         Left            =   6135
         TabIndex        =   59
         Top             =   3300
         Width           =   1500
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   9690
         TabIndex        =   58
         Top             =   3315
         Width           =   1005
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
         Height          =   240
         Left            =   8115
         TabIndex        =   57
         Top             =   1410
         Width           =   1185
      End
      Begin VB.Line Line3 
         X1              =   6150
         X2              =   12240
         Y1              =   3195
         Y2              =   3195
      End
      Begin VB.Line Line2 
         X1              =   6150
         X2              =   12240
         Y1              =   1650
         Y2              =   1650
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
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
         Left            =   11145
         TabIndex        =   56
         Top             =   1410
         Width           =   1050
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
         Height          =   270
         Left            =   9720
         TabIndex        =   55
         Top             =   1410
         Width           =   1320
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
         Left            =   9690
         TabIndex        =   54
         Top             =   1860
         Width           =   1005
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
         Left            =   6150
         TabIndex        =   53
         Top             =   1860
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
         Height          =   225
         Index           =   1
         Left            =   6135
         TabIndex        =   52
         Top             =   2310
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
         Height          =   225
         Index           =   1
         Left            =   9690
         TabIndex        =   51
         Top             =   2310
         Width           =   975
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
         Left            =   10650
         TabIndex        =   50
         Top             =   1860
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
         Height          =   225
         Index           =   1
         Left            =   10650
         TabIndex        =   49
         Top             =   2310
         Width           =   150
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
         Left            =   3675
         TabIndex        =   31
         Top             =   240
         Width           =   585
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   4335
         Picture         =   "frmVentaFruta.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   240
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
         Height          =   255
         Index           =   0
         Left            =   165
         TabIndex        =   29
         Top             =   645
         Width           =   570
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   915
         ToolTipText     =   "Buscar Socio"
         Top             =   660
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "NºAlbarán"
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
         Left            =   165
         TabIndex        =   28
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   8
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   48
      Top             =   885
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   47
      Top             =   885
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   6
      Left            =   4290
      MaxLength       =   10
      TabIndex        =   46
      Top             =   885
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   4290
      MaxLength       =   10
      TabIndex        =   45
      Top             =   840
      Width           =   1065
   End
   Begin VB.Frame FrameAux1 
      Caption         =   "Variedades"
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
      Height          =   3570
      Left            =   120
      TabIndex        =   32
      Top             =   4800
      Width           =   12745
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
         Height          =   315
         Index           =   9
         Left            =   60
         MaxLength       =   12
         TabIndex        =   70
         Tag             =   "Tipo Mov|T|N|||vtafrutalin|codtipom||S|"
         Text            =   "Tipo M"
         Top             =   2310
         Visible         =   0   'False
         Width           =   420
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
         Height          =   315
         Index           =   8
         Left            =   1260
         MaxLength       =   10
         TabIndex        =   69
         Tag             =   "Fec.Albaran|F|N|||vtafrutalin|fecalbar|dd/mm/yyyy|S|"
         Text            =   "Fec.Alb"
         Top             =   2310
         Visible         =   0   'False
         Width           =   525
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
         Height          =   315
         Index           =   2
         Left            =   8130
         MaxLength       =   7
         TabIndex        =   37
         Tag             =   "Peso Bruto|N|N|||vtafrutalin|pesobruto|###,##0||"
         Text            =   "pesobru"
         Top             =   2280
         Visible         =   0   'False
         Width           =   945
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
         Height          =   315
         Index           =   0
         Left            =   7080
         MaxLength       =   7
         TabIndex        =   36
         Tag             =   "Num.Palet|N|N|0|999999|vtafrutalin|numpalet|###,##0||"
         Text            =   "numpale"
         Top             =   2280
         Visible         =   0   'False
         Width           =   945
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
         Height          =   315
         Index           =   4
         Left            =   6390
         MaxLength       =   7
         TabIndex        =   35
         Tag             =   "Num Cajon|N|N|0|999999|vtafrutalin|numcajon|###,##0||"
         Text            =   "cajon"
         Top             =   2280
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
         Index           =   0
         Left            =   3060
         MaskColor       =   &H00000000&
         TabIndex        =   42
         ToolTipText     =   "Buscar Variedad"
         Top             =   2310
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
         Height          =   315
         Index           =   5
         Left            =   3300
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   41
         Text            =   "Nombre variedad"
         Top             =   2310
         Width           =   1980
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
         Height          =   315
         Index           =   1
         Left            =   510
         MaxLength       =   12
         TabIndex        =   40
         Tag             =   "Num.Albaran|N|N|||vtafrutalin|numalbar|0000000|S|"
         Text            =   "Albaran"
         Top             =   2310
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
         Height          =   315
         Index           =   5
         Left            =   2130
         MaxLength       =   6
         TabIndex        =   33
         Tag             =   "Variedad|N|N|||vtafrutalin|codvarie|000000|N|"
         Text            =   "variedad"
         Top             =   2310
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
         Height          =   315
         Index           =   3
         Left            =   1860
         MaxLength       =   12
         TabIndex        =   39
         Tag             =   "Num.Linea|N|N|||vtafrutalin|numlinea|000|S|"
         Text            =   "Linea"
         Top             =   2310
         Visible         =   0   'False
         Width           =   180
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
         Height          =   315
         Index           =   7
         Left            =   5400
         MaxLength       =   10
         TabIndex        =   34
         Tag             =   "Desc.Calibre|T|S|||vtafrutalin|descalibre|||"
         Text            =   "calibre"
         Top             =   2280
         Visible         =   0   'False
         Width           =   825
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
         Height          =   315
         Index           =   6
         Left            =   9180
         MaxLength       =   7
         TabIndex        =   38
         Tag             =   "Peso neto|N|N|||vtafrutalin|pesoneto|###,##0||"
         Text            =   "pesonet"
         Top             =   2280
         Visible         =   0   'False
         Width           =   945
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   1
         Left            =   270
         TabIndex        =   43
         Top             =   300
         Width           =   1440
         _ExtentX        =   2540
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
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "frmVentaFruta.frx":0097
         Height          =   2640
         Left            =   270
         TabIndex        =   44
         Top             =   780
         Width           =   12310
         _ExtentX        =   21722
         _ExtentY        =   4657
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
   End
   Begin VB.Frame Frame1 
      Height          =   570
      Index           =   0
      Left            =   120
      TabIndex        =   25
      Top             =   8415
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
         Left            =   120
         TabIndex        =   26
         Top             =   225
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
      Left            =   11805
      TabIndex        =   23
      Top             =   8505
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
      Left            =   10590
      TabIndex        =   22
      Top             =   8505
      Width           =   1065
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
      Left            =   11820
      TabIndex        =   24
      Top             =   8475
      Visible         =   0   'False
      Width           =   1065
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
      Left            =   12390
      TabIndex        =   84
      Top             =   165
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
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir Albarán"
         HelpContextID   =   2
         Shortcut        =   ^I
      End
      Begin VB.Menu mnListComprobacion 
         Caption         =   "Listado Comprobación"
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
Attribute VB_Name = "frmVentaFruta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public Albaran As String  ' venimos de albaranes para ver las facturas donde aparece el albaran

'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1

Private WithEvents frmVar As frmComVar 'variedades comercial
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmCli As frmBasico2 'clientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmTrans As frmManTranspor 'transportista
Attribute frmTrans.VB_VarHelpID = -1
Private WithEvents frmTar As frmManTarTra 'tarifa de transportista
Attribute frmTar.VB_VarHelpID = -1
Private WithEvents frmCamp As frmManCampos 'campos
Attribute frmCamp.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'mensajes
Attribute frmMens.VB_VarHelpID = -1

Private WithEvents frmVFruta As frmBasico2
Attribute frmVFruta.VB_VarHelpID = -1

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
Dim Indice As Byte

Dim TipoFactura As Byte
Private BuscaChekc As String

Dim FechaAnt As String
Dim TransporAnt As String
Dim CajonreaAnt As String
Dim NetoAnt As String

Dim v_cadena As String

Dim Tara1 As Currency
Dim Tara2 As Currency
Dim Cajon1 As String
Dim Cajon2 As String
Dim TaraVehiAnt As Long
Dim TaraPalotAnt As Long


Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 0 'Variedades
            Set frmVar = New frmComVar
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco txtAux(5)
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub


Private Sub cmdAceptar_Click()
Dim I As Integer

Dim V As Integer

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda

        Case 3  'AÑADIR
            If DatosOK Then InsertarCabecera

        Case 4  'MODIFICAR
            If DatosOK Then
                If ModificaCabecera Then
                    TerminaBloquear
                    PosicionarData
                    PonerCampos
                    PonerCamposLineas
                End If
            End If
            
         Case 5 'InsertarModificar LINEAS
            Select Case ModificaLineas
                Case 1 'afegir llínia
                    InsertarLinea
                Case 2 'modificar llínies
                    If ModificarLinea Then
                        V = Adoaux(1).Recordset.Fields(3) 'el 2 es el nº de llinia
                        CargaGrid DataGrid3, Adoaux(1), True
                
                        DataGrid3.SetFocus
                        Adoaux(1).Recordset.Find (Adoaux(1).Recordset.Fields(1).Name & " =" & V)
                
                        LLamaLineas ModificaLineas, 0, "DataGrid3"
                        
                        PosicionarData
                        PonerCampos
                        PonerCamposLineas
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
            PonerModo 2
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
    
    
    PonerModo 3
    
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    Text1(21).Text = CodTipoMov
    
    TaraVehiAnt = 0
    TaraPalotAnt = 0
    LimpiarDataGrids
    
    Text1(3).Enabled = True
    Text1(17).Enabled = True
    
    imgBuscar(0).Enabled = True
    imgBuscar(1).Enabled = True
    
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
        
        'Si pasamos el control aqui lo ponemos en amarillo
        Text1(21).Text = CodTipoMov
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
    If chkVistaPrevia.Value = 1 Then
        LimpiarCampos
        LimpiarDataGrids
        EsCabecera = True
        MandaBusquedaPrevia ""
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select vtafrutacab.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla
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

    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    If Text1(17).Text <> "" Then PonerFoco Text1(17)
    If Text1(3).Text <> "" Then PonerFoco Text1(3)
        
    TaraVehiAnt = CInt(ImporteSinFormato(ComprobarCero(Text1(11).Text)))
    TaraPalotAnt = CInt(ImporteSinFormato(ComprobarCero(Text1(22).Text)))
        
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
    

    PonerModo 5, Index
    
    vWhere = ObtenerWhereCP(False)
    vWhere = vWhere & " and numlinea=" & Adoaux(1).Recordset!numlinea
    If Not BloqueaRegistro("vtafrutalin", vWhere) Then
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
        anc = anc + 210
    Else
        anc = anc + DataGrid3.RowTop(DataGrid3.Row) + 10
    End If

    txtAux(9).Text = DataGrid3.Columns(0).Text ' tipo de movimiento AVF
    txtAux(1).Text = DataGrid3.Columns(1).Text ' nro de albaran
    txtAux(8).Text = DataGrid3.Columns(2).Text ' fecha de albaran
    txtAux(3).Text = DataGrid3.Columns(3).Text ' numlinea
    txtAux(5).Text = DataGrid3.Columns(4).Text ' variedad
    Text2(5).Text = DataGrid3.Columns(5).Text ' nombre de la variedad
    txtAux(7).Text = DataGrid3.Columns(6).Text ' descripcion de calibre
    txtAux(4).Text = DataGrid3.Columns(7).Text ' numcajon 1
    txtAux(0).Text = DataGrid3.Columns(8).Text ' numcajon 2
    txtAux(2).Text = DataGrid3.Columns(9).Text ' peso bruto
    txtAux(6).Text = DataGrid3.Columns(10).Text ' peso neto
    
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



Private Sub LLamaLineas(xModo As Byte, Optional alto As Single, Optional grid As String)
Dim jj As Integer
Dim B As Boolean
    
    Select Case grid
        Case "DataGrid3"
            DeseleccionaGrid Me.DataGrid3
            B = (xModo = 1 Or xModo = 2)
            
            For jj = 0 To 0
                txtAux(jj).Height = DataGrid3.RowHeight - 10
                txtAux(jj).Top = alto + 5
                txtAux(jj).visible = B
                txtAux(jj).Enabled = B
            Next jj
            For jj = 2 To 2
                txtAux(jj).Height = DataGrid3.RowHeight - 10
                txtAux(jj).Top = alto + 5
                txtAux(jj).visible = B
                txtAux(jj).Enabled = B
            Next jj
            For jj = 4 To 7
                txtAux(jj).Height = DataGrid3.RowHeight - 10
                txtAux(jj).Top = alto + 5
                txtAux(jj).visible = B
                txtAux(jj).Enabled = B
            Next jj
            
            Text2(5).Height = DataGrid3.RowHeight - 10
            Text2(5).Top = alto + 5
            Text2(5).visible = B
           
            For jj = 0 To btnBuscar.Count - 1
                btnBuscar(jj).Height = DataGrid3.RowHeight - 10
                btnBuscar(jj).Top = alto + 5
                btnBuscar(jj).visible = B
            Next jj
            
    End Select
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Facturas (scafac)
' y los registros correspondientes de las tablas cab. albaranes (scafac1)
' y las lineas de la factura (slifac)
Dim cad As String
'Dim vTipoMov As CTiposMov

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
'    'solo se puede modificar la factura si no esta contabilizada
'    If FactContabilizada Then Exit Sub
    
    cad = "Cabecera de Albarán Venta Fruta." & vbCrLf
    cad = cad & "-------------------------------------      " & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar el Albarán:            "
    cad = cad & vbCrLf & "Nº Albarán:  " & Format(Text1(0).Text, "0000000")
    cad = cad & vbCrLf & "Fecha:  " & Format(Text1(1).Text, "dd/mm/yyyy")

    cad = cad & vbCrLf & vbCrLf & " ¿Desea Eliminarlo? "

    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
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
Dim cad As String

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
        cad = Data1.Recordset.Fields(0) & "|"
        cad = cad & Data1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub

Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbYellow
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim I As Integer
    If Combo1(Index).BackColor = vbYellow Then Combo1(Index).BackColor = vbWhite
    Select Case Index
        Case 0
            I = Combo1(Index).ListIndex
            PosicionarCombo Combo1(Index), CInt(I)
    End Select

End Sub

Private Sub DataGrid3_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'    If Not Adoaux(1).Recordset.EOF And ModificaLineas <> 1 Then
'        If Not IsNull(Adoaux(1).Recordset.Fields(0).Value) Then
'            Text2(6).Text = DevuelveDesdeBDNew(cAgro, "rsocios", "nomsocio", "codsocio", Adoaux(1).Recordset!CodSocio, "N")
'            Text2(0).Text = DevuelveDesdeBDNew(cAgro, "rcapataz", "nomcapat", "codcapat", Adoaux(1).Recordset!codcapat, "N")
'            Text2(8).Text = DevuelveDesdeBDNew(cAgro, "rtarifatra", "nomtarif", "codtarif", Adoaux(1).Recordset!Codtarif, "N")
'            PonerDatosCampo CStr(Adoaux(1).Recordset!codcampo)
'        End If
'    Else
'        Text2(6).Text = ""
'        Text2(0).Text = ""
'        Text2(8).Text = ""
'
'        Text2(4).Text = ""
'        Text2(2).Text = ""
'    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Load()
Dim I As Integer
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
        'el 10  son separadors
        .Buttons(8).Image = 10  'Imprimir
    End With
    
    With Me.Toolbar5
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26 ' Listado de Comprobacion
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
        End With
    Next kCampo
   ' ***********************************
    
    CargarParametrosTaras
    CargaCombo
    
    LimpiarCampos   'Limpia los campos TextBox

    CodTipoMov = "AVF"
    VieneDeBuscar = False
    
    '## A mano
    NombreTabla = "vtafrutacab"
    NomTablaLineas = "vtafrutalin" 'Tabla notas de entrada
    Ordenacion = " ORDER BY vtafrutacab.numalbar"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    CadenaConsulta = "select * from vtafrutacab "
    If Albaran <> "" Then
        CadenaConsulta = CadenaConsulta & " where numalbar = " & Albaran
    Else
        CadenaConsulta = CadenaConsulta & " where numalbar = -1"
    End If
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
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



Private Sub CargarParametrosTaras()
Dim I As Integer
    
    Tara1 = 0
    Tara2 = 0
    Cajon1 = ""
    Cajon2 = ""

    For I = 0 To 1
        Me.Label15(I).Caption = ""
        Me.Label19(I).Caption = ""
    Next I
    
    ' cargamos los labels de parametros
    If vParamAplic.EsVtaFruta1 Then
        Me.Label15(0).Caption = vParamAplic.TipoCaja1
        Me.Label19(0).Caption = "x  " & vParamAplic.PesoCaja1
        Cajon1 = vParamAplic.TipoCaja1
        Tara1 = vParamAplic.PesoCaja1
        If vParamAplic.EsVtaFruta2 Then
            Me.Label15(1).Caption = vParamAplic.TipoCaja2
            Me.Label19(1).Caption = "x  " & vParamAplic.PesoCaja2
            Cajon2 = vParamAplic.TipoCaja2
            Tara2 = vParamAplic.PesoCaja2
        Else
            If vParamAplic.EsVtaFruta3 Then
                Me.Label15(1).Caption = vParamAplic.TipoCaja3
                Me.Label19(1).Caption = "x  " & vParamAplic.PesoCaja3
                Cajon2 = vParamAplic.TipoCaja3
                Tara2 = vParamAplic.PesoCaja3
            Else
                If vParamAplic.EsVtaFruta4 Then
                    Me.Label15(1).Caption = vParamAplic.TipoCaja4
                    Me.Label19(1).Caption = "x  " & vParamAplic.PesoCaja4
                    Cajon2 = vParamAplic.TipoCaja4
                    Tara2 = vParamAplic.PesoCaja4
                Else
                    If vParamAplic.EsVtaFruta5 Then
                        Me.Label15(1).Caption = vParamAplic.TipoCaja5
                        Me.Label19(1).Caption = "x  " & vParamAplic.PesoCaja5
                        Cajon2 = vParamAplic.TipoCaja5
                        Tara2 = vParamAplic.PesoCaja5
                    End If
                End If
            End If
        End If
    Else
        If vParamAplic.EsVtaFruta2 Then
            Me.Label15(0).Caption = vParamAplic.TipoCaja2
            Me.Label19(0).Caption = "x  " & vParamAplic.PesoCaja2
            Cajon1 = vParamAplic.TipoCaja2
            Tara1 = vParamAplic.PesoCaja2
            If vParamAplic.EsVtaFruta3 Then
                Me.Label15(1).Caption = vParamAplic.TipoCaja3
                Me.Label19(1).Caption = "x  " & vParamAplic.PesoCaja3
                Cajon2 = vParamAplic.TipoCaja3
                Tara2 = vParamAplic.PesoCaja3
            Else
                If vParamAplic.EsVtaFruta4 Then
                    Me.Label15(1).Caption = vParamAplic.TipoCaja4
                    Me.Label19(1).Caption = "x  " & vParamAplic.PesoCaja4
                    Cajon2 = vParamAplic.TipoCaja4
                    Tara2 = vParamAplic.PesoCaja4
                Else
                    If vParamAplic.EsVtaFruta5 Then
                        Me.Label15(1).Caption = vParamAplic.TipoCaja5
                        Me.Label19(1).Caption = "x  " & vParamAplic.PesoCaja5
                        Cajon2 = vParamAplic.TipoCaja5
                        Tara2 = vParamAplic.PesoCaja5
                    End If
                End If
            End If
        Else
            If vParamAplic.EsVtaFruta3 Then
                Me.Label15(0).Caption = vParamAplic.TipoCaja3
                Me.Label19(0).Caption = "x  " & vParamAplic.PesoCaja3
                Cajon1 = vParamAplic.TipoCaja3
                Tara1 = vParamAplic.PesoCaja3
                If vParamAplic.EsVtaFruta4 Then
                    Me.Label15(1).Caption = vParamAplic.TipoCaja4
                    Me.Label19(1).Caption = "x  " & vParamAplic.PesoCaja4
                    Cajon2 = vParamAplic.TipoCaja4
                    Tara2 = vParamAplic.PesoCaja4
                Else
                    If vParamAplic.EsVtaFruta5 Then
                        Me.Label15(1).Caption = vParamAplic.TipoCaja5
                        Me.Label19(1).Caption = "x  " & vParamAplic.PesoCaja5
                        Cajon2 = vParamAplic.TipoCaja5
                        Tara2 = vParamAplic.PesoCaja5
                    End If
                End If
            Else
                If vParamAplic.EsVtaFruta4 Then
                    Me.Label15(0).Caption = vParamAplic.TipoCaja4
                    Me.Label19(0).Caption = "x  " & vParamAplic.PesoCaja4
                    Cajon1 = vParamAplic.TipoCaja4
                    Tara1 = vParamAplic.PesoCaja4
                    If vParamAplic.EsVtaFruta5 Then
                        Me.Label15(1).Caption = vParamAplic.TipoCaja5
                        Me.Label19(1).Caption = "x  " & vParamAplic.PesoCaja5
                        Cajon2 = vParamAplic.TipoCaja5
                        Tara2 = vParamAplic.PesoCaja5
                    End If
                Else
                    If vParamAplic.EsVtaFruta5 Then
                        Me.Label15(0).Caption = vParamAplic.TipoCaja5
                        Me.Label19(0).Caption = "x  " & vParamAplic.PesoCaja5
                        Cajon1 = vParamAplic.TipoCaja5
                        Tara1 = vParamAplic.PesoCaja5
                    End If
                End If
            End If
        End If
    End If
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Combo1(0).ListIndex = -1
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Cancel = 0
    
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
        CadB = CadB & Aux
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
        CadB = CadB & " and " & Aux
        
        
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

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(7).Text = RecuperaValor(CadenaSeleccion, 1)
    PonerDatosCampo txtAux(7)
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codsocio
    FormateaCampo Text1(3)
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'nomsocio
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de cliente
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Codigo de clientes
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(5).Text = RecuperaValor(CadenaSeleccion, 1)  'Codigo de variedad
    Text2(5).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre
End Sub

Private Sub frmVFruta_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String

    If CadenaSeleccion <> "" Then
        Text1(0).Text = RecuperaValor(CadenaSeleccion, 1) 'albaran
        
        CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
        If CadB <> "" Then
            'Se muestran en el mismo form
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda
        End If
    End If

End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(Indice).Text = vCampo
End Sub


Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. de socio
            Indice = 3
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco Text1(3)
            
       Case 1 ' codigo de cliente
            Indice = 17
            Set frmCli = New frmBasico2
            AyudaClienteCom frmCli, Text1(Indice)
            Set frmCli = Nothing
            PonerFoco Text1(17)
            
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
        Indice = 16
        frmZ.pTitulo = "Observaciones del Albarán"
        frmZ.pValor = Text1(Indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco Text1(Indice)
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

Private Sub mnImprimir_Click()
    BotonImprimir
End Sub

Private Sub mnListComprobacion_Click()
    AbrirListado (35)
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


Private Sub BotonImprimir()
Dim cadFormula As String
Dim CadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim NroCopias As Integer
Dim Tipo As Byte

    If Text1(0).Text = "" Then
        MsgBox "Debe seleccionar una Factura para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    CadParam = ""
    cadSelect = ""
    numParam = 0
    
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de albaran
    '---------------------------------------------------
    'Tipo de factura
    devuelve = "{" & NombreTabla & ".codtipom}='" & CodTipoMov & "'"
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    devuelve = "codtipom = '" & CodTipoMov & "'"
    If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
    
    indRPT = 82
    
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
      
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
    
    'Nº Albaran
    devuelve = "{" & NombreTabla & ".numalbar}=" & Val(Text1(0).Text)
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    devuelve = "numalbar = " & Val(Text1(0).Text)
    If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
    
    'Fecha Albaran
    devuelve = "{" & NombreTabla & ".fecalbar}=Date(" & Year(Text1(1).Text) & "," & Month(Text1(1).Text) & "," & Day(Text1(1).Text) & ")"
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    devuelve = "fecalbar = " & DBSet(Text1(1).Text, "F")
    If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
    
    If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub
     
    With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = CadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 0
            .Titulo = "Impresión de Albarán de Fruta"
            .ConSubInforme = True
            .Show vbModal
    End With
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
Dim devuelve As String
Dim cadMen As String
Dim Sql As String
Dim Nregs As Long
Dim Tara As String
Dim PesoNeto As Long
Dim PesoPalot As Long
        
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1 'Fecha albaran
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index), True
            
        Case 3 ' Socio
            If Text1(Index).Text <> "" Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "rsocios", "nomsocio")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Socio: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmSoc = New frmManSocios
                        frmSoc.DatosADevolverBusqueda = "0|1|"
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmSoc.Show vbModal
                        Set frmSoc = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        Case 17 ' codclien
            If Text1(Index).Text <> "" Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "clientes", "nomclien")
                If Text2(Index).Text = "" Then
                    MsgBox "Código no existe. Revise.", vbExclamation
                    PonerFoco Text1(17)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
       Case 2, 10 ' numero de bultos 1 y 2
            PonerFormatoEntero Text1(Index)
            
       Case 11 ' tara de vehiculo
            If Modo = 1 Then Exit Sub
            PonerFormatoEntero Text1(Index)
'            Text1(12).Text = CInt(ImporteSinFormato(ComprobarCero(Text1(12).Text))) + TaraVehiAnt - CInt(ImporteSinFormato(ComprobarCero(Text1(Index).Text)))
'            If TaraVehiAnt <> ComprobarCero(Text1(11).Text) Then
                PesoNeto = CLng(ImporteSinFormato(ComprobarCero(Text1(12).Text))) + TaraVehiAnt - CLng(ImporteSinFormato(ComprobarCero(Text1(11).Text)))
                Text1(12).Text = Format(PesoNeto, "###,##0")
'            End If

        Case 23, 24 ' nro de palots, peso de palot
            If Modo = 1 Then Exit Sub
            If Index = 23 Then
                PonerFormatoEntero Text1(Index)
            Else
                PonerFormatoDecimal Text1(Index), 4
            End If
            If ComprobarCero(Text1(23).Text) <> "0" And ComprobarCero(Text1(24).Text) <> "0" Then
                PesoPalot = Round2(CLng(Text1(23).Text) * CCur(ImporteSinFormato(Text1(24).Text)), 0)
                Text1(22).Text = Format(PesoPalot, "###0")
            Else
                Text1(22).Text = ""
            End If
            
            PesoNeto = CLng(ImporteSinFormato(ComprobarCero(Text1(12).Text))) + TaraVehiAnt - CLng(ImporteSinFormato(ComprobarCero(Text1(11).Text))) + TaraPalotAnt - CInt(ImporteSinFormato(ComprobarCero(Text1(22).Text)))
            Text1(12).Text = Format(PesoNeto, "###,##0")
            
    End Select
    
    If (Index = 3 Or Index = 17) And Modo <> 1 And Modo <> 5 Then
        If Index = 3 Then
            Text1(17).Enabled = (Text1(3).Text = "")
            imgBuscar(1).Enabled = (Text1(3).Text = "")
            imgBuscar(1).visible = (Text1(3).Text = "")
        End If
        If Index = 17 Then
            Text1(3).Enabled = (Text1(17).Text = "")
            imgBuscar(0).Enabled = (Text1(17).Text = "")
            imgBuscar(0).visible = (Text1(17).Text = "")
        End If
    End If
    
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
        CadenaConsulta = "select vtafrutacab.* from " & NombreTabla
        CadenaConsulta = CadenaConsulta & " WHERE " & CadB & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
    'Llamamos a al form
    '##A mano
'    cad = ""
'    cad = cad & "Nº.Albarán|vtafrutacab.numalbar|N||15·"
'    cad = cad & "Cliente/Socio|concat(if(vtafrutacab.codclien is null,'',vtafrutacab.codclien),if(vtafrutacab.codsocio is null,'',vtafrutacab.codsocio)) as codigo|N||20·" 'ParaGrid(Text1(3), 10, "Cliente")
'    cad = cad & "Nombre Cliente/Socio|concat(if(clientes.nomclien is null,'',clientes.nomclien), if(rsocios.nomsocio is null,'',rsocios.nomsocio)) as nombre|N||45·"
'    cad = cad & ParaGrid(Text1(1), 15, "F.Albarán")
'    tabla = "(" & NombreTabla & " LEFT JOIN clientes ON vtafrutacab.codclien=clientes.codclien) "
'    tabla = tabla & " left join rsocios On vtafrutacab.codsocio = rsocios.codsocio "
'
'    Titulo = "Albaranes de Venta Fruta"
'    devuelve = "0|3|"
'
'    If cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vtabla = tabla
'        frmB.vSQL = CadB
'        HaDevueltoDatos = False
'        '###A mano
'        frmB.vDevuelve = "0|1|4|"
'        frmB.vDevuelve = devuelve
'        frmB.vTitulo = Titulo
'        frmB.vSelElem = 0
''        frmB.vConexionGrid = cAgro  'Conexión a BD: Ariagro
'        If Not EsCabecera Then frmB.Label1.FontSize = 11
''        frmB.vBuscaPrevia = chkVistaPrevia
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
''        If EsCabecera Then
''            PonerCadenaBusqueda
''            Text1(0).Text = Format(Text1(0).Text, "0000000")
''        End If
'        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
'        End If
'    End If
'    Screen.MousePointer = vbDefault

    Set frmVFruta = New frmBasico2
    
    AyudaVtaFruta frmVFruta
    
    Set frmVFruta = Nothing


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
'            Text1(0).BackColor = vbYellow
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
Dim B As Boolean
Dim b2 As Boolean
Dim I As Integer

    On Error GoTo EPonerLineas

    If Data1.Recordset.EOF Then Exit Sub

    Screen.MousePointer = vbHourglass
    
    If Data1.Recordset.RecordCount > 0 Then
        CargaGrid DataGrid3, Adoaux(1), True
    Else
        CargaGrid DataGrid3, Adoaux(1), False
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
Dim B As Boolean

    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    
'    b = PonerCamposForma2(Me, Data1, 2, "FrameDatosPesosTaras")
    B = PonerCamposForma(Me, Data1)
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
Dim I As Byte, NumReg As Byte
Dim B As Boolean

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
    B = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Or Albaran <> "" Then
        cmdRegresar.visible = B
    Else
        cmdRegresar.visible = False
    End If
        
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    DesplazamientoVisible B And Data1.Recordset.RecordCount > 1
          
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    BloquearCombo Me, Modo
    
    
    B = (Modo <> 1)
    'Campos Nº Albarán bloqueado y en azul
    BloquearTxt Text1(0), B, True
    
    B = (Modo <> 1) And (Modo <> 3)
    BloquearTxt Text1(1), B  'fecalbaran
    
    'Si no es modo lineas Boquear los TxtAux
    For I = 1 To 1
        txtAux(I).visible = False
        BloquearTxt txtAux(I), True
    Next I
    For I = 3 To 9
        txtAux(I).visible = False
        BloquearTxt txtAux(I), True
    Next I
    
    For I = 5 To 5
        Text2(I).visible = ((Modo = 5) And (indFrame = 1))
        Text2(I).Enabled = False
    Next I
    
    For I = 0 To 0
        BloquearBtn Me.btnBuscar(I), (ModificaLineas = 0)
    Next I
    
    
    '---------------------------------------------
'    b = (Modo <> 0 And Modo <> 2) Or (Modo = 5 And ModificaLineas <> 0)
    B = (Modo = 1) Or Modo = 3 Or Modo = 4 Or (Modo = 5 And ModificaLineas <> 0)
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    
    BloquearImgBuscar Me, Modo, ModificaLineas
    BloquearImgFec Me, 0, Modo
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    ' *** si n'hi han llínies i algún tab que no te datagrid ***
    BloquearFrameAux Me, "FrameAux1", Modo, 1
    
'    'Campos Nº entrada bloqueado y en azul
'    BloquearTxt Text1(0), b, True
    
    'taras desbloqueadas unicamente para buscar
    For I = 18 To 20
        BloquearTxt Text1(I), Not (Modo = 1)
        Text1(I).Enabled = Modo = 1
    Next I
    For I = 22 To 22
        BloquearTxt Text1(I), Not (Modo = 1)
        Text1(I).Enabled = Modo = 1
    Next I
    For I = 12 To 14
        BloquearTxt Text1(I), Not (Modo = 1)
        Text1(I).Enabled = Modo = 1
    Next I
    
        
    ' ***************************
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
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


Private Function DatosOK() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim B As Boolean
Dim Serie As String
Dim Sql As String

    On Error GoTo EDatosOK

    DatosOK = False
    
'    ComprobarDatosTotales

    'comprobamos datos OK de la tabla scaalb
    B = CompForm2(Me, 2, "Frame2") ' , 1) 'Comprobar formato datos ok de la cabecera: opcion=1
    If Not B Then Exit Function
    
    If Modo = 3 Or Modo = 4 Then
        If Text1(3).Text <> "" And Text1(17).Text <> "" Then
            MsgBox "El albarán sólo puede ser o de cliente o de socio, pero no de ambos.", vbExclamation
            B = False
            PonerFoco Text1(3)
        Else
            If Text1(3).Text = "" And Text1(17).Text = "" Then
                MsgBox "El albarán ha de ser de cliente o de socio. Revise.", vbExclamation
                B = False
                PonerFoco Text1(3)
            End If
        End If
    End If
    
    DatosOK = B
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
Dim B As Boolean
Dim I As Byte

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    B = True

    For I = 4 To 7
        If I = 4 Or I = 6 Or I = 7 Then
            If txtAux(I).Text = "" Then
                MsgBox "El campo " & txtAux(I).Tag & " no puede ser nulo", vbExclamation
                B = False
                PonerFoco txtAux(I)
                Exit Function
            End If
        End If
    Next I
            
    DatosOkLinea = B
    
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)

    If BloqueaRegistro(NombreTabla, "numalbar = " & Data1.Recordset!numalbar) Then
'    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
        Select Case Button.Index
            Case 1
                BotonAnyadirLinea Index
            Case 2
                BotonModificarLinea Index
            Case 3
                BotonEliminarLinea Index
            Case Else
        End Select
    End If

End Sub


Private Sub BotonEliminarLinea(Index As Integer)
Dim cad As String
Dim Sql As String
Dim Mens As String
Dim B As Boolean

    On Error GoTo EEliminarLinea

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    ' ***************************************************************************

    B = True

    ' *************** canviar la pregunta ****************
    cad = "¿Seguro que desea eliminar la Variedad del Albarán ?"
    cad = cad & vbCrLf & "Albarán: " & Adoaux(1).Recordset.Fields(1)
    cad = cad & vbCrLf & "Variedad: " & Adoaux(1).Recordset.Fields(4) & " " & Adoaux(1).Recordset.Fields(5)
    
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminarLinea
        Screen.MousePointer = vbHourglass
        NumRegElim = Adoaux(1).Recordset.AbsolutePosition
        
        If Not EliminarLinea Then
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            PosicionarData
            If SituarDataTrasEliminar(Adoaux(1), NumRegElim) Then
                PonerCampos
                PonerModo 2
            Else
                LimpiarCampos
                PonerModo 0
            End If
        End If
    End If
    Screen.MousePointer = vbDefault
       
EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Or Not B Then MuestraError Err.Number, "Eliminar Linea de Pesada", Err.Description

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Añadir
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
        Case 5  'Buscar
            mnBuscar_Click
        Case 6  'Todos
            BotonVerTodos
        Case 8  ' Impresion
            mnImprimir_Click
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
Dim B As Boolean
Dim Opcion As Byte
Dim Sql As String

    On Error GoTo ECargaGrid

    B = DataGrid3.Enabled
    Select Case vDataGrid.Name
        Case "DataGrid3" 'notas de entrada
            Opcion = 1
    End Select
    
    Sql = MontaSQLCarga(enlaza, Opcion)
    CargaGridGnral vDataGrid, vData, Sql, PrimeraVez
    
    vDataGrid.RowHeight = 270
    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
    
    B = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
    vDataGrid.Enabled = Not B
    
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim tots As String
Dim I As Integer

    On Error GoTo ECargaGrid

    Select Case vDataGrid.Name
         Case "DataGrid3" 'rentradas
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux(5)|T|Código|900|;S|btnBuscar(0)|B|||;"
            tots = tots & "S|Text2(5)|T|Variedad|2600|;"
            tots = tots & "S|txtAux(7)|T|Calibre|1100|;S|txtAux(4)|T|" & Mid(LCase(Trim(Cajon1)), 1, 15) & "|2100|;S|txtAux(0)|T|" & Mid(LCase(Trim(Cajon2)), 1, 15) & "|2100|;"
            tots = tots & "S|txtAux(2)|T|Peso Bruto|1400|;S|txtAux(6)|T|Peso Neto|1400|;"
            
            arregla tots, DataGrid3, Me, 350
    End Select
    
    For I = 2 To 6
        DataGrid3.Columns(I).Alignment = dbgLeft
    Next I
    
    For I = 7 To 9
        DataGrid3.Columns(I).Alignment = dbgRight
    Next I
        
    vDataGrid.HoldFields
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub

Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  ' Listado de Comprobacion
            mnListComprobacion_Click
    End Select
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub



Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Sql As String
Dim devuelve As String
Dim B As Boolean
Dim TipoDto As Byte


    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
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
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 4, 0, 2, 6 'Cajas 1, cajas 2, pesobruto  y pesoneto
            If PonerFormatoEntero(txtAux(Index)) Then
                If Index <> 6 Then                                                                'peso bruto distinto de 0
                    If txtAux(0).Text <> "" And txtAux(4).Text <> "" And txtAux(2).Text <> "" And ComprobarCero(txtAux(2).Text) <> 0 Then
                        txtAux(6).Text = CalcularPesoNetoLin
                    End If
                Else
                    If ComprobarCero(txtAux(6).Text) <> CalcularPesoNetoLin Then
                        ' limpiamos los valores de cajas y demás
'                        txtAux(2).Text = ""
'                        txtAux(4).Text = ""
'                        txtAux(0).Text = ""
                    End If
                    cmdAceptar.SetFocus
                End If
            End If
            
    End Select
    
End Sub




Private Function Eliminar() As Boolean
Dim Sql As String, LEtra As String, Sql2 As String
Dim Rs As ADODB.Recordset
Dim Sql1 As String
Dim B As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String
Dim NumF As Long
    On Error GoTo FinEliminar

    If Data1.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    B = True

    If B Then
        'Eliminar en tablas de cabecera de albaran
        '------------------------------------------
        Sql = " " & ObtenerWhereCP(True)
        
        'Lineas de variedades (vtafrutalin)
        conn.Execute "Delete from vtafrutalin " & Sql
        
        'Cabecera de albaran
        conn.Execute "Delete from " & NombreTabla & Sql
        
        'Decrementar contador si borramos el ult. albaran
        Set vTipoMov = New CTiposMov
        vTipoMov.DevolverContador CodTipoMov, Val(Text1(0).Text)
        Set vTipoMov = Nothing
    End If
    
FinEliminar:
    If Err.Number <> 0 Or Not B Then
        MuestraError Err.Number, "Eliminar Albarán Venta", Err.Description & " " & Mens
        B = False
    End If
    If Not B Then
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function

Private Function EliminarLinea() As Boolean
Dim Sql As String, LEtra As String, Sql1 As String
Dim B As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String
Dim Linea As Long
Dim vWhere As String

    On Error GoTo FinEliminar

    B = False
    If Adoaux(1).Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    Mens = ""
    
    
    'Eliminar en tablas de vtafrutalin
    '------------------------------------------
    Sql = " where codtipom = '" & Adoaux(1).Recordset.Fields(0) & "'"
    Sql = Sql & " and numalbar = " & Adoaux(1).Recordset.Fields(1)
    Sql = Sql & " and fecalbar = " & DBSet(Adoaux(1).Recordset.Fields(2), "F")
    Sql = Sql & " and numlinea = " & Adoaux(1).Recordset.Fields(3)

    'Lineas de variedades
    conn.Execute "Delete from vtafrutalin " & Sql
    
    vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
    B = True
    If B Then
        B = ActualizarPesonetoreal(vWhere)
    End If
    
    If B Then
        Mens = "Actualizando Cacecera"
        B = ActualizarCabecera("I", Mens)
    End If
    
FinEliminar:
    If Err.Number <> 0 Or Not B Then
        MuestraError Err.Number, "Eliminar Variedad del Albaran ", Err.Description & " " & Mens
        B = False
    End If
    If Not B Then
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
    
    Sql = "numalbar= " & DBSet(Text1(0).Text, "N")
    Sql = Sql & " and codtipom = " & DBSet(Text1(21).Text, "T")
    Sql = Sql & " and fecalbar = " & DBSet(Text1(1).Text, "F")
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
        Case 1  'vtafrutalin
            Sql = "SELECT codtipom, numalbar, fecalbar, numlinea, vtafrutalin.codvarie, variedades.nomvarie, descalibre, numcajon, numpalet, pesobruto, pesoneto "
            Sql = Sql & " FROM vtafrutalin, variedades "
            Sql = Sql & " WHERE vtafrutalin.codvarie = variedades.codvarie "
    End Select
    
    If enlaza Then
        Sql = Sql & " and " & ObtenerWhereCP(False)
    Else
        Sql = Sql & " and numalbar = -1"
    End If
    Sql = Sql & " ORDER BY numalbar, numlinea"
    MontaSQLCarga = Sql
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim B As Boolean, bAux As Boolean
Dim I As Integer

        B = ((Modo = 2) Or (Modo = 0)) And (Albaran = "") 'Or (Modo = 5 And ModificaLineas = 0)
        'Buscar
        Toolbar1.Buttons(5).Enabled = B
        Me.mnBuscar.Enabled = B
        'Vore Tots
        Toolbar1.Buttons(6).Enabled = B
        Me.mnVerTodos.Enabled = B
        'Añadir
        Toolbar1.Buttons(1).Enabled = B
        Me.mnModificar.Enabled = B
        
        
        B = (Modo = 2 And Data1.Recordset.RecordCount > 0) And (Albaran = "")
        'Modificar
        Toolbar1.Buttons(2).Enabled = B
        Me.mnModificar.Enabled = B
        'eliminar
        Toolbar1.Buttons(3).Enabled = B
        Me.mnEliminar.Enabled = B
        'Imprimir albaran
        Toolbar1.Buttons(8).Enabled = (Modo = 2) Or (Albaran <> "")
        Me.mnImprimir.Enabled = (Modo = 2) Or (Albaran <> "")
        'Listado de Comprobacion
        Toolbar5.Buttons(1).Enabled = (Modo = 2) Or (Modo = 0) Or (Albaran <> "")
        Me.mnListComprobacion.Enabled = (Modo = 2) Or (Modo = 0) Or (Albaran <> "")

    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
'++monica: si insertamos lo he quitado
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
'    b = (Modo = 5) And (Albaran = "")
'    For i = 1 To 1
'        ToolAux(i).Buttons(1).Enabled = b ' añadir y salir siempre activos
'        ToolAux(i).Buttons(4).Enabled = b
'
'        If b Then
'            bAux = (b And Me.Adoaux(1).Recordset.RecordCount > 0)
'        End If
'        ToolAux(i).Buttons(2).Enabled = bAux
'        ToolAux(i).Buttons(3).Enabled = bAux
'    Next i

    B = (Modo = 3 Or Modo = 4 Or Modo = 2) And DatosADevolverBusqueda = ""
    For I = 1 To 1
        ToolAux(I).Buttons(1).Enabled = B
        If B Then bAux = (B And Me.Adoaux(I).Recordset.RecordCount > 0)
        ToolAux(I).Buttons(2).Enabled = bAux
        ToolAux(I).Buttons(3).Enabled = bAux
    Next I


End Sub

Private Function ModificaCabecera() As Boolean
Dim B As Boolean
Dim MenError As String
Dim Sql As String
Dim vWhere As String

    On Error GoTo EModificarCab

    conn.BeginTrans
    
    
    B = ModificaDesdeFormulario2(Me, 2, "Frame2")
    
    vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
    
    If B Then
        B = ActualizarPesonetoreal(vWhere)
    End If
    
    If B Then
        MenError = "Actualizando Cabecera "
        B = ActualizarCabecera("I", MenError)
    End If

EModificarCab:
    If Err.Number <> 0 Then
        MenError = "Modificando Albaran de Fruta." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        B = False
    End If
    If B Then
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
        devuelve = DevuelveDesdeBDNew(cAgro, NombreTabla, "numalbar", "numalbar", Text1(0), "N")
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
    MenError = "Error al insertar en la tabla Cabecera de Venta de Fruta (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    MenError = "Error al actualizar el contador del Albarán."
    vTipoMov.IncrementarContador (CodTipoMov)
    
EInsertarOferta:
    If Err.Number <> 0 Then
        MenError = "Insertando Albarán de Venta." & vbCrLf & "----------------------------" & vbCrLf & MenError
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
Dim B As Boolean
Dim Mens As String

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
'        Case 0: nomFrame = "FrameAux0" 'variedades
    nomframe = "FrameAux1" 'lineas de albaran
    ' ***************************************************************
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        
        If InsertarLineaEnv(txtAux(3).Text) Then
'            CalcularDatosAlbaran
            B = BloqueaRegistro("vtafrutacab", "numalbar = " & Data1.Recordset!numalbar)
            CargaGrid DataGrid3, Adoaux(1), True
            
            PosicionarData
            PonerCampos
            PonerCamposLineas
            
            If B Then BotonAnyadirLinea 1
        End If
    End If
End Sub

Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vtabla As String
Dim anc As Single
Dim I As Integer
    
    ModificaLineas = 1 'Posem Modo Afegir Llínia
    
    If Me.Adoaux(1).Recordset.RecordCount >= 6 Then
        MsgBox "Sólo se permiten un máximo de 6 líneas por albarán para que quepa en la impresión." & vbCrLf & vbCrLf & "Cree un nuevo albarán con el resto de movimientos.", vbExclamation
        Exit Sub
    End If
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModificaLineas = 0 Then Exit Sub
    End If
       
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    BloquearTxt Text1(1), True
'    BloquearTxt Text1(1), True
'
    
    ' **************************************************

    ' *** posar el nom del les distintes taules de llínies ***
    vtabla = "vtafrutalin"
    ' ********************************************************
    
    vWhere = ObtenerWhereCab(False)
    
    NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)
    ' ***************************************************************

    AnyadirLinea DataGrid3, Adoaux(1)

    anc = DataGrid3.Top
    If DataGrid3.Row < 0 Then
        anc = anc + 215 '210
    Else
        anc = anc + DataGrid3.RowTop(DataGrid3.Row) + 5
    End If
  
    LLamaLineas ModificaLineas, anc, "DataGrid3"

    LimpiarCamposLin "FrameAux1"
    txtAux(9).Text = Text1(21).Text
    txtAux(1).Text = Text1(0).Text 'nro de albaran
    txtAux(8).Text = Text1(1).Text 'fecha de albaran
    
    txtAux(3).Text = NumF
    txtAux(0).Text = "0"
    txtAux(4).Text = "0"
    txtAux(2).Text = "0"
    
    PonerFoco txtAux(5)
    For I = 5 To 5
        Text2(I).Text = ""
    Next I
    For I = 0 To btnBuscar.Count - 1
        BloquearBtn Me.btnBuscar(I), False
    Next I
    
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
Dim cad As String
Dim Sql As String
Dim B As Boolean
Dim Mens As String
Dim vWhere As String

    
    On Error GoTo eModificarLinea

    ModificarLinea = False
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
        B = ModificaDesdeFormulario2(Me, 2, "FrameAux1")
            
            
        'Conseguir el siguiente numero de linea
        vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
    '    numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
            
            
        If B Then
            B = ActualizarPesonetoreal(vWhere)
        End If
            
        If B Then
            Mens = "Actualizando Cabecera "
            B = ActualizarCabecera("I", Mens)
        End If
            
        ModificaLineas = 0
        
    End If
        
eModificarLinea:
    If Err.Number <> 0 Or Not B Then
        MuestraError Err.Number, "Modificar Lineas Factura" & vbCrLf & Err.Description & vbCrLf & Mens
        B = False
    End If
    
    If B Then
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
Dim B As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte
Dim Cliente As String

    DatosOkLlin = True
    
    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False

    B = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not B Then Exit Function
    
    DatosOkLlin = B

EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " codtipom = " & DBSet(Text1(21).Text, "T")
    vWhere = vWhere & " and numalbar= " & Val(Text1(0).Text) & " and fecalbar = " & DBSet(Text1(1).Text, "F")
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
    

Private Function InsertarLineaEnv(numlinea As String) As Boolean
'Inserta un registro en la tabla de lineas de Albaranes: slialb
Dim Sql As String
Dim vWhere As String
Dim B As Boolean
Dim DentroTRANS As Boolean
Dim Mens As String

    On Error GoTo EInsertarLineaEnv
    
    
    
    InsertarLineaEnv = False
    Sql = ""
    
    'Conseguir el siguiente numero de linea
    vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
'    numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
    
    conn.BeginTrans
    
    
    B = InsertarLineaEntrada
    
    If B Then
        B = ActualizarPesonetoreal(vWhere)
    End If
    
    If B Then
        Mens = "Actualizando Cabecera "
        B = ActualizarCabecera("I", Mens)
    End If
    
    
    If B Then
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

Private Function ActualizarPesonetoreal(vWhere)
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim TaraVehi As Long
Dim PesoBrutoTot As Long
Dim PesoNeto As Long
Dim TaraTot As Long
Dim Tara As Long
Dim I As Long
Dim TaraPalot As Long

    On Error GoTo eActualizarPesonetoreal
    
    
    ActualizarPesonetoreal = False

    TaraVehi = DevuelveValor("select taravehi from vtafrutacab where " & vWhere)
    TaraPalot = DevuelveValor("select tarapalot from vtafrutacab where " & vWhere)
    TaraVehi = TaraVehi + TaraPalot
    
    PesoBrutoTot = DevuelveValor("select sum(pesobruto) from vtafrutalin where " & vWhere)
    
    '[Monica]29/12/2011: si no nos ponen peso bruto nos tienen que poner peso neto
    If PesoBrutoTot = 0 Then
        Sql = "update vtafrutalin set pesonetoreal = pesoneto where " & vWhere
        conn.Execute Sql
        ActualizarPesonetoreal = True
        Exit Function
    End If

    Sql = "select * from vtafrutalin where " & vWhere & " order by numlinea"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    TaraTot = 0
    While Not Rs.EOF
        I = DBLet(Rs!numlinea, "N")
        PesoNeto = DBLet(Rs!PesoBruto, "N") - Round2(DBLet(Rs!Numcajon, "N") * Tara1, 0) - Round2(DBLet(Rs!NumPalet, "N") * Tara2, 0)

        Tara = Round2(TaraVehi * DBLet(Rs!PesoBruto, "N") / PesoBrutoTot, 0)
        PesoNeto = PesoNeto - Tara
        
        TaraTot = TaraTot + Tara

        Sql = "update vtafrutalin set pesonetoreal = " & DBSet(PesoNeto, "N")
        Sql = Sql & " where " & vWhere
        Sql = Sql & " and numlinea = " & DBSet(I, "N")

        conn.Execute Sql

        Rs.MoveNext
    Wend
    If I <> 0 Then
        If TaraTot <> TaraVehi Then
            Sql = "update vtafrutalin set pesonetoreal = pesonetoreal + " & DBSet(TaraVehi - TaraTot, "N")
            Sql = Sql & " where " & vWhere
            Sql = Sql & " and numlinea = " & DBSet(I, "N")
            
            conn.Execute Sql
        End If
    
    End If
    Set Rs = Nothing

    ActualizarPesonetoreal = True
    Exit Function
    
eActualizarPesonetoreal:
    MuestraError Err.Number, "Actualizando Peso Real", Err.Description
End Function

Private Sub PonerCamposSocioVariedad()
Dim cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim Rs As ADODB.Recordset


    If txtAux(6).Text = "" Or txtAux(5).Text = "" Then Exit Sub
    

    cad = "rcampos.codsocio = " & DBSet(txtAux(6).Text, "N") & " and rcampos.fecbajas is null"
    cad = cad & " and rcampos.codvarie = " & DBSet(txtAux(5), "N")
     
    Cad1 = "select count(*) from rcampos where " & cad
     
    NumRegis = TotalRegistros(Cad1)
    
    If NumRegis = 0 Then Exit Sub
    If NumRegis = 1 Then
        Cad1 = "select codcampo from rcampos where " & cad
        Set Rs = New ADODB.Recordset
        Rs.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            txtAux(7).Text = DBLet(Rs.Fields(0).Value)
            PonerDatosCampo txtAux(7).Text
        End If
    Else
        Set frmMens = New frmMensajes
        frmMens.cadWHERE = " and " & cad '"rcampos.codsocio = " & NumCod & " and rcampos.fecbajas is null"
        frmMens.campo = txtAux(7).Text
        frmMens.OpcionMensaje = 6
        frmMens.Show vbModal
        Set frmMens = Nothing
    End If
    
End Sub

Private Sub PonerDatosCampo(campo As String)
Dim cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim Rs As ADODB.Recordset


    If campo = "" Then Exit Sub
    
'    If Not (Modo = 3 Or Modo = 4) Then Exit Sub

    cad = "rcampos.codcampo = " & DBSet(campo, "N") & " and rcampos.fecbajas is null"
     
    Cad1 = "select rcampos.codparti, rpartida.nomparti, rpartida.codzonas, rzonas.nomzonas, "
    Cad1 = Cad1 & " rpueblos.despobla from rcampos, rpartida, rzonas, rpueblos "
    Cad1 = Cad1 & " where " & cad
    Cad1 = Cad1 & " and rcampos.codparti = rpartida.codparti "
    Cad1 = Cad1 & " and rpartida.codzonas = rzonas.codzonas "
    Cad1 = Cad1 & " and rpartida.codpobla = rpueblos.codpobla "
     
    Set Rs = New ADODB.Recordset
    Rs.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Text1(5).Text = ""
    Text2(4).Text = ""
    Text2(2).Text = ""
    If Not Rs.EOF Then
        Text1(5).Text = campo
        PonerFormatoEntero Text1(5)
        Text2(4).Text = DBLet(Rs.Fields(0).Value, "N") ' codigo de partida
        If Text2(4).Text <> "" Then Text2(4).Text = Format(Text2(4).Text, "0000")
        Text2(2).Text = DBLet(Rs.Fields(1).Value, "T") ' nombre de partida
    End If
    
    Set Rs = Nothing
    
End Sub

Private Sub PosarDescripcions()
Dim NomEmple As String

    On Error GoTo EPosarDescripcions

    Text2(3).Text = ""
    If Text1(3).Text <> "" Then
        Text2(3).Text = PonerNombreDeCod(Text1(3), "rsocios", "nomsocio", "codsocio", "N")
    End If
    Text2(17).Text = ""
    If Text1(17).Text <> "" Then
        Text2(17).Text = PonerNombreDeCod(Text1(17), "clientes", "nomclien", "codclien", "N")
    End If
    
EPosarDescripcions:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo descripciones", Err.Description
End Sub



Private Function CalcularPesoNetoLin() As Long
Dim NCajas1 As Long
Dim NCajas2 As Long
Dim PBruto As Long
Dim PNeto As Long

    On Error GoTo eCalcularPesoNetoLin
    
    CalcularPesoNetoLin = 0

    NCajas1 = ComprobarCero(txtAux(4).Text)
    NCajas2 = ComprobarCero(txtAux(0).Text)
 
    PBruto = ComprobarCero(txtAux(2).Text)
    PNeto = PBruto - Round2(NCajas1 * Tara1, 0) - Round2(NCajas2 * Tara2, 0)
    
    CalcularPesoNetoLin = PNeto
    Exit Function
eCalcularPesoNetoLin:
    MuestraError Err.Number, "Calculando Peso Neto"
End Function


Private Function InsertarLineaEntrada() As Boolean
Dim Sql As String
    
    On Error GoTo EInsertarLineaEntrada

    InsertarLineaEntrada = False
    
    'Inserta en tabla "vtafrutalin"
    Sql = "INSERT INTO vtafrutalin "
    Sql = Sql & "(codtipom, numalbar, fecalbar, numlinea, codvarie, descalibre, "
    Sql = Sql & "pesoneto, pesobruto, numcajon, numpalet)"
    Sql = Sql & "VALUES (" & DBSet(txtAux(9).Text, "T") & ", " & DBSet(txtAux(1).Text, "N") & ", " & DBSet(txtAux(8).Text, "F") & ","
    Sql = Sql & DBSet(txtAux(3).Text, "N") & ", " ' numero de linea
    Sql = Sql & DBSet(txtAux(5).Text, "N") & ", " ' variedad
    Sql = Sql & DBSet(txtAux(7).Text, "T") & ", " & DBSet(txtAux(6).Text, "N") & ", "
    Sql = Sql & DBSet(txtAux(2).Text, "N") & ","
    Sql = Sql & DBSet(txtAux(4).Text, "N") & ","
    Sql = Sql & DBSet(txtAux(0).Text, "N") & ")"
    
    'insertar la linea
    conn.Execute Sql

    InsertarLineaEntrada = True
    Exit Function

EInsertarLineaEntrada:
    MuestraError Err.Number, "Insertar Linea Entrada", Err.Description
End Function



Private Function ActualizarCabecera(Operacion As String, Mens As String) As Boolean
Dim Linea As String
Dim Sql1 As String
Dim Numcajon As Long
Dim NumPalet As Long
Dim TaraCajon As Long
Dim TaraPalet As Long
Dim PesoBruto As Long
Dim PesoNeto As Long
Dim Sql As String

    On Error GoTo eActualizarCabecera
    
    ActualizarCabecera = False
    
    
    Sql = "select sum(if(pesobruto is null,0,pesobruto)) from vtafrutalin where codtipom = " & DBSet(CodTipoMov, "T")
    Sql = Sql & " and numalbar = " & DBSet(txtAux(1).Text, "N")
    Sql = Sql & " and fecalbar = " & DBSet(txtAux(8).Text, "F")
    
    PesoBruto = DevuelveValor(Sql)
    
'    Sql = "select sum(if(pesoneto is null,0,pesoneto)) from vtafrutalin where codtipom = " & DBSet(CodTipoMov, "T")
'    Sql = Sql & " and numalbar = " & DBSet(txtAux(1).Text, "N")
'    Sql = Sql & " and fecalbar = " & DBSet(txtAux(8).Text, "F")
'
'    PesoNeto = DevuelveValor(Sql)
'
'    If PesoNeto = 0 Then
        Sql = "select sum(if(pesonetoreal is null,0,pesonetoreal)) from vtafrutalin where codtipom = " & DBSet(CodTipoMov, "T")
        Sql = Sql & " and numalbar = " & DBSet(txtAux(1).Text, "N")
        Sql = Sql & " and fecalbar = " & DBSet(txtAux(8).Text, "F")
        
        PesoNeto = DevuelveValor(Sql)
    
    
'    End If
    
    Sql = "select sum(if(round(numcajon * " & DBSet(Tara1, "N") & ",0) is null,0,round(numcajon * " & DBSet(Tara1, "N") & ",0) )) from vtafrutalin where codtipom = " & DBSet(CodTipoMov, "T")
    Sql = Sql & " and numalbar = " & DBSet(txtAux(1).Text, "N")
    Sql = Sql & " and fecalbar = " & DBSet(txtAux(8).Text, "F")
    
    TaraCajon = DevuelveValor(Sql)
    
    Sql = "select sum(if(round(numpalet * " & DBSet(Tara2, "N") & ",0) is null,0,round(numpalet * " & DBSet(Tara2, "N") & ",0) )) from vtafrutalin where codtipom = " & DBSet(CodTipoMov, "T")
    Sql = Sql & " and numalbar = " & DBSet(txtAux(1).Text, "N")
    Sql = Sql & " and fecalbar = " & DBSet(txtAux(8).Text, "F")
    
    TaraPalet = DevuelveValor(Sql)
    
    
    
    Sql = "select sum(if(numcajon is null,0,numcajon)) from vtafrutalin where codtipom = " & DBSet(CodTipoMov, "T")
    Sql = Sql & " and numalbar = " & DBSet(txtAux(1).Text, "N")
    Sql = Sql & " and fecalbar = " & DBSet(txtAux(8).Text, "F")
    
    Numcajon = DevuelveValor(Sql)
    
    Sql = "select sum(if(numpalet is null,0,numpalet)) from vtafrutalin where codtipom = " & DBSet(CodTipoMov, "T")
    Sql = Sql & " and numalbar = " & DBSet(txtAux(1).Text, "N")
    Sql = Sql & " and fecalbar = " & DBSet(txtAux(8).Text, "F")
    
    NumPalet = DevuelveValor(Sql)
    
    
    Sql = "update vtafrutacab set pesobruto = " & DBSet(PesoBruto, "N")
    Sql = Sql & ", pesoneto = " & DBSet(PesoNeto, "N") '& " - if(taravehi is null,0,taravehi) "
    Sql = Sql & ", taracajon = " & DBSet(TaraCajon, "N")
    Sql = Sql & ", tarapalet = " & DBSet(TaraPalet, "N")
    Sql = Sql & ", numcajon = " & DBSet(Numcajon, "N")
    Sql = Sql & ", numpalet = " & DBSet(NumPalet, "N")
    Sql = Sql & " where codtipom = " & DBSet(CodTipoMov, "T")
    Sql = Sql & " and numalbar = " & DBSet(txtAux(1).Text, "N")
    Sql = Sql & " and fecalbar = " & DBSet(txtAux(8).Text, "F")
    
    conn.Execute Sql
    
    ActualizarCabecera = True
    Exit Function
    
eActualizarCabecera:
    Mens = Mens & " " & Err.Description
    ActualizarCabecera = False
End Function


Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim I As Byte
    
    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For I = 0 To Combo1.Count - 1
        Combo1(I).Clear
    Next I
    
    'tipo de albaran
    Combo1(0).AddItem "Venta de Fruta"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Precalibrado"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
End Sub



