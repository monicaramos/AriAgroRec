VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmManTraba 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trabajadores"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   13950
   Icon            =   "frmManTraba.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   13950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Cuenta Principal|N|N|0|1|cltebanc|ctaprpal|0||"
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
      Index           =   0
      Left            =   11370
      TabIndex        =   89
      Top             =   180
      Width           =   1605
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   210
      TabIndex        =   87
      Top             =   0
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   88
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
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3900
      TabIndex        =   85
      Top             =   0
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   86
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
      Index           =   12
      Left            =   4020
      MaxLength       =   10
      TabIndex        =   83
      Tag             =   "Móvil|T|S|||straba|movtraba|||"
      Top             =   3255
      Width           =   1665
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
      Index           =   11
      Left            =   1305
      MaxLength       =   10
      TabIndex        =   82
      Tag             =   "Teléfono|T|S|||straba|teltraba|||"
      Top             =   3255
      Width           =   1455
   End
   Begin VB.Frame FrameDatosDtoAdministracion 
      Caption         =   "Datos Relacionados Nóminas"
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
      Height          =   5100
      Left            =   5820
      TabIndex        =   45
      Top             =   1620
      Width           =   7980
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
         Index           =   35
         Left            =   4095
         MaxLength       =   3
         TabIndex        =   90
         Tag             =   "Código Gr|N|S|0|999|straba|codbanpr|000||"
         Top             =   4455
         Width           =   915
      End
      Begin VB.CheckBox chkEmbarga 
         Caption         =   "Hay Embargo"
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
         Index           =   1
         Left            =   180
         TabIndex        =   35
         Tag             =   "Hay embargo|N|N|||straba|hayembargo||N|"
         Top             =   4440
         Width           =   1815
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
         Index           =   34
         Left            =   1470
         MaxLength       =   4
         TabIndex        =   25
         Tag             =   "IBAN|T|S|||straba|iban|||"
         Text            =   "Text1"
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
         Index           =   33
         Left            =   6615
         MaxLength       =   8
         TabIndex        =   34
         Tag             =   "Pr.Hora Coste|N|N|||straba|prhoracoste|##0.0000||"
         Top             =   4020
         Width           =   1230
      End
      Begin VB.TextBox text2 
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
         Left            =   2070
         TabIndex        =   79
         Top             =   4020
         Width           =   3030
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
         Index           =   32
         Left            =   1470
         MaxLength       =   3
         TabIndex        =   33
         Tag             =   "Horario|N|S|||straba|codhorario|000||"
         Top             =   4020
         Width           =   555
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
         Left            =   6615
         MaxLength       =   6
         TabIndex        =   32
         Tag             =   "Nro.Tarjeta|N|S|||straba|idtarjeta|000000||"
         Top             =   3630
         Width           =   1230
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
         Index           =   30
         Left            =   4095
         MaxLength       =   5
         TabIndex        =   31
         Tag             =   "Cod.Asesoria|N|S|||straba|codasesoria|00000||"
         Top             =   3630
         Width           =   1020
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
         Left            =   1470
         MaxLength       =   6
         TabIndex        =   30
         Tag             =   "%Retencion|N|S|||straba|dtoreten|##0.00||"
         Top             =   3630
         Width           =   645
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
         Left            =   6630
         MaxLength       =   10
         TabIndex        =   22
         Tag             =   "Plus Capataz|N|S|||straba|pluscapataz|###,##0.00||"
         Top             =   2040
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
         Index           =   27
         Left            =   6630
         MaxLength       =   6
         TabIndex        =   16
         Tag             =   "%Antigüedad|N|S|||straba|porc_antig|##0.00||"
         Top             =   1260
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
         Index           =   26
         Left            =   4260
         MaxLength       =   6
         TabIndex        =   15
         Tag             =   "%Seg.Soc|N|S|||straba|dtosegso|##0.00||"
         Top             =   1260
         Width           =   735
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
         Index           =   25
         Left            =   1470
         MaxLength       =   6
         TabIndex        =   14
         Tag             =   "%IRPF|N|S|||straba|dtosirpf|##0.00||"
         Top             =   1260
         Width           =   765
      End
      Begin VB.TextBox text2 
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
         Left            =   2070
         TabIndex        =   70
         Top             =   2430
         Width           =   5760
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
         Left            =   1470
         MaxLength       =   3
         TabIndex        =   23
         Tag             =   "Almacén|N|N|0|99|straba|codalmac|00||"
         Top             =   2430
         Width           =   555
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
         Left            =   6630
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Tag             =   "Tipo|N|N|||straba|codsecci||N|"
         Top             =   1650
         Width           =   1215
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
         Left            =   1470
         MaxLength       =   15
         TabIndex        =   20
         Tag             =   "Seguridad Social|T|N|||straba|segsocial|||"
         Top             =   2055
         Width           =   1785
      End
      Begin VB.TextBox text2 
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
         Left            =   2070
         TabIndex        =   47
         Top             =   2850
         Width           =   5760
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
         Index           =   20
         Left            =   1470
         MaxLength       =   3
         TabIndex        =   24
         Tag             =   "Código F.Pago|N|N|0|999|straba|codforpa|000||"
         Top             =   2850
         Width           =   555
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
         Left            =   3435
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Tag             =   "Tipo|N|N|||straba|tipotraba||N|"
         Top             =   1635
         Width           =   1575
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
         Left            =   2216
         MaxLength       =   4
         TabIndex        =   26
         Tag             =   "Banco|N|S|0|9999|straba|codbanco|0000||"
         Text            =   "Text1"
         Top             =   3240
         Width           =   675
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
         Index           =   19
         Left            =   2962
         MaxLength       =   4
         TabIndex        =   27
         Tag             =   "Sucursal|N|S|0|9999|straba|codsucur|0000||"
         Text            =   "Text1"
         Top             =   3240
         Width           =   675
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
         Index           =   8
         Left            =   3708
         MaxLength       =   2
         TabIndex        =   28
         Tag             =   "Digito Control|T|S|||straba|digcontr|00||"
         Text            =   "Text1"
         Top             =   3240
         Width           =   495
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
         Index           =   7
         Left            =   4275
         MaxLength       =   10
         TabIndex        =   29
         Tag             =   "Cuenta Bancaria|T|S|||straba|cuentaba|0000000000||"
         Text            =   "Text1"
         Top             =   3240
         Width           =   2055
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
         Index           =   9
         Left            =   1470
         MaxLength       =   3
         TabIndex        =   13
         Tag             =   "Categoria|N|N|0|99|straba|codcateg|00||"
         Top             =   840
         Width           =   555
      End
      Begin VB.TextBox text2 
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
         Left            =   2040
         TabIndex        =   46
         Top             =   840
         Width           =   5820
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
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "Fec.Alta|F|S|||straba|fechaalta|dd/mm/yyyy||"
         Top             =   450
         Width           =   1350
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
         Left            =   4050
         MaxLength       =   10
         TabIndex        =   11
         Tag             =   "Fecha Baja|F|S|||straba|fechabaja|dd/mm/yyyy||"
         Top             =   450
         Width           =   1350
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
         Left            =   6525
         MaxLength       =   10
         TabIndex        =   12
         Tag             =   "Fecha Antig|F|S|||straba|fecantig|dd/mm/yyyy||"
         Top             =   450
         Width           =   1350
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
         Left            =   1470
         MaxLength       =   3
         TabIndex        =   17
         Tag             =   "Grupo Cotizacion|N|N|0|999|straba|grupocot|000||"
         Top             =   1650
         Width           =   765
      End
      Begin VB.CheckBox chkAbonos 
         Caption         =   "Contrato"
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
         Left            =   3810
         TabIndex        =   21
         Tag             =   "Tipo|N|N|||straba|contrato||N|"
         Top             =   2040
         Width           =   1245
      End
      Begin VB.Label Label31 
         Caption         =   "Grupo trabajo"
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
         Left            =   2385
         TabIndex        =   91
         Top             =   4455
         Width           =   1425
      End
      Begin VB.Label Label30 
         Caption         =   "Pr.Coste Hora"
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
         Left            =   5160
         TabIndex        =   81
         Top             =   4050
         Width           =   1515
      End
      Begin VB.Label Label27 
         Caption         =   "Horario"
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
         TabIndex        =   80
         Top             =   4020
         Width           =   750
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1185
         ToolTipText     =   "Buscar Horario"
         Top             =   4050
         Width           =   240
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   0
         Left            =   6360
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   3690
         Width           =   240
      End
      Begin VB.Label Label26 
         Caption         =   "Nro.Tarjeta"
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
         Left            =   5160
         TabIndex        =   78
         Top             =   3660
         Width           =   1185
      End
      Begin VB.Label Label25 
         Caption         =   "Código Asesoria"
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
         Left            =   2385
         TabIndex        =   77
         Top             =   3660
         Width           =   1635
      End
      Begin VB.Label Label24 
         Caption         =   "% Retención"
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
         Left            =   150
         TabIndex        =   76
         Top             =   3660
         Width           =   1350
      End
      Begin VB.Label Label22 
         Caption         =   "Plus Capataz"
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
         Left            =   5220
         TabIndex        =   75
         Top             =   2040
         Width           =   1320
      End
      Begin VB.Label Label19 
         Caption         =   "% Antiguedad"
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
         Left            =   5190
         TabIndex        =   74
         Top             =   1260
         Width           =   1380
      End
      Begin VB.Label Label18 
         Caption         =   "% Seg.Social"
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
         Left            =   2910
         TabIndex        =   73
         Top             =   1260
         Width           =   1380
      End
      Begin VB.Label Label17 
         Caption         =   "% IRPF"
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
         TabIndex        =   72
         Top             =   1260
         Width           =   1020
      End
      Begin VB.Label Label15 
         Caption         =   "Almacén"
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
         TabIndex        =   71
         Top             =   2430
         Width           =   870
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1185
         ToolTipText     =   "Buscar Almacén"
         Top             =   2430
         Width           =   240
      End
      Begin VB.Label Label8 
         Caption         =   "Sección "
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
         Left            =   5220
         TabIndex        =   66
         Top             =   1650
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "Seg.Social"
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
         TabIndex        =   65
         Top             =   2055
         Width           =   1050
      End
      Begin VB.Label Label21 
         Caption         =   "Tipo "
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
         Left            =   2895
         TabIndex        =   55
         Top             =   1650
         Width           =   615
      End
      Begin VB.Label Label23 
         Caption         =   "F.Pago"
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
         TabIndex        =   54
         Top             =   2895
         Width           =   885
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1185
         ToolTipText     =   "Buscar F.Pago"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "IBAN Trab."
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
         Index           =   21
         Left            =   180
         TabIndex        =   53
         Top             =   3315
         Width           =   1005
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1185
         ToolTipText     =   "Buscar Categoria"
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Categoria"
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
         TabIndex        =   52
         Top             =   840
         Width           =   990
      End
      Begin VB.Label Label7 
         Caption         =   "Fec.Alta"
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
         TabIndex        =   51
         Top             =   450
         Width           =   945
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1185
         Picture         =   "frmManTraba.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   450
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Fec.Baja"
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
         Left            =   2895
         TabIndex        =   50
         Top             =   450
         Width           =   870
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   3780
         Picture         =   "frmManTraba.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   450
         Width           =   240
      End
      Begin VB.Label Label11 
         Caption         =   "F.Antig"
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
         Left            =   5430
         TabIndex        =   49
         Top             =   450
         Width           =   750
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   6255
         Picture         =   "frmManTraba.frx":0122
         ToolTipText     =   "Buscar fecha"
         Top             =   450
         Width           =   240
      End
      Begin VB.Label Label13 
         Caption         =   "Grupo Cot."
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
         TabIndex        =   48
         Top             =   1650
         Width           =   1170
      End
   End
   Begin VB.TextBox text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   23
      Left            =   8370
      TabIndex        =   68
      Top             =   3735
      Width           =   3840
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   23
      Left            =   7095
      MaxLength       =   10
      TabIndex        =   67
      Tag             =   "Cta.Contable|T|S|||straba|codmacta|||"
      Top             =   3735
      Width           =   1215
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
      Index           =   13
      Left            =   1305
      MaxLength       =   10
      TabIndex        =   7
      Tag             =   "Fax|T|S|||straba|faxtraba|||"
      Top             =   3675
      Width           =   1455
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
      Index           =   14
      Left            =   1305
      MaxLength       =   40
      TabIndex        =   8
      Tag             =   "E-mail|T|S|||straba|mailtraba|||"
      Top             =   4080
      Width           =   4395
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
      Index           =   3
      Left            =   1305
      MaxLength       =   15
      TabIndex        =   2
      Tag             =   "NIF / CIF|T|N|||straba|niftraba|||"
      Top             =   1620
      Width           =   1920
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
      Left            =   1305
      MaxLength       =   35
      TabIndex        =   3
      Tag             =   "Domicilio|T|S|||straba|domtraba|||"
      Top             =   2025
      Width           =   4395
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
      Height          =   1815
      Index           =   21
      Left            =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Tag             =   "Observaciones|T|S|||straba|observac|||"
      Top             =   4905
      Width           =   5385
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
      Left            =   1305
      MaxLength       =   6
      TabIndex        =   4
      Tag             =   "C.Postal|T|S|||straba|codpobla|||"
      Top             =   2430
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
      Index           =   18
      Left            =   2115
      MaxLength       =   30
      TabIndex        =   5
      Tag             =   "Población|T|S|||straba|pobtraba|||"
      Top             =   2430
      Width           =   3585
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
      Index           =   22
      Left            =   1305
      MaxLength       =   30
      TabIndex        =   6
      Tag             =   "Provincia|T|S|||atraba|protraba|||"
      Top             =   2850
      Width           =   4380
   End
   Begin VB.Frame Frame2 
      Height          =   765
      Index           =   0
      Left            =   225
      TabIndex        =   40
      Top             =   750
      Width           =   13575
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
         Left            =   1095
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "Código|N|N|1|999999|straba|codtraba|000000|S|"
         Top             =   255
         Width           =   855
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
         Index           =   2
         Left            =   3510
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "Nombre|T|N|||straba|nomtraba|||"
         Top             =   255
         Width           =   6375
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre "
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
         Left            =   2595
         TabIndex        =   42
         Top             =   285
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
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
         TabIndex        =   41
         Top             =   285
         Width           =   705
      End
   End
   Begin VB.TextBox text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   25
      Left            =   9600
      TabIndex        =   44
      Top             =   765
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   300
      TabIndex        =   38
      Top             =   6900
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
         TabIndex        =   39
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
      Left            =   12735
      TabIndex        =   37
      Top             =   7005
      Width           =   1035
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
      Left            =   11535
      TabIndex        =   36
      Top             =   7005
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4185
      Top             =   5490
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
      Left            =   12720
      TabIndex        =   43
      Top             =   7005
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   13350
      TabIndex        =   84
      Top             =   150
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
   Begin VB.Label Label20 
      Caption         =   "Cta.Conta."
      Height          =   255
      Left            =   6075
      TabIndex        =   69
      Top             =   3735
      Width           =   735
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   6825
      ToolTipText     =   "Buscar Cta.Contable"
      Top             =   3735
      Width           =   240
   End
   Begin VB.Image imgMail 
      Height          =   240
      Index           =   0
      Left            =   1050
      Top             =   4110
      Width           =   240
   End
   Begin VB.Label Label10 
      Caption         =   "Teléfono"
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
      Left            =   315
      TabIndex        =   64
      Top             =   3285
      Width           =   945
   End
   Begin VB.Label Label12 
      Caption         =   "Móvil"
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
      Left            =   3285
      TabIndex        =   63
      Top             =   3270
      Width           =   645
   End
   Begin VB.Label Label14 
      Caption         =   "Fax"
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
      Left            =   315
      TabIndex        =   62
      Top             =   3705
      Width           =   735
   End
   Begin VB.Label Label16 
      Caption         =   "E-mail"
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
      Left            =   315
      TabIndex        =   61
      Top             =   4110
      Width           =   675
   End
   Begin VB.Label Label5 
      Caption         =   "NIF"
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
      Left            =   330
      TabIndex        =   60
      Top             =   1650
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Domicilio"
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
      Left            =   330
      TabIndex        =   59
      Top             =   2055
      Width           =   945
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
      Left            =   330
      TabIndex        =   58
      Top             =   4605
      Width           =   1470
   End
   Begin VB.Image imgZoom 
      Height          =   240
      Index           =   0
      Left            =   1875
      Tag             =   "-1"
      ToolTipText     =   "Zoom descripción"
      Top             =   4605
      Width           =   240
   End
   Begin VB.Label Label28 
      Caption         =   "Provincia"
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
      Left            =   315
      TabIndex        =   57
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Población"
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
      Left            =   315
      TabIndex        =   56
      Top             =   2475
      Width           =   1035
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
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmManTraba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MONICA                   -+-+
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

Private Const IdPrograma = 8002

'Dim T1 As Single

Public DatosADevolverBusqueda As String    'Tindrà el nº de text que vol que torne, empipat
Public Event DatoSeleccionado(CadenaSeleccion As String)

' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmTraPrev As frmManTrabaPrev ' Trabajadores vista previa
Attribute frmTraPrev.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1

Private WithEvents frmCtas As frmCtasConta 'cuenta contable
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmFPa As frmBasico2 'Formas de Pago
Attribute frmFPa.VB_VarHelpID = -1
Private WithEvents frmMSal As frmManSalarios 'Salarios
Attribute frmMSal.VB_VarHelpID = -1
Private WithEvents frmAlm As frmBasico2 'almacenes propios
Attribute frmAlm.VB_VarHelpID = -1
Private WithEvents frmHor As frmBasico2 'horario de costes de comercial
Attribute frmHor.VB_VarHelpID = -1
' *****************************************************


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

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
'Dim CadAncho() As Boolean  'array, per a quan cridem al form de llínies
Dim Indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim CadB As String

Dim CategAnt As String


Private BuscaChekc As String


Dim NombreAnt As String
Dim IbanAnt As String
Dim BancoAnt As String
Dim SucurAnt As String
Dim DigitoAnt As String
Dim CuentaAnt As String
    
Dim DirecAnt As String
Dim cPostalAnt As String
Dim PoblaAnt As String
Dim ProviAnt As String
Dim NifAnt As String
Dim EMaiAnt As String

Dim Tel1Ant As String
Dim Tel2Ant As String




Private Sub chkAbonos_GotFocus(Index As Integer)
    PonerFocoChk Me.chkAbonos(Index)
End Sub

Private Sub chkAbonos_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkAbonos(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "chkAbonos(" & Index & ")|"
    End If
End Sub

Private Sub chkAbonos_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkAbonos_LostFocus(Index As Integer)
    If Index = 1 And (Modo = 3 Or Modo = 4) Then
        If chkAbonos(Index).Value = 1 Then Text1(25).Text = ""
    End If
End Sub




Private Sub chkEmbarga_GotFocus(Index As Integer)
    PonerFocoChk Me.chkEmbarga(Index)
End Sub

Private Sub chkEmbarga_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkEmbarga(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "chkEmbarga(" & Index & ")|"
    End If
End Sub

Private Sub chkEmbarga_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub




Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BÚSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm2(Me, 1) Then
                
                    If vParamAplic.HayAppAriagro Then PushTrabajador Text1(0).Text, "I"
                
                
                
                    ' *** canviar o llevar el WHERE, repasar codEmpre ****
                    Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
                    'Data1.RecordSource = "Select * from " & NombreTabla & " where codempre = " & codEmpre & Ordenacion
                    ' ***************************************************************
                    PosicionarData
                End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario2(Me, 1) Then
                    
                    '[Monica]13/08/2018: para el caso de modificacion de un socio no lo habiamos lanzado
                    If vParamAplic.HayAppAriagro Then PushTrabajador Text1(0).Text, "M"
                    
                    
                    TerminaBloquear
                    PosicionarData
                End If
            Else
                ModoLineas = 0
            End If
        ' **************************
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



' *** si n'hi han combos a la capçalera ***
Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbLightBlue
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Combo1(Index).BackColor = vbLightBlue Then Combo1(Index).BackColor = vbWhite
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then PrimeraVez = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
    If Modo = 4 Then TerminaBloquear
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim i As Integer
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    
    ' ICONETS DE LA BARRA
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
        .Buttons(8).Image = 10  'Imprimir
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
    
    'cargar IMAGES de busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    'carga IMAGES de mail
    For i = 0 To Me.imgMail.Count - 1
        Me.imgMail(i).Picture = frmPpal.imgListImages16.ListImages(2).Picture
    Next i
    'IMAGES para zoom
    For i = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(i).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next i
    ' Imagenes para ayuda
    For i = 0 To imgAyuda.Count - 1
        imgAyuda(i).Picture = frmPpal.ImageListB.ListImages(10).Picture
    Next i
    
    ' Si hay control de costes el nro de tarjeta es requerido
    If vParamAplic.HayCCostes Then
        Text1(31).Tag = "Nro.Tarjeta|N|N|||straba|idtarjeta|000000||"
    End If
    
    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han llínies *******
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "straba"
    Ordenacion = " ORDER BY codtraba"
    '************************************************
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    '***** canviar el nom de la PK de la capçalera; repasar codEmpre *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where codtraba=-1"
    Data1.Refresh
       
    ' ******* si n'hi han llinies en datagrid *******
'    ReDim CadAncho(DataGridAux.Count) 'redimensione l'array a la quantitat de datagrids
'    CadAncho(0) = False
'    CadAncho(1) = False
'    CadAncho(2) = False
'    CadAncho(4) = False
    
    ModoLineas = 0
       
    ' **** si n'hi ha algun frame que no te datagrids ***
'    CargaFrame 3, False
    ' *************************************************
         
    ' *** si n'hi han combos (capçalera o llínies) ***
    CargaCombo
    ' ************************************************
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1 'búsqueda
        ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
        Text1(0).BackColor = vbLightBlue 'codtraba
        ' ****************************************************************************
    End If
    
    '[Monica]07/05/2018: banco propio que es la categoria
    Label31.Enabled = (vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19)
    Label31.visible = (vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19)
    Text1(35).Enabled = (vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19)
    Text1(35).visible = (vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19)
    
    
End Sub

Private Sub frmTraPrev_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
    
    If CadenaSeleccion <> "" Then
        CadB = "codtraba = " & DBSet(RecuperaValor(CadenaSeleccion, 1), "N")
        
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If

End Sub


Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub


Private Sub LimpiarCampos()
Dim i As Integer

    On Error Resume Next
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
    For i = 0 To Combo1.Count - 1
        Combo1(i).ListIndex = -1
    Next i
    Me.chkAbonos(0).Value = 0
    Me.chkEmbarga(1).Value = 0
    
    ' *** si n'hi han combos a la capçalera ***
    ' *****************************************

    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO s'habiliten, o no, els diversos camps del
'   formulari en funció del modo en que anem a treballar
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim i As Integer, NumReg As Byte
Dim B As Boolean

    On Error GoTo EPonerModo
 
    'Actualisa Iconos Insertar,Modificar,Eliminar
    'ActualizarToolbar Modo, Kmodo
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo, ModoLineas
    
    BuscaChekc = ""
       
    'Modo 2. N'hi han datos i estam visualisant-los
    '=========================================
    'Posem visible, si es formulari de búsqueda, el botó "Regresar" quan n'hi han datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    
    '=======================================
    B = (Modo = 2)
    'Posar Fleches de desplasament visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Només es per a saber que n'hi ha + d'1 registre
    End If
    DesplazamientoVisible B And Data1.Recordset.RecordCount > 1
    
    '---------------------------------------------
    B = Modo <> 0 And Modo <> 2
    CmdCancelar.visible = B
    cmdAceptar.visible = B
       
    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a més neteja els camps Text1
    BloquearText1 Me, Modo
    BloquearCombo Me, Modo
    BloquearChk Me.chkAbonos(0), (Modo = 0 Or Modo = 2 Or Modo = 5)
    BloquearChk Me.chkEmbarga(1), (Modo = 0 Or Modo = 2 Or Modo = 5)
    
    ' ***** bloquejar tots els controls visibles de la clau primaria de la capçalera ***
'    If Modo = 4 Then _
'        BloquearTxt Text1(0), True 'si estic en  modificar, bloqueja la clau primaria
    ' **********************************************************************************
    
    ' **** si n'hi han imagens de buscar en la capçalera *****
    BloquearImgBuscar Me, Modo, ModoLineas
'    BloquearImgFec Me, 25, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
    ' ********************************************************
    
            
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    PonerLongCampos

    
      
'    ' ****** si n'hi han combos a la capçalera ***********************
'    If (Modo = 0) Or (Modo = 2) Or (Modo = 4) Or (Modo = 5) Then
'        Combo1(0).Enabled = False
'        Combo1(0).BackColor = &H80000018 'groc
'    ElseIf (Modo = 1) Or (Modo = 3) Then
'        Combo1(0).Enabled = True
'        Combo1(0).BackColor = &H80000005 'blanc
'    End If
'    ' ****************************************************************
    
    ' *** si n'hi han llínies i algún tab que no te datagrid ***
'    BloquearFrameAux Me, "FrameAux3", Modo, NumTabMto
'    BloquearFrameAux2 Me, "FrameAux3", (Modo <> 5) Or (Modo = 5 And indFrame <> 3) 'frame datos viaje indiv.
    ' ***************************
        
    PonerModoOpcionesMenu (Modo) 'Activar opcions menú según modo
    PonerOpcionesMenu   'Activar opcions de menú según nivell
                        'de permisos de l'usuari

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
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los TEXT1
    PonerLongCamposGnral Me, Modo, 1
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
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
    Toolbar1.Buttons(1).Enabled = B
    Me.mnNuevo.Enabled = B
    
    B = (Modo = 2 And Data1.Recordset.RecordCount > 0) 'And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    Me.mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(3).Enabled = B
    Me.mnEliminar.Enabled = B
    
    'Imprimir
    'Toolbar1.Buttons(12).Enabled = (b Or Modo = 0)
    Toolbar1.Buttons(8).Enabled = B
       
    
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botons de Desplaçament; per a desplaçar-se pels registres de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index, True
    PonerCampos
End Sub




Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento alacenes propios
    Text1(24).Text = RecuperaValor(CadenaSeleccion, 1) 'codalmac
    FormateaCampo Text1(24)
    Text2(24).Text = RecuperaValor(CadenaSeleccion, 2) 'nomalmac
End Sub

Private Sub frmFPa_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento Formas de pago
    Text1(20).Text = RecuperaValor(CadenaSeleccion, 1) 'codforpa
    FormateaCampo Text1(20)
    Text2(20).Text = RecuperaValor(CadenaSeleccion, 2) 'nomforpa
End Sub

Private Sub frmHor_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento de horarios para costes
    Text1(32).Text = RecuperaValor(CadenaSeleccion, 1) 'codhorario
    FormateaCampo Text1(32)
    Text2(32).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmMSal_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento Salarios
    Text1(9).Text = RecuperaValor(CadenaSeleccion, 1) 'codcateg
    FormateaCampo Text1(9)
    Text2(9).Text = RecuperaValor(CadenaSeleccion, 2) 'nomcateg
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(Indice).Text = vCampo
End Sub


Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "Identificador de tarjeta para fichadas de tareas para el proceso de" & vbCrLf & _
                      "costes. " & vbCrLf & vbCrLf & _
                      "Sólo se utiliza si hay control de costes en la aplicación." & vbCrLf & _
                      vbCrLf & vbCrLf
                                            
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
    
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

    Select Case Index
        Case 0
            imgFec(1).Tag = 10 '<===
        
            ' *** repasar si el camp es Text3 o Text1 ***
            If Text1(10).Text <> "" Then frmC.NovaData = Text1(10).Text
            ' ********************************************
        Case 1
            imgFec(1).Tag = 15 '<===
            
            ' *** repasar si el camp es Text3 o Text1 ***
            If Text1(15).Text <> "" Then frmC.NovaData = Text1(15).Text
            ' ********************************************
        Case 2
            imgFec(1).Tag = 16 '<===
            
            ' *** repasar si el camp es Text3 o Text1 ***
            If Text1(16).Text <> "" Then frmC.NovaData = Text1(16).Text
            ' ********************************************
    End Select
    
    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es Text3 o Text1 ***
    PonerFoco Text1(CByte(imgFec(1).Tag)) '<===
    ' ********************************************
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es Text3 o Text1 ***
    Select Case imgFec(1).Tag
        Case 10
            Text1(10).Text = Format(vFecha, "dd/mm/yyyy") '<===
        Case 15
            Text1(15).Text = Format(vFecha, "dd/mm/yyyy") '<===
        Case 16
            Text1(16).Text = Format(vFecha, "dd/mm/yyyy") '<===
    End Select
    ' ********************************************
End Sub
' *****************************************************


Private Sub imgMail_Click(Index As Integer)
'Abrir Outlook para enviar e-mail
Dim dirMail As String

'    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 0: dirMail = Text1(14).Text
    End Select

    If LanzaMailGnral(dirMail) Then espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        Indice = 21
        frmZ.pTitulo = "Obsservaciones Trabajador"
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
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************
    
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 1  'Nou
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
        Case 5  'Búscar
           mnBuscar_Click
        Case 6  'Tots
            mnVerTodos_Click
        Case 8 'Imprimir
            printNou
    End Select
End Sub

Private Sub BotonBuscar()
Dim i As Integer
' ***** Si la clau primaria de la capçalera no es Text1(0), canviar-ho en <=== *****
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonerFoco Text1(0) ' <===
        Text1(0).BackColor = vbLightBlue ' <===
        ' *** si n'hi han combos a la capçalera ***
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbLightBlue
            PonerFoco Text1(kCampo)
        End If
    End If
' ******************************************************************************
End Sub

Private Sub HacerBusqueda()

    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
    If chkVistaPrevia(0) = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        ' *** foco al 1r camp visible de la capçalera que siga clau primaria ***
        PonerFoco Text1(0)
        ' **********************************************************************
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
'    Dim cad As String
'
'    'Cridem al form
'    ' **************** arreglar-ho per a vore lo que es desije ****************
'    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
'    cad = ""
'    cad = cad & ParaGrid(Text1(2), 45, "Nombre")
'    cad = cad & ParaGrid(Text1(0), 10, "Cód.")
'    cad = cad & ParaGrid(Text1(3), 15, "NIF")
'    cad = cad & ParaGrid(Text1(11), 15, "Teléfono")
'    cad = cad & ParaGrid(Text1(12), 15, "Móvil")
'    If cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vtabla = NombreTabla
'        frmB.vSQL = CadB
'        HaDevueltoDatos = False
'        frmB.vDevuelve = "1|" '*** els camps que volen que torne ***
'        frmB.vTitulo = "Trabajadores" ' ***** repasa açò: títol de BuscaGrid *****
'        frmB.vSelElem = 0
'
'        frmB.Show vbModal
'        Set frmB = Nothing
'        'Si ha posat valors i tenim que es formulari de búsqueda llavors
'        'tindrem que tancar el form llançant l'event
'        If HaDevueltoDatos Then
'            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                cmdRegresar_Click
'        Else   'de ha retornat datos, es a decir NO ha retornat datos
'            PonerFoco Text1(kCampo)
'        End If
'    End If
    
    Set frmTraPrev = New frmManTrabaPrev
    
    frmTraPrev.cWhere = CadB
    frmTraPrev.DatosADevolverBusqueda = "0|1|2|"
    frmTraPrev.Show vbModal
    
    Set frmTraPrev = Nothing
    
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


Private Sub BotonVerTodos()
'Vore tots
    LimpiarCampos 'Neteja els Text1
    CadB = ""
    
    If chkVistaPrevia(0).Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        ' *** canviar o llevar, si cal, el WHERE; repasar codEmpre ***
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        'CadenaConsulta = "Select * from " & NombreTabla & " where codempre = " & codEmpre & Ordenacion
        ' ******************************************
        PonerCadenaBusqueda
        ' *** si n'hi han llínies sense grids ***
'        CargaFrame 0, True
        ' ************************************
    End If
End Sub


Private Sub BotonAnyadir()

    LimpiarCampos 'Huida els TextBox
    PonerModo 3
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la capçalera *******
    Text1(0).Text = SugerirCodigoSiguienteStr("straba", "codtraba")
    FormateaCampo Text1(0)
       
    CategAnt = ""
    
    Text1(33).Text = Format(0, "##0.0000")
       
    PonerFoco Text1(0) '*** 1r camp visible que siga PK ***
    ' ***********************************************************
    
    ' *** si n'hi han camps de descripció a la capçalera ***
    PosarDescripcions
    ' ******************************************************
    
End Sub


Private Sub BotonModificar()

    PonerModo 4

    CategAnt = Text1(9).Text
    
    
    '[Monica]17/08/2018: para la parte de aripush
    NombreAnt = Text1(2).Text
    IbanAnt = Text1(34).Text
    BancoAnt = Text1(1).Text
    SucurAnt = Text1(19).Text
    DigitoAnt = Text1(8).Text
    CuentaAnt = Text1(7).Text
    
    DirecAnt = Text1(4).Text
    cPostalAnt = Text1(5).Text
    PoblaAnt = Text1(18).Text
    ProviAnt = Text1(22).Text
    NifAnt = Text1(3).Text
    EMaiAnt = Text1(14).Text

    Tel1Ant = Text1(11).Text
    Tel2Ant = Text1(12).Text
    
    

    ' *** bloquejar els camps visibles de la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    ' *************************************************************************
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(2)
    ' *********************************************************
End Sub


Private Sub BotonEliminar()
Dim cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    If Not SepuedeBorrar Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************

    ' *************** canviar la pregunta ****************
    cad = "¿Seguro que desea eliminar el Trabajador?"
    cad = cad & vbCrLf & "Código: " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
    cad = cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(1)
    ' **************************************************************************
    
    'borrem
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ' **** repasar si es diu Data1 l'adodc de la capçalera ***
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
        ' ********************************************************
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Trabajador", Err.Description
End Sub


Private Sub PonerCampos()
Dim i As Integer
Dim CodPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la capçalera
    

    ' *** si n'hi han llínies sense datagrid ***
'    CargaFrame 3, True
    ' ***************************************
    
    ' ************* configurar els camps de les descripcions de la capçalera *************
'    txtAux2(22).Text = PonerNombreDeCod(txtAux(22), "poblacio", "despobla", "codpobla", "N")

    PosarDescripcions

'    codPobla = DBLet(Data1.Recordset!codPobla, "T")
'    DatosPoblacion codPobla, desPobla, CPostal, desProvi, desPais
'    text1(5).Text = codPobla 'Devuelve el campo formateado
'    text2(5).Text = desPobla
''    text1(8).Text = CPostal
'    text2(1).Text = desProvi
'    text2(2).Text = desPais
'
'    text2(7).Text = PonerNombreDeCod(text1(7), "activida", "desactiv")
'    text2(8).Text = PonerNombreDeCod(text1(8), "grupempr", "desgrupo", "codgrupo", "N")
    ' ********************************************************************************
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
End Sub


Private Sub cmdCancelar_Click()
Dim i As Integer
Dim V

    Select Case Modo
        Case 1, 3 'Búsqueda, Insertar
                LimpiarCampos
                If Data1.Recordset.EOF Then
                    PonerModo 0
                Else
                    PonerModo 2
                    PonerCampos
                End If
                ' *** foco al primer camp visible de la capçalera ***
                PonerFoco Text1(0)
                ' ***************************************************

        Case 4  'Modificar
                TerminaBloquear
                PonerModo 2
                PonerCampos
                ' *** primer camp visible de la capçalera ***
                PonerFoco Text1(0)
                ' *******************************************
        
    End Select
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean
Dim SQL As String
Dim cadMen As String
Dim cta As String
Dim TipoForp As String

'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOk = False
    B = CompForm2(Me, 1)
    If Not B Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
        'comprobar si existe ya el cod. del campo clave primaria
        If ExisteCP(Text1(0)) Then B = False
    End If
    
    If B And (Modo = 3 Or Modo = 4) Then
        TipoForp = ""
        TipoForp = DevuelveDesdeBDNew(cAgro, "forpago", "tipoforp", "codforpa", Text1(20).Text, "N")
        If CByte(TipoForp) = 1 Then ' transferencia
            If Text1(1).Text = "" Or Text1(19).Text = "" Or Text1(8).Text = "" Or Text1(7).Text = "" Then
                Text1(1).Text = ""
                Text1(19).Text = ""
                Text1(8).Text = ""
                Text1(7).Text = ""
                B = False
                cadMen = "El trabajador no tiene asignada cuenta bancaria."
            Else
                cta = Format(Text1(1).Text, "0000") & Format(Text1(19).Text, "0000") & Format(Text1(8).Text, "00") & Format(Text1(7).Text, "0000000000")
                If Val(ComprobarCero(cta)) = 0 Then
                    cadMen = "El trabajador no tiene asignada cuenta bancaria."
                    B = False
                End If
                If Not Comprueba_CC(cta) Then
                    cadMen = "La cuenta bancaria del trabajador no es correcta."
                    B = False
                Else
    '                '[Monica]20/11/2013: añadimos el tema de la comprobacion del IBAN
    '                If Not Comprueba_CC_IBAN(cta, Text1(42).Text) Then
    '                    cadMen = "La cuenta IBAN del cliente no es correcta. ¿ Desea continuar ?."
    '                    If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
    '                        b = True
    '                    Else
    '                        PonerFoco Text1(42)
    '                        b = False
    '                    End If
    '                End If
    
    '       sustituido por lo de David
                    BuscaChekc = ""
                    If Me.Text1(34).Text <> "" Then BuscaChekc = Mid(Text1(34).Text, 1, 2)
                        
                    If DevuelveIBAN2(BuscaChekc, cta, cta) Then
                        If Me.Text1(34).Text = "" Then
                            If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.Text1(34).Text = BuscaChekc & cta
                        Else
                            If Mid(Text1(34).Text, 3) <> cta Then
                                cta = "Calculado : " & BuscaChekc & cta
                                cta = "Introducido: " & Me.Text1(34).Text & vbCrLf & cta & vbCrLf
                                cta = "Error en codigo IBAN" & vbCrLf & cta & "Continuar?"
                                If MsgBox(cta, vbQuestion + vbYesNo) = vbNo Then
                                    PonerFoco Text1(34)
                                    B = False
                                End If
                            End If
                        End If
                    End If
                End If
                    
            
            End If
            If cadMen <> "" Then
                MsgBox cadMen, vbExclamation
            End If
        End If
    End If
    
    'control de costes
    If B And vParamAplic.HayCCostes Then
        If Modo = 3 Or Modo = 4 Then
            SQL = "select count(*) from straba where codtraba <> " & DBSet(Text1(0).Text, "N")
            If Text1(31).Text <> "" Then
                SQL = SQL & " and idtarjeta = " & DBSet(Text1(31).Text, "N")
            Else
                SQL = SQL & " and idtarjeta is null "
            End If
    
            If DevuelveValor(SQL) <> 0 Then
                If MsgBox("Hay otro trabajador con el mismo Nro.Tarjeta asignado." & vbCrLf & vbCrLf & "               ¿Desea Continuar?.", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                    B = False
                End If
            End If
        End If
    End If
    
    If B And vParamAplic.HayCCostes Then
        If Modo = 3 Or Modo = 4 Then
            If Text1(32).Text = "" Then
                MsgBox "Debe introducir obligatoriamente un tipo de horario para el trabajador.", vbExclamation
                PonerFoco Text1(32)
                B = False
            Else
                Text2(32).Text = PonerNombreDeCod(Text1(32), "cchorario", "descripc")
                If Text2(32).Text = "" Then
                    MsgBox "Código de Horario no existe. Reintroduzca.", vbExclamation
                    PonerFoco Text1(32)
                    B = False
                End If
            End If
        End If
    End If
    
    DatosOk = B
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub PosicionarData()
Dim cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    cad = "(codtraba=" & Text1(0).Text & ")"
    ' ***************************************
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarData(Data1, cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
    ' ***********************************************************************************
End Sub


Private Function Eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar

    conn.BeginTrans
'    ' ***** canviar el nom de la PK de la capçalera, repasar codEmpre *******
'    vWhere = " WHERE codclien=" & Data1.Recordset!Codclien
'        ' ***********************************************************************
'
'    ' ***** elimina les llínies ****
'    Conn.Execute "DELETE FROM destinos " & vWhere
'
'    ' *******************************
        
    'Eliminar la CAPÇALERA
    vWhere = " WHERE codtraba=" & Data1.Recordset!CodTraba
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
Dim campo2 As String
Dim SQL As String



    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    Select Case Index
        Case 0 'cod cliente
            PonerFormatoEntero Text1(0)

        Case 2 'NOMBRE
            Text1(Index).Text = UCase(Text1(Index).Text)
        
        Case 3 'NIF
            Text1(Index).Text = UCase(Text1(Index).Text)
            ValidarNIF Text1(Index).Text
                
                
        Case 9 'CATEGORIA
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "salarios", "nomcateg")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Salario: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmMSal = New frmManSalarios
                        frmMSal.DatosADevolverBusqueda = "0|1|"
                        frmMSal.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmMSal.Show vbModal
                        Set frmMSal = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                Else ' traemos el porcentaje de irpf y de seguridad social
                    If Modo = 3 Or Modo = 4 Then
                        campo2 = "dtosegso"
                        SQL = DevuelveDesdeBDNew(cAgro, "salarios", "dtosirpf", "codcateg", Text1(9).Text, "N", campo2)
                        If SQL <> "" Then
                            If Text1(9).Text <> CategAnt Then Text1(25).Text = Format(ImporteSinFormato(SQL), "##0.00")
                            If Text1(9).Text <> CategAnt Then Text1(26).Text = Format(ImporteSinFormato(campo2), "##0.00")
                            If Text1(9).Text <> CategAnt Then Text1(27).Text = DevuelveDesdeBDNew(cAgro, "salarios", "pluscapataz", "codcateg", Text1(9).Text, "N")
                            If Text1(9).Text <> CategAnt Then Text1(27).Text = Format(ImporteSinFormato(Text1(27).Text), "###,##0.00")
                            If Text1(9).Text <> CategAnt Then Text1(29).Text = DevuelveDesdeBDNew(cAgro, "salarios", "dtoreten", "codcateg", Text1(9).Text, "N")
                            If Text1(9).Text <> CategAnt Then Text1(29).Text = Format(ImporteSinFormato(Text1(29).Text), "###,##0.00")
                        End If
                    End If
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 20 'FORMA DE PAGO
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "forpago", "nomforpa")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Forma de Pago: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmFPa = New frmComercial
                        frmFPa.DatosADevolverBusqueda = "0|1|"
                        frmFPa.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmFPa.Show vbModal
                        Set frmFPa = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 10, 15, 16 'Fechas
            PonerFormatoFecha Text1(Index)
            
        Case 23 'cuenta contable
            If Text1(Index).Text = "" Then Exit Sub
            If Modo = 3 Then
                Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo, "") 'text1(0).Text)
            Else
                Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo, Text1(2).Text)
            End If
            
'        Case 23, 26 'porcentajes de comision
'            cadMen = TransformaPuntosComas(Text1(Index).Text)
'            Text1(Index).Text = Format(cadMen, "##0.00")
'
'        Case 25 'tipo de movimiento
'            If Text1(Index).Text <> "" Then Text1(Index).Text = UCase(Text1(Index).Text)
'
        Case 1, 19 'ENTIDAD Y SUCURSAL BANCARIA
            PonerFormatoEntero Text1(Index)
          
          
        Case 24 'ALMACENES PROPIOS
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "salmpr", "nomalmac")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el almacén: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmAlm = New frmComercial
                        frmAlm.DatosADevolverBusqueda = "0|1|"
                        frmAlm.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmAlm.Show vbModal
                        Set frmFPa = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
          
       Case 25, 26 ' dtoirpf y dto seguridad social
            PonerFormatoDecimal Text1(Index), 4
          
       Case 27 ' % antiguedad
            PonerFormatoDecimal Text1(Index), 4
            
       Case 28 ' plus del capataz
            PonerFormatoDecimal Text1(Index), 3
       
       Case 29 ' % dtoretencion para Picassent
            PonerFormatoDecimal Text1(Index), 4
    
       Case 30 'CODIGO DE ASESORIA PARA PICASSENT
            PonerFormatoEntero Text1(Index)
    
       Case 32 'CODIGO DE HORARIO
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "cchorario", "descripc")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Horario " & Text1(Index).Text & ". Reintroduzca." & vbCrLf
                    MsgBox cadMen, vbExclamation
                    PonerFoco Text1(Index)
                End If
            End If
            
       Case 33 ' precio hora coste
            PonerFormatoDecimal Text1(Index), 7
            
       Case 34 ' codigo de iban
            Text1(Index).Text = UCase(Text1(Index).Text)
            
       Case 35
            PonerFormatoEntero Text1(Index)
    End Select
    
    '[Monica]: calculo del iban si no lo ponen
    If Index = 1 Or Index = 19 Or Index = 8 Or Index = 7 Then
        Dim cta As String
        Dim CC As String
        If Text1(1).Text <> "" And Text1(19).Text <> "" And Text1(8).Text <> "" And Text1(7).Text <> "" Then
            
            cta = Format(Text1(1).Text, "0000") & Format(Text1(19).Text, "0000") & Format(Text1(8).Text, "00") & Format(Text1(7).Text, "0000000000")
            If Len(cta) = 20 Then
    '        Text1(42).Text = Calculo_CC_IBAN(cta, Text1(42).Text)
    
                If Text1(34).Text = "" Then
                    'NO ha puesto IBAN
                    If DevuelveIBAN2("ES", cta, cta) Then Text1(34).Text = "ES" & cta
                Else
                    CC = CStr(Mid(Text1(34).Text, 1, 2))
                    If DevuelveIBAN2(CStr(CC), cta, cta) Then
                        If Mid(Text1(34).Text, 3) <> cta Then
                            
                            MsgBox "Codigo IBAN distinto del calculado [" & CC & cta & "]", vbExclamation
                        End If
                    End If
                End If
                
                
            End If
        End If
    End If
    
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 21 Then
        If KeyAscii = teclaBuscar Then
            If Modo = 1 Or Modo = 3 Or Modo = 4 Then
                Select Case Index
                    Case 9: KEYBusqueda KeyAscii, 0 'categoria
                    Case 24: KEYBusqueda KeyAscii, 3 'almacen
                    Case 20: KEYBusqueda KeyAscii, 2 'forma de pago
                    Case 32: KEYBusqueda KeyAscii, 4 'horario
                    Case 10: KEYFecha KeyAscii, 0 'fecha de alta
                    Case 15: KEYFecha KeyAscii, 1 'fecha de baja
                    Case 16: KEYFecha KeyAscii, 2 'fecha antiguedad
                End Select
            End If
        Else
            KEYpress KeyAscii
        End If
    Else
        If Text1(21).Text = "" And KeyAscii = teclaBuscar Then
            imgZoom_Click 0
        Else
            KEYpress KeyAscii
        End If
    End If
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

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFec_Click (Indice)
End Sub
Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
End Sub

' **** si n'hi han camps de descripció a la capçalera ****
Private Sub PosarDescripcions()
Dim NomEmple As String

    On Error GoTo EPosarDescripcions
    
    Text2(9).Text = PonerNombreDeCod(Text1(9), "salarios", "nomcateg", "codcateg", "N")
    
    Text2(20).Text = PonerNombreDeCod(Text1(20), "forpago", "nomforpa", "codforpa", "N")
    If vParamAplic.NumeroConta <> 0 Then
        Text2(23).Text = PonerNombreCuenta(Text1(23), Modo)
    End If
    Text2(24).Text = PonerNombreDeCod(Text1(24), "salmpr", "nomalmac", "codalmac", "N")
   
    Text2(32).Text = PonerNombreDeCod(Text1(32), "cchorario", "descripc", "codhorario", "N")
       
   
EPosarDescripcions:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo descripciones", Err.Description
End Sub
' ************************************************************





' ***** Si n'hi han combos *****
' per a seleccionar la opcio del combo quan estem modificant; només per a "si" i "no"
'Private Sub SelComboBool(valor As Integer, combo As ComboBox)
'Private Sub SelComboBool(valor, combo As ComboBox)
'    Dim i As Integer
'    Dim j As Integer
'
'    i = valor
'    For j = 0 To combo.ListCount - 1
'        If combo.ItemData(j) = i Then
'            combo.ListIndex = j
'            Exit For
'        End If
'    Next j
'End Sub
' ********************************



' *** si n'hi han formularis de buscar codi a les llínies ***
Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
    
    Select Case Index
       Case 0 'salarios
            Set frmMSal = New frmManSalarios
            frmMSal.DatosADevolverBusqueda = "0|1|"
            frmMSal.CodigoActual = Text1(9).Text
            frmMSal.Show vbModal
            Set frmMSal = Nothing
            PonerFoco Text1(9)
        
       Case 1 'Cuentas Contables (de contabilidad)
            If vParamAplic.NumeroConta = 0 Then Exit Sub
            
            Indice = Index + 22
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = Text1(Indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco Text1(Indice)
       
       Case 2 'formas de pago

            Set frmFPa = New frmBasico2
            
            AyudaFPagoCom frmFPa, Text1(20).Text
            
            Set frmFPa = Nothing

            PonerFoco Text1(20)
       
       Case 3 'almacén

            Set frmAlm = New frmBasico2
            
            AyudaAlmacenCom frmAlm, Text1(24).Text
            
            Set frmAlm = Nothing

            PonerFoco Text1(24)
       
       Case 4 'horario de coste
            Set frmHor = New frmBasico2
            
            AyudaHorarioCom frmHor, Text1(32).Text
            
            Set frmHor = Nothing

            PonerFoco Text1(32)
       
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub



Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    Text1(23).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    FormateaCampo Text1(23)
    Text2(23).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

' *********************************************************************************
Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " codclien=" & Val(Text1(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function

Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "straba"
        .Informe2 = "rManTraba.rpt"
        If CadB <> "" Then
            '.cadRegSelec = Replace(SQL2SF(CadB), "clientes", "clientes_1")
            .cadRegSelec = SQL2SF(CadB)
        Else
            .cadRegSelec = ""
        End If
        ' *** repasar el nom de l'adodc ***
        '.cadRegActua = Replace(POS2SF(Data1, Me), "clientes", "clientes_1")
        .cadRegActua = POS2SF(Data1, Me)
        '[Monica]13/07/2012: falla si hay un solo registro seleccionado y apretamos registros buscados
        If Data1.Recordset.RecordCount = 1 Then .cadRegSelec = .cadRegActua
        
        ' *** repasar codEmpre ***
        .cadTodosReg = ""
        '.cadTodosReg = "{itinerar.codempre} = " & codEmpre
        ' *** repasar si li pose ordre o no ****
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|pOrden={straba.codtraba}|"
        '.OtrosParametros2 = "pEmpresa='" & vEmpresa.NomEmpre & "'|"
        ' *** posar el nº de paràmetres que he posat en OtrosParametros2 ***
        '.NumeroParametros2 = 1
        .NumeroParametros2 = 2
        ' ******************************************************************
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = False
        .SubInformeConta = ""
        .Show vbModal
    End With
End Sub


'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGridAux_GotFocus(Index As Integer)
'  WheelHook DataGridAux(Index)
'End Sub
'Private Sub DataGridAux_LostFocus(Index As Integer)
'  WheelUnHook
'End Sub

' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 0 To Combo1.Count - 1
        Combo1(i).Clear
    Next i
    
    Combo1(0).AddItem "Socio"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Trabajador"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
    Combo1(1).AddItem "Campo"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Almacén"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    
    
End Sub

Private Function SepuedeBorrar() As Boolean
Dim SQL As String

    SepuedeBorrar = False
    
    ' *** si cal comprovar alguna cosa abans de borrar ***
    SQL = ""
    SQL = DevuelveDesdeBDNew(cAgro, "horas", "codtraba", "codtraba", Data1.Recordset!CodTraba, "N")
    If SQL <> "" Then
        MsgBox "No puede borrar el trabajador porque tiene horas asignadas.", vbExclamation
        Exit Function
    End If
    ' ****************************************************
    
    SepuedeBorrar = True
    
End Function

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub


Private Function PushTrabajador(Codigo As String, Operacion As String) As Boolean
Dim SQL As String


    On Error GoTo ePushTrabajador

    PushTrabajador = False

'    If Not AbrirConexionApp Then Exit Function
    
    If Not ExisteAripush Then Exit Function

    Select Case Operacion
        Case "I" ' insertar
            SQL = DevuelveDesdeBDNew(cAgro, "aripush.usuariospush", "login", "login", "C" & CLng(Text1(0).Text), "T")
            If SQL = "" Then
                SQL = "insert into aripush.usuariospush (nif,nombre,comunId,ariagroId,tiendaId,gasolineraId,telefoniaId,tratamientosId,login,"
                SQL = SQL & "password,email,playerId,direccion,codPostal,poblacion,provincia,telefono1,telefono2,iban, solomensajes, estrabajador)"
                SQL = SQL & " values ("
                SQL = SQL & DBSet(Text1(3), "T") & "," ' nif
                SQL = SQL & DBSet(Text1(2), "T") & "," 'nombre
                SQL = SQL & ValorNulo & "," 'codtraba
                SQL = SQL & ValorNulo & "," 'ariagroid
                SQL = SQL & ValorNulo & "," 'tiendaid
                SQL = SQL & ValorNulo & "," 'gasolineraid
                SQL = SQL & ValorNulo & "," 'telefoniaid
                SQL = SQL & ValorNulo & "," 'tratamientosid
                SQL = SQL & DBSet("C" & CLng(Text1(0)), "T") & "," 'login
                SQL = SQL & DBSet(Text1(3), "T") & "," 'password el nif
                SQL = SQL & DBSet(Text1(14), "T") & "," 'email
                SQL = SQL & ValorNulo & "," 'playerid
                SQL = SQL & DBSet(Text1(4), "T") & "," 'dirsocio (direccion fiscal)
                SQL = SQL & DBSet(Text1(5), "T") & "," 'codigo postal
                SQL = SQL & DBSet(Text1(18), "T") & "," 'poblacion
                SQL = SQL & DBSet(Text1(22), "T") & "," 'provincia
                SQL = SQL & DBSet(Text1(11), "T") & "," 'telefono1
                SQL = SQL & DBSet(Text1(12), "T") & "," 'telefono2
                SQL = SQL & DBSet(Text1(34).Text & Text1(1).Text & Text1(19).Text & Text1(8).Text & Text1(7).Text, "T") & ","
                SQL = SQL & "1,1)"
            Else
                SQL = "update aripush.usuariospush set "
                SQL = SQL & " nif = " & DBSet(Text1(3), "T")
                SQL = SQL & ",nombre = " & DBSet(Text1(2).Text, "T")
                SQL = SQL & ",email = " & DBSet(Text1(14).Text, "T")
                SQL = SQL & ",direccion = " & DBSet(Text1(4), "T")  'dirsocio (direccion fiscal)
                SQL = SQL & ",codpostal = " & DBSet(Text1(5), "T")  'codigo postal
                SQL = SQL & ",poblacion = " & DBSet(Text1(18), "T")  'poblacion
                SQL = SQL & ",provincia = " & DBSet(Text1(22), "T")  'provincia
                SQL = SQL & ",telefono1 = " & DBSet(Text1(11), "T")  'telefono1
                SQL = SQL & ",telefono2 = " & DBSet(Text1(12), "T")  'telefono2
                SQL = SQL & ",iban = " & DBSet(Text1(34).Text & Text1(1).Text & Text1(19).Text & Text1(8).Text & Text1(7).Text, "T")
                SQL = SQL & ",solomensajes = 1"
                SQL = SQL & ",estrabajador = 1"
                SQL = SQL & " where login = " & DBSet("C" & CLng(Text1(0).Text), "T")
            End If

            conn.Execute SQL
            
        
        Case "M" ' modificar
            If NombreAnt <> Text1(2).Text Or IbanAnt <> Text1(34).Text Or BancoAnt <> Text1(1).Text Or SucurAnt <> Text1(19).Text Or _
               DigitoAnt <> Text1(8).Text Or CuentaAnt <> Text1(7).Text Or DirecAnt <> Text1(4).Text Or cPostalAnt <> Text1(5).Text Or _
               PoblaAnt <> Text1(18).Text Or ProviAnt <> Text1(22).Text Or NifAnt <> Text1(3).Text Or EMaiAnt <> Text1(14).Text Or _
               Tel1Ant <> Text1(11).Text Or Tel2Ant <> Text1(12).Text Then

                    SQL = "update aripush.usuariospush set "
                    SQL = SQL & " nif = " & DBSet(Text1(3), "T")
                    SQL = SQL & ",nombre = " & DBSet(Text1(2).Text, "T")
                    SQL = SQL & ",email = " & DBSet(Text1(14).Text, "T")
                    SQL = SQL & ",direccion = " & DBSet(Text1(4), "T") 'dirsocio (direccion fiscal)
                    SQL = SQL & ",codpostal = " & DBSet(Text1(5), "T") 'codigo postal
                    SQL = SQL & ",poblacion = " & DBSet(Text1(18), "T") 'poblacion
                    SQL = SQL & ",provincia = " & DBSet(Text1(22), "T") 'provincia
                    SQL = SQL & ",telefono1 = " & DBSet(Text1(11), "T") 'telefono1
                    SQL = SQL & ",telefono2 = " & DBSet(Text1(12), "T") 'telefono2
                    SQL = SQL & ",iban = " & DBSet(Text1(34).Text & Text1(1).Text & Text1(19).Text & Text1(8).Text & Text1(7).Text, "T")
                    SQL = SQL & " where login = " & DBSet("C" & CLng(Text1(0).Text), "T")
                    
                    conn.Execute SQL
            End If
        
        Case "E" ' eliminar
    End Select

'    CerrarConexionApp = True
    
    Exit Function

ePushTrabajador:
    MuestraError Err.Number, "Push Trabajador en App", Err.Description
End Function


Private Function ExisteAripush() As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset

    On Error GoTo eExisteAripush

    SQL = "select * from aripush.usuariospush "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Set Rs = Nothing
    
    ExisteAripush = True
    Exit Function

eExisteAripush:
    ExisteAripush = False
End Function


