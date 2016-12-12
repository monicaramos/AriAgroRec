VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManCoope 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cooperativas"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   14610
   Icon            =   "frmManCoope.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   14610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   120
      TabIndex        =   50
      Top             =   30
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   51
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
               Object.Tag             =   "0"
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
   Begin VB.Frame Frame4 
      Caption         =   "Factura Retirada"
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
      Height          =   825
      Left            =   7770
      TabIndex        =   45
      Top             =   5280
      Width           =   6675
      Begin VB.TextBox txtAux 
         BeginProperty Font 
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
         Left            =   5910
         MaxLength       =   3
         TabIndex        =   21
         Tag             =   "Tipo Mov Fra.Retirada Bodega|T|N|||rcoope|codtipomfacbod|||"
         Top             =   270
         Width           =   585
      End
      Begin VB.TextBox txtAux 
         BeginProperty Font 
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
         Left            =   4380
         MaxLength       =   3
         TabIndex        =   20
         Tag             =   "Tipo Mov Fra.Retirada Almazara|T|N|||rcoope|codtipomfacalmz|||"
         Top             =   270
         Width           =   555
      End
      Begin VB.Label Label14 
         Caption         =   "Bodega"
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
         Left            =   5130
         TabIndex        =   47
         Top             =   285
         Width           =   720
      End
      Begin VB.Label Label13 
         Caption         =   "Almazara"
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
         Left            =   3450
         TabIndex        =   46
         Top             =   255
         Width           =   900
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Liquidación"
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
      Height          =   795
      Left            =   7770
      TabIndex        =   39
      Top             =   4440
      Width           =   6675
      Begin VB.TextBox txtAux 
         BeginProperty Font 
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
         Left            =   5910
         MaxLength       =   3
         TabIndex        =   19
         Tag             =   "Tipo Mov Anticipos Bodega|T|N|||rcoope|codtipomliqbod|||"
         Top             =   270
         Width           =   585
      End
      Begin VB.TextBox txtAux 
         BeginProperty Font 
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
         Left            =   4395
         MaxLength       =   3
         TabIndex        =   18
         Tag             =   "Tipo Mov Liquidación Almazara|T|N|||rcoope|codtipomliqalmz|||"
         Top             =   270
         Width           =   585
      End
      Begin VB.TextBox txtAux 
         BeginProperty Font 
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
         Left            =   1020
         MaxLength       =   3
         TabIndex        =   16
         Tag             =   "Tipo Mov Liquidación|T|N|||rcoope|codtipomliq|||"
         Top             =   270
         Width           =   585
      End
      Begin VB.TextBox txtAux 
         BeginProperty Font 
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
         Left            =   2685
         MaxLength       =   3
         TabIndex        =   17
         Tag             =   "Tipo Mov Liquidación V.Campo|T|N|||rcoope|codtipomliqvc|||"
         Top             =   270
         Width           =   585
      End
      Begin VB.Label Label16 
         Caption         =   "Bodega"
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
         Left            =   5115
         TabIndex        =   49
         Top             =   285
         Width           =   720
      End
      Begin VB.Label Label11 
         Caption         =   "Almazara"
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
         Left            =   3450
         TabIndex        =   43
         Top             =   285
         Width           =   930
      End
      Begin VB.Label Label8 
         Caption         =   "Socios"
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
         TabIndex        =   42
         Top             =   285
         Width           =   630
      End
      Begin VB.Label Label10 
         Caption         =   "V.Campo"
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
         Left            =   1740
         TabIndex        =   40
         Top             =   285
         Width           =   945
      End
   End
   Begin VB.TextBox txtAux 
      BeginProperty Font 
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
      Left            =   12660
      MaxLength       =   10
      TabIndex        =   9
      Tag             =   "Fax|T|S|||rcoope|faxcoope|||"
      Top             =   2400
      Width           =   1740
   End
   Begin VB.TextBox txtAux 
      BeginProperty Font 
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
      Left            =   8805
      MaxLength       =   10
      TabIndex        =   8
      Tag             =   "Teléfono|T|S|||rcoope|telcoope|||"
      Top             =   2400
      Width           =   1725
   End
   Begin VB.TextBox txtAux 
      BeginProperty Font 
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
      Left            =   8805
      MaxLength       =   30
      TabIndex        =   10
      Tag             =   "Web|T|S|||rcoope|wwwcoope|||"
      Top             =   2805
      Width           =   5580
   End
   Begin VB.TextBox txtAux 
      BeginProperty Font 
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
      Left            =   8805
      MaxLength       =   40
      TabIndex        =   11
      Tag             =   "Mail|T|S|||rcoope|maicoope|||"
      Top             =   3225
      Width           =   5580
   End
   Begin VB.Frame Frame2 
      Caption         =   "Anticipos"
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
      Height          =   765
      Left            =   7770
      TabIndex        =   33
      Top             =   3690
      Width           =   6690
      Begin VB.TextBox txtAux 
         BeginProperty Font 
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
         Left            =   5910
         MaxLength       =   3
         TabIndex        =   15
         Tag             =   "Tipo Mov Anticipos Bodega|T|N|||rcoope|codtipomantbod|||"
         Top             =   240
         Width           =   585
      End
      Begin VB.TextBox txtAux 
         BeginProperty Font 
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
         Left            =   4395
         MaxLength       =   3
         TabIndex        =   14
         Tag             =   "Tipo Mov Anticipos Almazara|T|N|||rcoope|codtipomantalmz|||"
         Top             =   240
         Width           =   585
      End
      Begin VB.TextBox txtAux 
         BeginProperty Font 
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
         Left            =   2685
         MaxLength       =   3
         TabIndex        =   13
         Tag             =   "Tipo Mov Anticipos V.Campo|T|N|||rcoope|codtipomantvc|||"
         Top             =   240
         Width           =   585
      End
      Begin VB.TextBox txtAux 
         BeginProperty Font 
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
         Left            =   1050
         MaxLength       =   3
         TabIndex        =   12
         Tag             =   "Tipo Mov Anticipos|T|N|||rcoope|codtipomant|||"
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label15 
         Caption         =   "Bodega"
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
         Left            =   5130
         TabIndex        =   48
         Top             =   270
         Width           =   840
      End
      Begin VB.Label Label12 
         Caption         =   "Almazara"
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
         Left            =   3450
         TabIndex        =   44
         Top             =   270
         Width           =   1260
      End
      Begin VB.Label Label9 
         Caption         =   "V.Campo "
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
         Left            =   1770
         TabIndex        =   41
         Top             =   270
         Width           =   945
      End
      Begin VB.Label Label4 
         Caption         =   "Socios"
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
         TabIndex        =   38
         Top             =   270
         Width           =   765
      End
   End
   Begin VB.TextBox txtAux 
      BeginProperty Font 
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
      Left            =   8820
      MaxLength       =   9
      TabIndex        =   3
      Tag             =   "CIF|T|N|||rcoope|cifcoope|||"
      Top             =   780
      Width           =   1455
   End
   Begin VB.TextBox txtAux 
      BeginProperty Font 
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
      Left            =   13110
      MaxLength       =   6
      TabIndex        =   4
      Tag             =   "Código Postal|T|N|||rcoope|codposta|||"
      Top             =   765
      Width           =   1260
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
      Index           =   2
      Left            =   6210
      MaxLength       =   6
      TabIndex        =   2
      Tag             =   "Porc.Gasto|N|S|||rcoope|porcgast|##0.00||"
      Top             =   4410
      Width           =   690
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
      Left            =   240
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "Código|N|N|1|99|rcoope|codcoope|00|S|"
      Top             =   4425
      Width           =   795
   End
   Begin VB.TextBox txtAux 
      BeginProperty Font 
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
      Left            =   8805
      MaxLength       =   30
      TabIndex        =   7
      Tag             =   "Provincia|T|N|||rcoope|procoope|||"
      Top             =   1995
      Width           =   5580
   End
   Begin VB.TextBox txtAux 
      BeginProperty Font 
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
      Left            =   8805
      MaxLength       =   30
      TabIndex        =   6
      Tag             =   "Población|T|N|||rcoope|pobcoope|||"
      Top             =   1590
      Width           =   5580
   End
   Begin VB.TextBox txtAux 
      BeginProperty Font 
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
      Left            =   8805
      MaxLength       =   35
      TabIndex        =   5
      Tag             =   "Domicilio|T|N|||rcoope|domcoope|||"
      Top             =   1185
      Width           =   5580
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
      Left            =   12120
      TabIndex        =   22
      Top             =   6240
      Visible         =   0   'False
      Width           =   1035
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
      Left            =   13350
      TabIndex        =   23
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   1215
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "Nombre|T|N|||rcoope|nomcoope|||"
      Text            =   "1234567890123456789012345678901234567890"
      Top             =   4425
      Width           =   4905
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmManCoope.frx":000C
      Height          =   5280
      Left            =   120
      TabIndex        =   26
      Top             =   810
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   9313
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
      Left            =   13350
      TabIndex        =   32
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   120
      TabIndex        =   24
      Top             =   6150
      Width           =   2385
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
         Height          =   255
         Left            =   45
         TabIndex        =   25
         Top             =   210
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   330
      Left            =   4440
      Top             =   120
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
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   13830
      TabIndex        =   52
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
   Begin VB.Label Label5 
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
      Left            =   12150
      TabIndex        =   37
      Top             =   2460
      Width           =   495
   End
   Begin VB.Label Label1 
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
      Height          =   285
      Index           =   1
      Left            =   7800
      TabIndex        =   36
      Top             =   2430
      Width           =   945
   End
   Begin VB.Label Label6 
      Caption         =   "Web"
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
      Left            =   7800
      TabIndex        =   35
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Mail"
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
      Left            =   7800
      TabIndex        =   34
      Top             =   3285
      Width           =   495
   End
   Begin VB.Image imgWeb 
      Height          =   240
      Index           =   0
      Left            =   8430
      Picture         =   "frmManCoope.frx":0021
      Top             =   2880
      Width           =   240
   End
   Begin VB.Image imgMail 
      Height          =   240
      Index           =   0
      Left            =   8430
      Top             =   3300
      Width           =   240
   End
   Begin VB.Label Label3 
      Caption         =   "CIF"
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
      Left            =   7800
      TabIndex        =   31
      Top             =   795
      Width           =   1215
   End
   Begin VB.Label Label1 
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
      Index           =   7
      Left            =   7800
      TabIndex        =   30
      Top             =   1185
      Width           =   855
   End
   Begin VB.Label Label2 
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
      Left            =   7800
      TabIndex        =   29
      Top             =   2025
      Width           =   945
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
      Index           =   9
      Left            =   7800
      TabIndex        =   28
      Top             =   1635
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "C.Postal"
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
      Index           =   8
      Left            =   12150
      TabIndex        =   27
      Top             =   795
      Width           =   975
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
         Caption         =   "&Ver todos"
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
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
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
Attribute VB_Name = "frmManCoope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: Monica (Se lo copia)                          +-+-
' +-+- Menú: Cooperativas (con un par)                      +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-

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

Private Const IdPrograma = 2001



Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'altre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String

Public DeConsulta As Boolean
Public CodigoActual As String

' *** adrede: per a quan busque suplements/desconters des de frmViagrc ***
Public ExpedBusca As Long
Public TipoSuplem As Integer
' *********************************************************************

' *** declarar els formularis als que vaig a cridar ***
'Private WithEvents frmB As frmBuscaGrid

Private CadenaConsulta As String
Private CadB As String

' ### [Monica] 08/09/2006
Dim Indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos

Private kCampo As Integer

Dim Modo As Byte
'----------- MODOS --------------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'--------------------------------------------------

Private Sub PonerModo(vModo)
Dim B As Boolean
Dim I As Integer
    
    Modo = vModo
'    PonerIndicador lblIndicador, Modo
    
    B = (Modo = 2)
    If B Then
        PonerContRegIndicador lblIndicador, adodc1, CadB
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    ' **** posar tots els controls (botons inclosos) que siguen del Grid
    txtAux(0).visible = Not B
    txtAux(1).visible = Not B
    txtAux(2).visible = Not B
    ' **************************************************
    
    ' **** si n'hi han camps fora del grid, bloquejar-los ****
    For I = 3 To txtAux.Count - 1
        BloquearTxt txtAux(I), B
    Next I
    
    ' ********************************************************

    cmdAceptar.visible = Not B
    cmdCancelar.visible = Not B
    DataGrid1.Enabled = B
    
    'Si es retornar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = B
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botons de menu según Modo
    PonerOpcionesMenu 'Activar/Desact botons de menu según permissos de l'usuari
    
    ' *** bloquejar tota la PK quan estem en modificar  ***
    BloquearTxt txtAux(0), (Modo = 4) 'codbanpr
    
    BloquearImgBuscar Me, Modo

End Sub

Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botons de la toolbar i del menu, según el modo en que estiguem
Dim B As Boolean

    ' *** adrede: per a que no es puga fer res si estic cridant des de frmViagrc ***

    B = (Modo = 2) And ExpedBusca = 0
    'Busqueda
    Toolbar1.Buttons(5).Enabled = B
    Me.mnBuscar.Enabled = B
    'Vore Tots
    Toolbar1.Buttons(6).Enabled = B
    Me.mnVerTodos.Enabled = B
    
    'Insertar
    Toolbar1.Buttons(1).Enabled = B And Not DeConsulta
    Me.mnNuevo.Enabled = B And Not DeConsulta
    
    B = (B And adodc1.Recordset.RecordCount > 0) And Not DeConsulta And ExpedBusca = 0
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    Me.mnModificar.Enabled = B

    'Eliminar
    Toolbar1.Buttons(3).Enabled = B
    Me.mnEliminar.Enabled = B
    'Imprimir
    Toolbar1.Buttons(8).Enabled = B
    Me.mnImprimir.Enabled = B

    ' ******************************************************************************
End Sub

Private Sub BotonAnyadir()
Dim NumF As String
Dim anc As Single
Dim I As Integer
    
    CargaGrid 'primer de tot carregue tot el grid
    CadB = ""
    '********* canviar taula i camp; repasar codEmpre ************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = SugerirCodigoSiguienteStr("rcoope", "codcoope")
    End If
    '***************************************************************
    'Situem el grid al final
    AnyadirLinea DataGrid1, adodc1

    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 206
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    
    ' *** valors per defecte a l'afegir (dins i fora del grid); repasar codEmpre ***
    txtAux(0).Text = NumF
    For I = 1 To txtAux.Count - 1
        txtAux(I).Text = ""
    Next I
    ' **************************************************

    LLamaLineas anc, 3
       
    ' *** posar el foco ***
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        PonerFoco txtAux(1) '**** 1r camp visible que NO siga PK ****
    Else
        PonerFoco txtAux(0) '**** 1r camp visible que siga PK ****
    End If
    ' ******************************************************
End Sub

Private Sub BotonVerTodos()
    CargaGrid ""
    CadB = ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    Dim I As Integer
    
    ' *** canviar per la PK (no posar codempre si està a Form_Load) ***
    CargaGrid "codcoope = -1"
    '*******************************************************************************

    ' *** canviar-ho pels valors per defecte al buscar (dins i fora del grid);
    For I = 0 To txtAux.Count - 1
        txtAux(I).Text = ""
    Next I

    LLamaLineas DataGrid1.Top + 250, 1 '206, 1
    
    ' *** posar el foco al 1r camp visible que siga PK ***
    PonerFoco txtAux(0)
    ' ***************************************************************
End Sub

Private Sub BotonModificar()
    Dim anc As Single
    Dim I As Integer

    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top 'DataGrid1.RowTop(DataGrid1.Row) + 625 '545
    End If

    ' *** asignar als controls del grid, els valors de les columnes ***
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux(2).Text = ComprobarCero(Trim(DataGrid1.Columns(2).Text))
    txtAux(3).Text = ComprobarCero(Trim(DataGrid1.Columns(3).Text))
    txtAux(4).Text = DataGrid1.Columns(4).Text
    txtAux(5).Text = DataGrid1.Columns(5).Text
    ' ********************************************************

    LLamaLineas anc, 4 'modo 4
   
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco txtAux(1)
    ' *********************************************************
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
Dim I As Integer

    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo

    ' *** posar el Top a tots els controls del grid (botons també) ***
    'Me.imgFec(2).Top = alto
    For I = 0 To 2
        txtAux(I).Top = alto
    Next I
    ' ***************************************************
End Sub

Private Sub BotonEliminar()
Dim SQL As String
Dim temp As Boolean

    On Error GoTo Error2
    
    'Certes comprovacions
    If adodc1.Recordset.EOF Then Exit Sub
'    If Not SepuedeBorrar Then Exit Sub
    
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(adodc1.Recordset.Fields(1).Value), FormatoCampo(txtAux(1))) Then Exit Sub
    ' ***************************************************************************
    
    '*** canviar la pregunta, els noms dels camps i el DELETE; repasar codEmpre ***
    SQL = "¿Seguro que desea eliminar la Cooperativa?"
    'SQL = SQL & vbCrLf & "Código: " & Format(adodc1.Recordset.Fields(0), "000")
    SQL = SQL & vbCrLf & "Código: " & adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Nombre: " & adodc1.Recordset.Fields(1)
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'N'hi ha que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        SQL = "Delete from rcoope where codcoope = " & adodc1.Recordset!codcoope
        
        conn.Execute SQL
        CargaGrid CadB
'        If CadB <> "" Then
'            CargaGrid CadB
'            lblIndicador.Caption = "RESULTADO BUSQUEDA"
'        Else
'            CargaGrid ""
'            lblIndicador.Caption = ""
'        End If
        temp = SituarDataTrasEliminar(adodc1, NumRegElim, True)
        PonerModoOpcionesMenu
        adodc1.Recordset.Cancel
    End If
    Exit Sub
    
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los txtAux
    PonerLongCamposGnral Me, Modo, 3
End Sub

Private Sub cmdAceptar_Click()
Dim I As Long

    Select Case Modo
        Case 3 'INSERTAR
            If DatosOK Then
                'If InsertarDesdeForm(Me) Then
                If InsertarDesdeForm2(Me, 0) Then
                    CargaGrid
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                        cmdCancelar_Click
                        If Not adodc1.Recordset.EOF Then
                            ' *** filtrar per tota la PK; repasar codEmpre **
                            'adodc1.Recordset.Filter = "codempre = " & txtAux(0).Text & " AND codsupdt = " & txtAux(1).Text
                            adodc1.Recordset.Filter = "codcoope = " & txtAux(0).Text
                            ' ****************************************************
                        End If
                        cmdRegresar_Click
                    Else
                        BotonAnyadir
                    End If
                    CadB = ""
                End If
            End If
            
        Case 4 'MODIFICAR
            If DatosOK Then
                'If ModificaDesdeFormulario(Me) Then
                If ModificaDesdeFormulario2(Me, 0) Then
                    I = adodc1.Recordset.AbsolutePosition
                    TerminaBloquear
                    PonerModo 2
                    CargaGrid CadB
'                    If CadB <> "" Then
'                        CargaGrid CadB
'                        lblIndicador.Caption = "RESULTADO BUSQUEDA"
'                    Else
'                        CargaGrid
'                        lblIndicador.Caption = ""
'                    End If
                    adodc1.Recordset.Move I - 1
                    PonerFocoGrid Me.DataGrid1
                End If
            End If
            
        Case 1  'BUSQUEDA
            CadB = ObtenerBusqueda(Me)
            If CadB <> "" Then
                CargaGrid CadB
                PonerModo 2
'                lblIndicador.Caption = "RESULTADO BUSQUEDA"
                PonerFocoGrid Me.DataGrid1
            End If
    End Select
End Sub

Private Sub cmdCancelar_Click()
'On Error Resume Next

    Select Case Modo
        Case 3 'INSERTAR
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        Case 4 'MODIFICAR
            TerminaBloquear
        Case 1 'BUSQUEDA
            CargaGrid CadB
    End Select
    
    If Not adodc1.Recordset.EOF Then
        CargaForaGrid
    Else
        LimpiarCampos
    End If
    
    PonerModo 2
'    If CadB <> "" Then
'        lblIndicador.Caption = "RESULTADO BUSQUEDA"
'    Else
'        lblIndicador.Caption = ""
'    End If
    PonerFocoGrid Me.DataGrid1
'    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim I As Integer
Dim J As Integer
Dim Aux As String

    If adodc1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    cad = ""
    I = 0
    Do
        J = I + 1
        I = InStr(J, DatosADevolverBusqueda, "|")
        If I > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, I - J)
            J = Val(Aux)
            cad = cad & adodc1.Recordset.Fields(J) & "|"
        End If
    Loop Until I = 0
    ' *** adrede: per a tornar el TipoSuplem ***
    ' cad = cad & TipoSuplem & "|"
    ' ******************************************
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    'Posem el foco
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        PonerFoco txtAux(1)
    End If
End Sub

Private Sub Form_Load()
Dim I As Integer
    
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    '******* repasar si n'hi ha botó d'imprimir o no******
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
        'el 4 es separadors
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2   'Tots
        'el 7 es separadors
        .Buttons(8).Image = 10  'Imprimir
    End With
    '*****************************************************
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 12
    End With
    


    'carga IMAGES de mail
    For I = 0 To Me.imgMail.Count - 1
        Me.imgMail(I).Picture = frmPpal.imgListImages16.ListImages(2).Picture
    Next I

'    chkVistaPrevia.Value = CheckValueLeer(Name)
    ' *** SI N'HI HAN COMBOS ***
    ' CargaCombo 0
    ' **************************
    
    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT codcoope, nomcoope, porcgast, domcoope, codposta, "
    CadenaConsulta = CadenaConsulta & "pobcoope, procoope, cifcoope, telcoope, "
    CadenaConsulta = CadenaConsulta & "faxcoope, wwwcoope, maicoope, "
    CadenaConsulta = CadenaConsulta & "codtipomant, codtipomliq, codtipomantvc, codtipomliqvc, codtipomantalmz, codtipomliqalmz, "
    CadenaConsulta = CadenaConsulta & "codtipomantbod, codtipomliqbod, codtipomfacalmz, codtipomfacbod "
    CadenaConsulta = CadenaConsulta & " FROM rcoope "
    '************************************************************************
    
    CadB = ""
    CargaGrid
    
    ' ****** Si n'hi han camps fora del grid ******
    'CargaForaGrid
    ' *********************************************
    
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        BotonAnyadir
    Else
        PonerModo 2
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
    If Modo = 4 Then TerminaBloquear
    
End Sub

Private Sub imgMail_Click(Index As Integer)
    If Index = 0 Then
        If txtAux(11).Text <> "" Then
            LanzaMailGnral txtAux(11).Text
        End If
    End If
End Sub

Private Sub imgWeb_Click(Index As Integer)
    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    If LanzaHomeGnral(txtAux(10).Text) Then espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    printNou
End Sub

Private Sub mnModificar_Click()
    'Comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(adodc1.Recordset.Fields(1).Value), FormatoCampo(txtAux(1))) Then Exit Sub
    ' ***************************************************************************
    
    
    'Prepara para modificar
    '-----------------------
    If BLOQUEADesdeFormulario2(Me, adodc1, 1) Then BotonModificar
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
        Case 1
                mnNuevo_Click
        Case 2
                mnModificar_Click
        Case 3
                mnEliminar_Click
        Case 5
                mnBuscar_Click
        Case 6
                mnVerTodos_Click
        Case 8 'Imprimir
                mnImprimir_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String)
    Dim I As Integer
    Dim SQL As String
    Dim tots As String
    
'    adodc1.ConnectionString = Conn
    ' *** si en Form_load ya li he posat clausula WHERE, canviar el `WHERE` de
    ' `SQL = CadenaConsulta & " WHERE " & vSQL` per un `AND`
    If vSQL <> "" Then
        SQL = CadenaConsulta & " WHERE " & vSQL  ' ### [Monica] 08/09/2006: antes habia AND
    Else
        SQL = CadenaConsulta
    End If
    '********************* canviar el ORDER BY *********************++
    'SQL = SQL & " ORDER BY codempre, codsupdt"
    SQL = SQL & " ORDER BY codcoope"
    '**************************************************************++
    
'    adodc1.RecordSource = SQL
'    adodc1.CursorType = adOpenDynamic
'    adodc1.LockType = adLockOptimistic
'    DataGrid1.ScrollBars = dbgNone
'    adodc1.Refresh
'    Set DataGrid1.DataSource = adodc1 ' per a que no ixca l'error de "la fila actual no está disponible"
       
    CargaGridGnral Me.DataGrid1, Me.adodc1, SQL, False
       
       
    ' *** posar només els controls del grid ***
    tots = "S|txtAux(0)|T|Código|850|;S|txtAux(1)|T|Denominación|4650|;S|txtAux(2)|T|% Gastos|1450|;"
    For I = 3 To txtAux.Count - 1
        tots = tots & "N||||0|;"
    Next I
    arregla tots, DataGrid1, Me, 350
    DataGrid1.ScrollBars = dbgAutomatic
    ' **********************************************************
    
    ' *** alliniar les columnes que siguen numèriques a la dreta ***
    'DataGrid1.Columns(1).Alignment = dbgRight
    'DataGrid1.Columns(2).Alignment = dbgRight
    'DataGrid1.Columns(4).Alignment = dbgRight
    ' *****************************
    
    
    ' *** Si n'hi han camps fora del grid ***
    If Not adodc1.Recordset.EOF Then
        CargaForaGrid
    Else
        LimpiarCampos
    End If
    ' **************************************
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub



Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
'    If Index = 3 And KeyAscii = 43 Then '+
'        KeyAscii = 0
'    Else
'        KEYpress KeyAscii
'    End If
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 12: KEYBusqueda KeyAscii, 0 'cuenta contable
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
    
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
'    imgBuscar_Click (indice)
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    '*** configurar el LostFocus dels camps (de dins i de fora del grid) ***
    Select Case Index
        Case 0
            PonerFormatoEntero txtAux(Index)
        
        Case 1
            If txtAux(Index).Text = "" Then Exit Sub
            txtAux(Index).Text = UCase(txtAux(Index).Text)
        
        Case 2
            PonerFormatoDecimal txtAux(Index), 4 'tipo 4: Decimal(5,2)
            
            
'        Case 12 'cuenta contable
'            If txtAux(Index).Text = "" Then Exit Sub
'            Text2(Index).Text = PonerNombreCuenta(txtAux(Index), Modo)

'       Case 6, 7 'dates
'           If txtAux(Index).Text <> "" Then PonerFormatoFecha txtAux(Index)
            
'       Case 8 'Importe
'           If txtAux(Index).Text <> "" Then PonerFormatoDecimal txtAux(Index), 1 'tipo1: Decimal(12,2)
            
'       Case 9
'           PonerFormatoDecimal txtAux(Index), 4 'tipo 4: Decimal(5,2)
    End Select
    '**************************************************************************
End Sub


Private Function DatosOK() As Boolean
Dim Datos As String
Dim B As Boolean
' *** només per ad este manteniment ***
Dim Rs As Recordset
Dim cad As String
Dim I As Integer
Dim SQL As String
Dim Nregs As Long
Dim Tipo As Byte
'Dim exped As String
' *************************************

    B = CompForm(Me)
    If Not B Then Exit Function


    If B And (Modo = 3) Then
        'Estem insertant
        'aço es com posar: select codvarie from svarie where codvarie = txtAux(0)
        'la N es pa dir que es numèric
         
        ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
        Datos = DevuelveDesdeBD("codcoope", "rcoope", "codcoope", txtAux(0).Text, "N")
         
        If Datos <> "" Then
            MsgBox "Ya existe el Código de Cooperativa: " & txtAux(0).Text, vbExclamation
            DatosOK = False
            PonerFoco txtAux(1) '*** posar el foco al 1r camp visible de la PK de la capçalera ***
            Exit Function
        End If
        '*************************************************************************************
    End If

    ' *** Si cal fer altres comprovacions ***
    ' los tipos de movimientos deben de existir en la stipom
    If B And (Modo = 3 Or Modo = 4) Then
        For I = 12 To 19
            Select Case I
                Case 12 ', 16, 18
                    Tipo = 1
                Case 13 ', 17, 19
                    Tipo = 2
                Case 14
                    Tipo = 3
                Case 15
                    Tipo = 4
                Case 16
                    Tipo = 7
                Case 17
                    Tipo = 8
                Case 18
                    Tipo = 9
                Case 19
                    Tipo = 10
            End Select
        
            SQL = "select count(*) from usuarios.stipom where codtipom = " & DBSet(txtAux(I).Text, "T") & " and tipodocu = " & Tipo
            Nregs = TotalRegistros(SQL)
            If Nregs = 0 Then
                cad = "No existe el Tipo de Movimiento de "
                Select Case I
                    Case 12
                        cad = cad & " Anticipos,"
                    Case 14
                        cad = cad & " Anticipos de Venta Campo,"
                    Case 13
                        cad = cad & " Liquidación,"
                    Case 15
                        cad = cad & " Liquidación de Venta Campo,"
                    Case 16
                        cad = cad & " Anticipos de Almazara,"
                    Case 17
                        cad = cad & " Liquidación de Almazara,"
                    Case 18
                        cad = cad & " Anticipos de Bodega,"
                    Case 19
                        cad = cad & " Liquidación de Bodega,"
                End Select
                cad = cad & " o no es del tipo correcto."
                cad = cad & vbCrLf & vbCrLf
                cad = cad & "LLame a Ariadna.            "
                
                MsgBox cad, vbExclamation
                DatosOK = False
                PonerFoco txtAux(I)
                Exit Function
            End If
        Next I
    
    End If
    ' *********************************************

    DatosOK = B
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me

End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim I As Byte

    If Modo <> 4 Then 'Modificar
        CargaForaGrid
    Else
        For I = 0 To txtAux.Count - 1
            txtAux(I).Text = ""
        Next I
    End If
    
    If Modo = 2 Then PonerContRegIndicador lblIndicador, adodc1, CadB
    
End Sub

Private Sub CargaForaGrid()
Dim I As Integer

    If DataGrid1.Columns.Count <= 2 Then Exit Sub
    
    ' *** posar als camps de fora del grid el valor de la columna corresponent ***
    
    For I = 3 To 21
        txtAux(I) = DataGrid1.Columns(I).Text
    Next I
 End Sub

Private Sub LimpiarCampos()
Dim I As Integer
On Error Resume Next

    ' *** posar a huit tots els camps de fora del grid ***
    For I = 3 To txtAux.Count - 1
        txtAux(I).Text = ""
    Next I
    ' ****************************************************

    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "rcoope"
        .Informe2 = "rManCoope.rpt"
        If CadB <> "" Then
            '.cadRegSelec = Replace(SQL2SF(CadB), "clientes", "clientes_1")
            .cadRegSelec = SQL2SF(CadB)
        Else
            .cadRegSelec = ""
        End If
        ' *** repasar el nom de l'adodc ***
        '.cadRegActua = Replace(POS2SF(Data1, Me), "clientes", "clientes_1")
        .cadRegActua = POS2SF(adodc1, Me)
        '[Monica]13/07/2012: falla si hay un solo registro seleccionado y apretamos registros buscados
        If adodc1.Recordset.RecordCount = 1 Then .cadRegSelec = .cadRegActua
        
        ' *** repasar codEmpre ***
        .cadTodosReg = ""
        '.cadTodosReg = "{sbanco.codbanpr} = " & vSesion.Empresa
        ' *** repasar si li pose ordre o no ****
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|pOrden={rcoope.codcoope}|"
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
'Private Sub DataGrid1_GotFocus()
'  WheelHook DataGrid1
'End Sub
'Private Sub DataGrid1_Lostfocus()
'  WheelUnHook
'End Sub

