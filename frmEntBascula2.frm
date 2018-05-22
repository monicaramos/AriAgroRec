VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEntBascula2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entrada de Paletización"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   8280
   Icon            =   "frmEntBascula2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   45
      TabIndex        =   47
      Top             =   90
      Width           =   1470
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   48
         Top             =   180
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar Líneas"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Palets/Palots"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FramePalets 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   4950
      TabIndex        =   44
      Top             =   2925
      Width           =   3030
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
         Index           =   10
         Left            =   180
         MaxLength       =   3
         TabIndex        =   0
         Top             =   225
         Width           =   1200
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
         Index           =   11
         Left            =   1665
         MaxLength       =   7
         TabIndex        =   1
         Top             =   225
         Width           =   1200
      End
      Begin VB.Label Label2 
         Caption         =   "Palets"
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
         TabIndex        =   46
         Top             =   0
         Width           =   1050
      End
      Begin VB.Label Label8 
         Caption         =   "Palots"
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
         Left            =   1665
         TabIndex        =   45
         Top             =   0
         Width           =   1185
      End
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
      Index           =   10
      Left            =   6075
      MaxLength       =   20
      TabIndex        =   12
      Tag             =   "CRFID|T|N|||trzpalets|crfid|||"
      Text            =   "crfid"
      Top             =   6570
      Width           =   495
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
      Index           =   9
      Left            =   5670
      MaxLength       =   7
      TabIndex        =   11
      Tag             =   "Nro Nota|N|N|||trzpalets|numnotac|0000000||"
      Text            =   "numnot"
      Top             =   6570
      Width           =   405
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
      Index           =   8
      Left            =   5220
      MaxLength       =   6
      TabIndex        =   10
      Tag             =   "Hora|FH|N|||trzpalets|hora|dd/mm/yyyy hh:mm:ss||"
      Text            =   "hora"
      Top             =   6570
      Width           =   405
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
      Index           =   7
      Left            =   4455
      MaxLength       =   6
      TabIndex        =   9
      Tag             =   "Fecha|F|N|||trzpalets|fecha|dd/mm/yyyy||"
      Text            =   "fecha"
      Top             =   6570
      Width           =   675
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
      Index           =   6
      Left            =   4050
      MaxLength       =   6
      TabIndex        =   8
      Tag             =   "Variedad|N|N|||trzpalets|codvarie|||"
      Text            =   "varie"
      Top             =   6570
      Width           =   360
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
      Index           =   5
      Left            =   3510
      MaxLength       =   6
      TabIndex        =   7
      Tag             =   "Cod.Campo|N|N|||trzpalets|codcampo|||"
      Text            =   "campo"
      Top             =   6570
      Width           =   495
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
      Left            =   2295
      MaxLength       =   7
      TabIndex        =   5
      Tag             =   "Num.Kilos|N|N|||trzpalets|numkilos|###,##0||"
      Text            =   "kil"
      Top             =   6570
      Width           =   675
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   3225
      Index           =   0
      Left            =   60
      TabIndex        =   19
      Top             =   885
      Width           =   8055
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
         Index           =   8
         Left            =   180
         MaxLength       =   3
         TabIndex        =   40
         Top             =   2700
         Width           =   1305
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
         MaxLength       =   30
         TabIndex        =   34
         Top             =   2085
         Width           =   5595
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
         Index           =   6
         Left            =   1170
         MaxLength       =   6
         TabIndex        =   33
         Top             =   2070
         Width           =   900
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
         Index           =   3
         Left            =   1170
         MaxLength       =   6
         TabIndex        =   32
         Top             =   855
         Width           =   885
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
         Left            =   2160
         MaxLength       =   30
         TabIndex        =   31
         Top             =   855
         Width           =   5610
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
         Index           =   4
         Left            =   1170
         MaxLength       =   6
         TabIndex        =   30
         Top             =   1260
         Width           =   900
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
         Left            =   2160
         MaxLength       =   40
         TabIndex        =   29
         Top             =   1275
         Width           =   5610
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
         Index           =   7
         Left            =   6525
         MaxLength       =   8
         TabIndex        =   28
         Top             =   405
         Width           =   1260
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
         Left            =   2160
         MaxLength       =   30
         TabIndex        =   27
         Top             =   1665
         Width           =   5595
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
         Index           =   5
         Left            =   1170
         MaxLength       =   4
         TabIndex        =   26
         Top             =   1665
         Width           =   900
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
         Index           =   0
         Left            =   180
         MaxLength       =   7
         TabIndex        =   22
         Top             =   405
         Width           =   1365
      End
      Begin VB.TextBox Text1 
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
         Index           =   1
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   21
         Top             =   405
         Width           =   1260
      End
      Begin VB.TextBox Text1 
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
         Left            =   3780
         MaxLength       =   10
         TabIndex        =   20
         Top             =   405
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
         Left            =   2160
         MaxLength       =   7
         TabIndex        =   42
         Top             =   2700
         Width           =   1650
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
         Left            =   2160
         TabIndex        =   43
         Top             =   2475
         Width           =   1320
      End
      Begin VB.Label Label13 
         Caption         =   "Nro. Cajas"
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
         TabIndex        =   41
         Top             =   2475
         Width           =   1185
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
         Left            =   180
         TabIndex        =   39
         Top             =   2115
         Width           =   645
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
         Left            =   180
         TabIndex        =   38
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label Label5 
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
         Height          =   255
         Left            =   180
         TabIndex        =   37
         Top             =   900
         Width           =   960
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
         Left            =   180
         TabIndex        =   36
         Top             =   1305
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Nº Campo"
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
         Left            =   6525
         TabIndex        =   35
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Nota"
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
         Left            =   180
         TabIndex        =   25
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label18 
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
         Left            =   2160
         TabIndex        =   24
         Top             =   180
         Width           =   885
      End
      Begin VB.Label Label26 
         Caption         =   "Hora"
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
         Left            =   3780
         TabIndex        =   23
         Top             =   180
         Width           =   570
      End
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
      Left            =   1440
      MaxLength       =   7
      TabIndex        =   4
      Tag             =   "Num.Cajones|N|N|||trzpalets|numcajones|###,##0||"
      Text            =   "caj"
      Top             =   6570
      Width           =   900
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
      Index           =   4
      Left            =   3015
      MaxLength       =   6
      TabIndex        =   6
      Tag             =   "Cod.Socio|N|N|||trzpalets|codsocio|||"
      Text            =   "socio"
      Top             =   6570
      Width           =   450
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
      Left            =   5850
      TabIndex        =   13
      Tag             =   "   "
      Top             =   7965
      Visible         =   0   'False
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
      Left            =   6990
      TabIndex        =   14
      Top             =   7965
      Visible         =   0   'False
      Width           =   1065
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
      Left            =   810
      MaxLength       =   30
      TabIndex        =   3
      Tag             =   "Tipo|N|N|||trzpalets|tipo|||"
      Text            =   "tipo"
      Top             =   6570
      Width           =   540
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
      Left            =   315
      MaxLength       =   11
      TabIndex        =   2
      Tag             =   "IdPalet|N|N|||trzpalets|idpalet|00000000000|S|"
      Text            =   "idpale"
      Top             =   6570
      Width           =   480
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmEntBascula2.frx":000C
      Height          =   3510
      Left            =   90
      TabIndex        =   17
      Top             =   4185
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   6191
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
      Left            =   6975
      TabIndex        =   18
      Top             =   7965
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   90
      TabIndex        =   15
      Top             =   7785
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
         TabIndex        =   16
         Top             =   195
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   3150
      Top             =   945
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
   Begin VB.Menu mnOpciones 
      Caption         =   "Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnModificar 
         Caption         =   "Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnPaletsPalots 
         Caption         =   "Palets/Palots"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmEntBascula2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MANOLO  +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'atre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public crear As Byte ' 0=false 1=true
Public NumNota As String
Public NumCajones As String
Public NumKilos As String
Public Codsocio As String
Public codcampo As String
Public Codvarie As String
Public Fecha As String
Public Hora As String

Private CadenaConsulta As String
Private CadB As String


Dim Modo As Byte
'----------- MODOS --------------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'--------------------------------------------------
Dim PrimeraVez As Boolean
Dim Indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim I As Integer

Private Sub PonerModo(vModo)
Dim B As Boolean

    Modo = vModo
    
    B = (Modo = 2)
    If B Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    For I = 0 To txtAux.Count - 1
        txtAux(I).Enabled = False
        txtAux(I).visible = False
    Next I
    
'    txtAux(0).Enabled = (Modo = 4)
    txtAux(0).visible = (Modo = 4)
    txtAux(2).Enabled = (Modo = 4)
    txtAux(2).visible = (Modo = 4)
    txtAux(3).visible = (Modo = 4)
    txtAux(3).Enabled = (Modo = 4)
    txtAux(10).visible = (Modo = 4)
    txtAux(10).Enabled = (Modo = 4)
    
    CmdAceptar.visible = Not B
    CmdCancelar.visible = Not B
'    DataGrid1.Enabled = b
    
    FramePalets.Enabled = (Modo = 5)
    Text1(10).Enabled = (Modo = 5) And Val(NumCajones) <> 0
    Text1(11).Enabled = (Modo = 5) And Val(NumCajones) = 0
    'Si es regresar
    'If DatosADevolverBusqueda <> "" Then
    cmdRegresar.visible = B
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu  'En funcion del usuario
    
    'Si estamos modo Modificar bloquear clave primaria
'    BloquearTxt txtAux(0), (Modo = 4)
End Sub


Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botones del la toobar y del menu, segun el modo en que estemos
Dim B As Boolean

    B = (Modo = 2)

    B = (B And adodc1.Recordset.RecordCount > 0)
    'Modificar
    Toolbar1.Buttons(1).Enabled = B
    Me.mnModificar.Enabled = B
    'Modificar Palets/Palots
    Toolbar1.Buttons(2).Enabled = B
    Me.mnPaletsPalots.Enabled = B

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
        anc = anc + 240 ' 320
    Else
        anc = anc + DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 5 'DataGrid1.RowTop(DataGrid1.Row) + 3330 'DataGrid1.RowTop(DataGrid1.Row) + 545
    End If

    'Llamamos al form
    For I = 0 To 10
        txtAux(I).Text = DataGrid1.Columns(I).Text
    Next I
    
    ' ***** canviar-ho pel nom del camp del combo *********
'    SelComboBool DataGrid1.Columns(2).Text, Combo1(0)
    ' *****************************************************

    'PosicionarCombo Me.Combo1(0), i
    'PosicionarCombo Me.Combo1(1), i

    LLamaLineas anc, 4 'Pone el form en Modo=4, Modificar
   
    'Como es modificar
    PonerFoco txtAux(10)
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    'Fijamos el ancho
    For I = 0 To txtAux.Count - 1
        txtAux(I).Top = alto
    Next I
    ' ### [Monica] 12/09/2006
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los txtAux
    PonerLongCamposGnral Me, Modo, 3
End Sub

Private Sub cmdAceptar_Click()
    Dim I As Double

    Select Case Modo
        Case 4 'MODIFICAR lineas
            If DatosOkLin Then
                If ModificaDesdeFormulario(Me) Then
                    TerminaBloquear
                    I = adodc1.Recordset.Fields(0)
                    PonerModo 2
                    CargaGrid CadB
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & I)
                    PonerFocoGrid Me.DataGrid1
                    
                    PasarSigReg
                End If
            Else
                PonerFoco txtAux(10)
            End If
            
        Case 5 ' modificar el numero de palets o palots
            If Val(NumCajones) <> 0 Then
                InsertarPalets DBSet(Text1(10).Text, "N")
            Else
                InsertarPalets DBSet(Text1(11).Text, "N")
            End If
            TerminaBloquear
            PonerModo 2
            CargaGrid CadB
            BotonModificar
    End Select
End Sub

Private Sub cmdCancelar_Click()
    On Error Resume Next
    
    Select Case Modo
        Case 4 'modificar
            TerminaBloquear
    End Select
    
    PonerModo 2
    
'    If CadB <> "" Then
'        lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
''    Else
''        lblIndicador.Caption = ""
'    End If
    
    PonerFocoGrid Me.DataGrid1
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim I As Integer
Dim J As Integer
Dim Aux As String

    If DatosOK Then
        Unload Me
    Else
        mnPaletsPalots_Click
    End If

End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    PonerContRegIndicador
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault

    If PrimeraVez Then
        
        PrimeraVez = False
        CadB = ""
        CargaGrid
        mnPaletsPalots_Click

    End If
End Sub

Private Sub Form_Load()
Dim Sql As String

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True

'    With Me.Toolbar1
'        .HotImageList = frmPpal.imgListComun_OM
'        .DisabledImageList = frmPpal.imgListComun_BN
'        .ImageList = frmPpal.imgListComun
'        'el 1 es separadors
'        .Buttons(2).Image = 1   'Buscar
'        .Buttons(3).Image = 2   'Todos
'        'el 4 i el 5 son separadors
'        .Buttons(6).Image = 3   'Insertar
'        .Buttons(7).Image = 4   'Modificar
'        .Buttons(8).Image = 19   'Modificar Palets / Palots
'        'el 9 i el 10 son separadors
'        .Buttons(11).Image = 10  'imprimir
'        .Buttons(12).Image = 11  'Salir
'    End With
    
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 4    'Modificar
        .Buttons(2).Image = 19   'Modificar Palets / Palots
    End With
    

    '## A mano
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT idpalet,tipo,numcajones,numkilos,codsocio,codcampo,codvarie,fecha,hora,numnotac,CRFID "
    CadenaConsulta = CadenaConsulta & " FROM trzpalets"
    CadenaConsulta = CadenaConsulta & " WHERE numnotac = " & Trim(NumNota) & ""
    '************************************************************************
    
    CargarDatosCabecera
    
    If crear = 1 Then
        If Val(NumCajones) <> 0 Then
            InsertarPalets 0
        Else
            InsertarPalets 0
        End If
    Else
        Sql = "select count(*) from trzpalets where numnotac = " & Trim(NumNota)
        If Val(NumCajones) <> 0 Then
            Text1(10).Text = Format(TotalRegistros(Sql), "###,##0")
            Text1(11).Text = ""
        Else
            Text1(10).Text = ""
            Text1(11).Text = Format(TotalRegistros(Sql), "###,##0")
        End If
    End If
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
   If SalirFormulario Then
        Cancel = 0
   Else
        Cancel = 1
   End If
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub mnModificar_Click()
    'Comprobaciones
    '--------------
    If adodc1.Recordset.EOF Then Exit Sub
    
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCampo(txtAux(0))) Then Exit Sub
    
    
    'Preparamos para modificar
    '-------------------------
    TerminaBloquear
    
    'If BLOQUEADesdeFormulario2(Me, adodc1, 1) Then BotonModificar
    If BloqueaRegistro("trzpalets", "numnotac = " & Trim(NumNota)) Then BotonModificar
    
End Sub

Private Sub mnPaletsPalots_Click()
    'boton modificar palets o palots
    TerminaBloquear
    
    If BloqueaRegistro("trzpalets", "numnotac = " & Trim(NumNota)) Then
        PonerModo 5
        If Val(NumCajones) <> 0 Then
            PonerFoco Text1(10)
        Else
            txtAux(2).Text = "0"
            txtAux(3).Text = "0"
            PonerFoco Text1(11)
        End If
    End If
End Sub

Private Sub mnSalir_Click()
    If DatosOK Then Unload Me
End Sub


Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean


If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    
    ' ***************** configurar els LostFocus dels camps de la capçalera *****************
    Select Case Index
        Case 10, 11 'palets y palots
            If PonerFormatoEntero(Text1(Index)) Then
                If CCur(Text1(Index).Text) <> 0 Then
                    CmdAceptar.SetFocus
                Else
                    MsgBox "Debe introducir un valor distinto de cero", vbExclamation
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
            End If
    End Select
        ' ***************************************************************************
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
                mnModificar_Click
        Case 2
                mnPaletsPalots_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String)
    Dim Sql As String
    Dim tots As String
    
'    adodc1.ConnectionString = Conn
    If vSQL <> "" Then
        Sql = CadenaConsulta & " AND " & vSQL
    Else
        Sql = CadenaConsulta
    End If
    '********************* canviar el ORDER BY *********************++
    Sql = Sql & " ORDER BY trzpalets.idpalet"
    '**************************************************************++
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, Sql, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "S|txtAux(0)|T|Id.Palet|1500|;N|||||;"
    tots = tots & "S|txtAux(2)|T|Cajas|1500|;S|txtAux(3)|T|Kilos|1800|;N|||||;"
    tots = tots & "N|||||;N|||||;N|||||;N|||||;N|||||;"
    tots = tots & "S|txtAux(10)|T|CRFID|2600|;"
    
    arregla tots, DataGrid1, Me, 350
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgRight
    DataGrid1.Columns(2).Alignment = dbgRight
    DataGrid1.Columns(3).Alignment = dbgRight
    DataGrid1.Columns(10).Alignment = dbgLeft
    
End Sub


Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub


Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo EKeyD
    If Index <> 10 Then Exit Sub
    
    If (KeyCode = 38 Or KeyCode = 40) Then
        cmdAceptar_Click 'ModificarExistencia
    End If

    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
                If DataGrid1.Bookmark > 0 Then
                    DataGrid1.Bookmark = DataGrid1.Bookmark - 1
                End If

        Case 40 'Desplazamiento Flecha Hacia Abajo
                PasarSigReg

    End Select
EKeyD:
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)

   
   If Index <> 10 Then
        KEYpress KeyAscii
        Exit Sub
   End If

   If KeyAscii = 13 Then 'ENTER
        cmdAceptar_Click 'ModificarExistencia
   ElseIf KeyAscii = 27 Then
        cmdCancelar_Click 'ESC
   End If

End Sub

Private Sub txtAux_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 10
            If txtAux(Index).Text = "" Then Exit Sub
            
            txtAux(Index).Text = UCase(txtAux(Index).Text)
        Case 2, 3
            PonerFormatoEntero txtAux(Index)
    End Select
    
End Sub


Private Function DatosOK() As Boolean
'Dim Datos As String
Dim B As Boolean
Dim Sql As String
Dim Mens As String
Dim Rs As ADODB.Recordset

    Sql = "select numnotac, sum(numcajones), sum(numkilos) from trzpalets where numnotac = " & DBSet(NumNota, "N")
    Sql = Sql & " group by 1 order by 1"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    B = False
    
    If Rs.EOF Then
        MsgBox "Debe haber un reparto de cajas y kilos", vbExclamation
        
    Else
        If DBLet(Rs.Fields(1).Value, "N") <> Val(NumCajones) Then
            MsgBox "La suma de número de cajas no cuadra con el total. Revise.", vbExclamation
        Else
            If DBLet(Rs.Fields(2).Value, "N") <> Val(NumKilos) Then
                MsgBox "La suma de número de kilos no cuadra con el total. Revise.", vbExclamation
            Else
                B = True
            End If
        End If
    End If
    Rs.Close
        
        
    DatosOK = B
End Function



Private Function DatosOkLin() As Boolean
'Dim Datos As String
Dim B As Boolean
Dim Sql As String
Dim Mens As String
Dim Rs As ADODB.Recordset

    B = CompForm(Me)
    If Not B Then Exit Function

    Sql = "select count(*) from trzpalets where idpalet <> " & DBSet(txtAux(0).Text, "N")
    Sql = Sql & " and crfid = " & DBSet(txtAux(10).Text, "T")
    
    If TotalRegistros(Sql) <> 0 Then
        MsgBox "Este Numero CRFID está asignado a otro palet. Revise.", vbExclamation
        txtAux(10).Text = ""
        B = False
    End If
    
    DatosOkLin = B
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub


Private Sub PonerContRegIndicador()
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    If (Modo = 2 Or Modo = 0) Then
        cadReg = PonerContRegistros(Me.adodc1)
        If CadB = "" Then
            lblIndicador.Caption = cadReg
        Else
            lblIndicador.Caption = "BUSQUEDA: " & cadReg
        End If
    End If
End Sub



Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me

End Sub


Private Sub InsertarPalets(Palets As Long)
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Kilos As Currency
Dim Cajas As Currency
Dim nroPalets As Integer
Dim RestoCajas As Currency
Dim KilosporPalet As Currency
Dim TotKilos As Currency
Dim RestoKilos As Currency
Dim NumF As String
Dim Tipo As Byte

    Sql = "delete from trzpalets where numnotac = " & Trim(NumNota)
    conn.Execute Sql
    
    If Val(Text1(8).Text) <> 0 Then ' si palets
        ' se reparte en palets las cajas
        Tipo = 0
        
        If Palets = 0 Then ' partimos del nro de cajas
            nroPalets = Val(NumCajones) \ vParamAplic.CajasporPalet
            RestoCajas = Val(NumCajones) Mod vParamAplic.CajasporPalet
            
            KilosporPalet = (vParamAplic.CajasporPalet * NumKilos) \ Val(NumCajones)
            TotKilos = 0
        
            For I = 1 To nroPalets
                NumF = SugerirCodigoSiguienteStr("trzpalets", "idpalet")
                
                TotKilos = TotKilos + KilosporPalet
                
                Sql = "insert into trzpalets (idpalet,tipo,numcajones,numkilos,"
                Sql = Sql & "codsocio,codcampo,codvarie,fecha,hora,numnotac,CRFID) values ("
                Sql = Sql & DBSet(NumF, "N") & "," & DBSet(Tipo, "N") & "," & DBSet(vParamAplic.CajasporPalet, "N") & ","
                Sql = Sql & DBSet(KilosporPalet, "N") & "," & DBSet(Codsocio, "N") & "," & DBSet(codcampo, "N") & ","
                Sql = Sql & DBSet(Codvarie, "N") & "," & DBSet(Fecha, "F") & "," & DBSet(Fecha & " " & Hora, "FH", "S") & ","
                Sql = Sql & DBSet(NumNota, "N") & "," & ValorNulo & ")"
                
                conn.Execute Sql
            Next I
            
            If RestoCajas <> 0 Then ' insertamos el ultimo palet con el resto
                NumF = SugerirCodigoSiguienteStr("trzpalets", "idpalet")
                
                RestoKilos = NumKilos - (KilosporPalet * nroPalets)
                
                TotKilos = TotKilos + RestoKilos
                
                Sql = "insert into trzpalets (idpalet,tipo,numcajones,numkilos,"
                Sql = Sql & "codsocio,codcampo,codvarie,fecha,hora,numnotac,CRFID) values ("
                Sql = Sql & DBSet(NumF, "N") & "," & DBSet(Tipo, "N") & "," & DBSet(RestoCajas, "N") & ","
                Sql = Sql & DBSet(RestoKilos, "N") & "," & DBSet(Codsocio, "N") & "," & DBSet(codcampo, "N") & ","
                Sql = Sql & DBSet(Codvarie, "N") & "," & DBSet(Fecha, "F") & "," & DBSet(Fecha & " " & Hora, "FH", "S") & ","
                Sql = Sql & DBSet(NumNota, "T") & "," & ValorNulo & ")"
                
                conn.Execute Sql
                
                nroPalets = nroPalets + 1
            End If
            
            RestoKilos = NumKilos - TotKilos
            
            If RestoKilos <> 0 Then ' actualizamos el ultimo registro si hay resto de kilos
                Sql = "update trzpalets set numkilos = numkilos + " & DBSet(RestoKilos, "N")
                Sql = Sql & " where idpalet = " & DBSet(NumF, "N")
                
                conn.Execute Sql
            End If
        Else ' partimos del nro de palets
            nroPalets = Palets
            Kilos = NumKilos \ nroPalets
            Cajas = Val(NumCajones) \ nroPalets
            
            For I = 1 To nroPalets
                NumF = SugerirCodigoSiguienteStr("trzpalets", "idpalet")
                
                Sql = "insert into trzpalets (idpalet,tipo,numcajones,numkilos,"
                Sql = Sql & "codsocio,codcampo,codvarie,fecha,hora,numnotac,CRFID) values ("
                Sql = Sql & DBSet(NumF, "N") & "," & DBSet(Tipo, "N") & "," & DBSet(Cajas, "N") & ","
                Sql = Sql & DBSet(Kilos, "N") & "," & DBSet(Codsocio, "N") & "," & DBSet(codcampo, "N") & ","
                Sql = Sql & DBSet(Codvarie, "N") & "," & DBSet(Fecha, "F") & "," & DBSet(Fecha & " " & Hora, "FH", "S") & ","
                Sql = Sql & DBSet(NumNota, "N") & "," & ValorNulo & ")"
                
                conn.Execute Sql
            Next I
            
            Sql = "update trzpalets set numcajones = numcajones + " & (CCur(NumCajones) - (Cajas * nroPalets))
            Sql = Sql & ", numkilos = numkilos + " & CCur(NumKilos) - (Kilos * nroPalets)
            Sql = Sql & " where numnotac = " & DBSet(NumNota, "N")
            Sql = Sql & " and idpalet = " & DBSet(NumF, "N")
            
            conn.Execute Sql
        End If
    
        Text1(10).Text = Format(nroPalets, "##,##0")
    
    Else ' si palots
        If Palets = 0 Then ' no me han introducido aun el numero de palots
            Exit Sub
        End If
        ' me han dado el numero de palots y los reparto
        Tipo = 1
        
        nroPalets = Val(Text1(11).Text) ' en realidad son palots
        KilosporPalet = NumKilos \ nroPalets
        RestoKilos = NumKilos - (KilosporPalet * nroPalets)
    
        For I = 1 To nroPalets
            NumF = SugerirCodigoSiguienteStr("trzpalets", "idpalet")
            
            Sql = "insert into trzpalets (idpalet,tipo,numcajones,numkilos,"
            Sql = Sql & "codsocio,codcampo,codvarie,fecha,hora,numnotac,CRFID) values ("
            Sql = Sql & DBSet(NumF, "N") & "," & DBSet(Tipo, "N") & ",0,"
            Sql = Sql & DBSet(KilosporPalet, "N") & "," & DBSet(Codsocio, "N") & "," & DBSet(codcampo, "N") & ","
            Sql = Sql & DBSet(Codvarie, "N") & "," & DBSet(Fecha, "F") & "," & DBSet(Fecha & " " & Hora, "FH", "S") & ","
            Sql = Sql & DBSet(NumNota, "N") & "," & ValorNulo & ")"
            
            conn.Execute Sql
        Next I
        
        If RestoKilos <> 0 Then ' actualizamos los kilos del ultimo palot
            Sql = "update trzpalets set numkilos = numkilos + " & DBSet(RestoKilos, "N")
            Sql = Sql & " where idpalet = " & DBSet(NumF, "N")
            
            conn.Execute Sql
        End If
    
    End If
    
'    CadB = ""
'    CargaGrid
'    BotonModificar
    
End Sub

Private Sub CargarDatosCabecera()

    Text1(0).Text = NumNota
    Text1(8).Text = Val(NumCajones)
    Text1(9).Text = NumKilos
    Text1(4).Text = Codsocio
    Text1(7).Text = codcampo
    Text1(3).Text = Codvarie
    Text1(1).Text = Fecha
    Text1(2).Text = Hora
    Text2(0).Text = DevuelveDesdeBDNew(cAgro, "rsocios", "nomsocio", "codsocio", Codsocio, "N")
    Text2(1).Text = DevuelveDesdeBDNew(cAgro, "variedades", "nomvarie", "codvarie", Codvarie, "N")

    PonerDatosCampo codcampo

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
    
    Text2(5).Text = ""
    Text2(2).Text = ""
    Text3(6).Text = ""
    Text2(3).Text = ""
    If Not Rs.EOF Then
        Text2(5).Text = DBLet(Rs.Fields(0).Value, "N") ' codigo de partida
        If Text2(5).Text <> "" Then Text2(5).Text = Format(Text2(5).Text, "0000")
        Text2(2).Text = DBLet(Rs.Fields(1).Value, "T") ' nombre de partida
        Text3(6).Text = DBLet(Rs.Fields(2).Value, "N") ' codigo de zona
        If Text3(6).Text <> "" Then Text3(6).Text = Format(Text3(6).Text, "0000")
        Text2(3).Text = DBLet(Rs.Fields(3).Value, "T") ' nombre de zona
    End If
    
    Set Rs = Nothing
    
End Sub

Private Sub PasarSigReg()
'Nos situamos en el siguiente registro
    If DataGrid1.Bookmark < adodc1.Recordset.RecordCount Then
'        DataGrid1.Row = DataGrid1.Row + 1
        DataGrid1.Bookmark = DataGrid1.Bookmark + 1
        BotonModificar
    ElseIf DataGrid1.Bookmark = adodc1.Recordset.RecordCount Then
        PonerFocoBtn Me.CmdAceptar
    End If
End Sub

Private Function ModificarExistencia() As Boolean
Dim NumReg As Long
Dim Indicador As String

    If DatosOK Then
        If ModificaDesdeFormulario(Me) Then
            TerminaBloquear
            NumReg = adodc1.Recordset.AbsolutePosition
            CargaGrid ""
            If SituarDataPosicion(adodc1, NumReg, Indicador) Then

            End If
            ModificarExistencia = True
        Else
            ModificarExistencia = False
        End If
    End If
End Function



Private Function SalirFormulario() As Boolean
Dim Sql As String
Dim B As Boolean
Dim Rs As ADODB.Recordset
Dim cad As String

    On Error GoTo eSalirFormulario

    SalirFormulario = True
    
    B = True
    
    cad = "Todos los registros deben de tener el CRFID. "
    
    Set Rs = New ADODB.Recordset
    Sql = "select crfid from trzpalets where numnotac = " & DBSet(NumNota, "N")
    
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Rs.EOF Then B = False
    
    While Not Rs.EOF And B
        If DBLet(Rs.Fields(0).Value, "T") = "" Then
            B = False
        End If
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    If Not B Then
        If MsgBox(cad & "¿ Desea salir de todos modos ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            SalirFormulario = False
        Else
            ' borramos los registros antes de salir
            Sql = "delete from trzpalets where numnotac = " & DBSet(NumNota, "N")
            conn.Execute Sql
        End If
    End If
    Exit Function
    
eSalirFormulario:
    MuestraError Err.Number, "Salir del Formulario", Err.Description
End Function
    
