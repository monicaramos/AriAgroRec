VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBodHcoFacturas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico de Facturas Retirada Bodega/Almazara"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   12690
   Icon            =   "frmBodHcoFacturas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   12690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   135
      TabIndex        =   96
      Top             =   45
      Width           =   3000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   97
         Top             =   180
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
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
      Left            =   3150
      TabIndex        =   94
      Top             =   45
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   95
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
      Left            =   10080
      TabIndex        =   93
      Top             =   225
      Width           =   1605
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4800
      Left            =   120
      TabIndex        =   13
      Tag             =   "Fecha Oferta|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
      Top             =   1680
      Width           =   12420
      _ExtentX        =   21908
      _ExtentY        =   8467
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmBodHcoFacturas.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(11)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(12)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(30)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(31)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text1(16)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1(15)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Albaranes"
      TabPicture(1)   =   "frmBodHcoFacturas.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FrameObserva"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "DataGrid2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "DataGrid1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtAux(6)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtAux(5)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtAux(4)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdObserva"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtAux(0)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtAux(1)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtAux(2)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtAux(3)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtAux3(0)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtAux3(1)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).ControlCount=   13
      Begin VB.Frame Frame2 
         Height          =   4230
         Index           =   1
         Left            =   -74940
         TabIndex        =   45
         Top             =   420
         Width           =   12175
         Begin VB.Frame FrameCliente 
            Caption         =   "Datos Socio"
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
            Height          =   2010
            Left            =   60
            TabIndex        =   67
            Top             =   150
            Width           =   12055
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
               Left            =   1125
               MaxLength       =   35
               TabIndex        =   75
               Tag             =   "Domicilio|T|N|||rbodfacturas|dirsocio||N|"
               Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
               Top             =   690
               Width           =   4800
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
               Index           =   10
               Left            =   8385
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   74
               Text            =   "Text2"
               Top             =   645
               Width           =   3285
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
               Left            =   7800
               MaxLength       =   3
               TabIndex        =   73
               Tag             =   "Forma de Pago|N|N|0|999|rbodfacturas|codforpa|000|N|"
               Text            =   "Text1"
               Top             =   645
               Width           =   540
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
               Left            =   1125
               MaxLength       =   15
               TabIndex        =   72
               Tag             =   "NIF socio|T|N|||rbodfacturas|nifsocio||N|"
               Text            =   "123456789"
               Top             =   285
               Width           =   1200
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
               Left            =   3375
               MaxLength       =   20
               TabIndex        =   71
               Tag             =   "teléfono socio|T|S|||rbodfacturas|telsoci1||N|"
               Text            =   "12345678911234567899"
               Top             =   285
               Width           =   2550
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
               Left            =   2070
               MaxLength       =   30
               TabIndex        =   70
               Tag             =   "Población|T|N|||rbodfacturas|pobsocio||N|"
               Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
               Top             =   1080
               Width           =   3855
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
               Left            =   1125
               MaxLength       =   6
               TabIndex        =   69
               Tag             =   "CPostal|T|N|||rbodfacturas|codpostal||N|"
               Text            =   "Text15"
               Top             =   1080
               Width           =   900
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
               Left            =   1125
               MaxLength       =   30
               TabIndex        =   68
               Tag             =   "Provincia|T|N|||rbodfacturas|prosocio||N|"
               Text            =   "Text1 Text1 Text1 Text1 Text22"
               Top             =   1485
               Width           =   3435
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
               Left            =   120
               TabIndex        =   81
               Top             =   690
               Width           =   960
            End
            Begin VB.Image imgBuscar 
               Height          =   240
               Index           =   3
               Left            =   7515
               ToolTipText     =   "Buscar forma de pago"
               Top             =   675
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
               Index           =   15
               Left            =   6270
               TabIndex        =   80
               Top             =   645
               Width           =   1215
            End
            Begin VB.Image imgBuscar 
               Height          =   240
               Index           =   1
               Left            =   855
               ToolTipText     =   "Buscar proveedor varios"
               Top             =   300
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.Label Label1 
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
               Index           =   20
               Left            =   120
               TabIndex        =   79
               Top             =   285
               Width           =   615
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
               Height          =   255
               Index           =   19
               Left            =   2445
               TabIndex        =   78
               Top             =   285
               Width           =   915
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
               Index           =   16
               Left            =   120
               TabIndex        =   77
               Top             =   1080
               Width           =   1005
            End
            Begin VB.Label Label1 
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
               Index           =   17
               Left            =   120
               TabIndex        =   76
               Top             =   1485
               Width           =   1005
            End
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   11
            Left            =   7380
            MaxLength       =   5
            TabIndex        =   84
            Text            =   "Text1 7"
            Top             =   1530
            Width           =   525
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   12
            Left            =   8745
            MaxLength       =   5
            TabIndex        =   83
            Text            =   "Text1 7"
            Top             =   1530
            Width           =   525
         End
         Begin VB.Frame FrameFactura 
            Height          =   2040
            Left            =   60
            TabIndex        =   46
            Top             =   2130
            Width           =   12055
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
               Index           =   30
               Left            =   7155
               MaxLength       =   15
               TabIndex        =   100
               Tag             =   "Total Factura|N|N|||rbodfacturas|totalfac|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   1275
               Width           =   1695
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
               Index           =   29
               Left            =   5355
               MaxLength       =   15
               TabIndex        =   99
               Tag             =   "Importe IVA 3|N|S|||rbodfacturas|impoiva3|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   1275
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
               Left            =   240
               MaxLength       =   15
               TabIndex        =   58
               Tag             =   "Imp.Bruto|N|N|||rbodfacturas|brutofac|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   435
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
               Left            =   3555
               MaxLength       =   15
               TabIndex        =   57
               Tag             =   "Base Imponible 1|N|N|||rbodfacturas|baseimp1|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   450
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
               Left            =   1890
               MaxLength       =   3
               TabIndex        =   56
               Tag             =   "Cod. IVA 1|N|S|0|999|rbodfacturas|codiiva1|000|N|"
               Text            =   "Text1 7"
               Top             =   450
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
               Index           =   21
               Left            =   2835
               MaxLength       =   5
               TabIndex        =   55
               Tag             =   "% IVA 1|N|S|0|99.90|rbodfacturas|porciva1|#0.00|N|"
               Text            =   "Text1 7"
               Top             =   450
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
               Index           =   27
               Left            =   5355
               MaxLength       =   15
               TabIndex        =   54
               Tag             =   "Importe IVA 1|N|N|||rbodfacturas|impoiva1|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   450
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
               Index           =   25
               Left            =   3555
               MaxLength       =   15
               TabIndex        =   53
               Tag             =   "Base Imponible 2 |N|S|||rbodfacturas|baseimp2|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   855
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
               Index           =   19
               Left            =   1890
               MaxLength       =   3
               TabIndex        =   52
               Tag             =   "Cod. IVA 2|N|S|0|999|rbodfacturas|codiiva2|000|N|"
               Text            =   "Text1 7"
               Top             =   855
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
               Index           =   22
               Left            =   2835
               MaxLength       =   5
               TabIndex        =   51
               Tag             =   "& IVA 2|N|S|0|99.90|rbodfacturas|porciva2|#0.00|N|"
               Text            =   "Text1 7"
               Top             =   855
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
               Index           =   28
               Left            =   5355
               MaxLength       =   15
               TabIndex        =   50
               Tag             =   "Importe IVA 2|N|S|||rbodfacturas|impoiva2|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   855
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
               Index           =   26
               Left            =   3555
               MaxLength       =   15
               TabIndex        =   49
               Tag             =   "Base Imponible 3|N|S|||rbodfacturas|baseimp3|#,###,###,##0.00|N|"
               Text            =   "Text1 7"
               Top             =   1275
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
               Index           =   20
               Left            =   1890
               MaxLength       =   3
               TabIndex        =   48
               Tag             =   "Cod. IVA 3|N|S|0|999|rbodfacturas|codiiva3|000|N|"
               Text            =   "Text1 7"
               Top             =   1275
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
               Index           =   23
               Left            =   2835
               MaxLength       =   5
               TabIndex        =   47
               Tag             =   "% IVA 3|N|S|0|99.90|rbodfacturas|porciva3|#0.00|N|"
               Text            =   "Text1 7"
               Top             =   1275
               Width           =   705
            End
            Begin VB.Label Label1 
               Caption         =   "Importe Bruto"
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
               Left            =   240
               TabIndex        =   66
               Top             =   195
               Width           =   1530
            End
            Begin VB.Label Label1 
               Caption         =   "Imp. IVA"
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
               Left            =   5475
               TabIndex        =   65
               Top             =   195
               Width           =   1335
            End
            Begin VB.Label Label1 
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
               Height          =   255
               Index           =   37
               Left            =   5115
               TabIndex        =   64
               Top             =   375
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
               TabIndex        =   63
               Top             =   2160
               Width           =   135
            End
            Begin VB.Label Label1 
               Caption         =   "="
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
               Index           =   38
               Left            =   6915
               TabIndex        =   62
               Top             =   1230
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
               Left            =   7125
               TabIndex        =   61
               Top             =   990
               Width           =   1755
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
               Left            =   2835
               TabIndex        =   60
               Top             =   195
               Width           =   720
            End
            Begin VB.Label Label1 
               Caption         =   "Cod. IVA"
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
               Index           =   42
               Left            =   1890
               TabIndex        =   59
               Top             =   210
               Width           =   960
            End
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. P.P"
            Height          =   255
            Index           =   25
            Left            =   6720
            TabIndex        =   86
            Top             =   1530
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Dto. Gral"
            Height          =   255
            Index           =   26
            Left            =   8055
            TabIndex        =   85
            Top             =   1530
            Width           =   735
         End
      End
      Begin VB.TextBox txtAux3 
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
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   30
         Tag             =   "Fecha Albarán|F|N|||rbodfacturas_alb|fechaalb|dd/mm/yyyy|N|"
         Text            =   "fecalbar"
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux3 
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
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   29
         Tag             =   "Nº Albarán|N|N|||rbodfacturas_alb|numalbar|0|N|"
         Text            =   "numparte"
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
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
         Left            =   3240
         MaxLength       =   12
         TabIndex        =   23
         Tag             =   "Cantidad|N|N|0||rbodfacturas_lineas|cantidad|#,###,###,##0.00|N|"
         Text            =   "cantidad"
         Top             =   4320
         Visible         =   0   'False
         Width           =   735
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
         Index           =   2
         Left            =   2160
         MaxLength       =   12
         TabIndex        =   22
         Text            =   "nomvarie"
         Top             =   4320
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
         Index           =   1
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   21
         Tag             =   "Variedad|N|N|||rbodfacturas_lineas|codvarie|000000|N|"
         Text            =   "codvarie"
         Top             =   4320
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
         Index           =   0
         Left            =   360
         MaxLength       =   12
         TabIndex        =   20
         Tag             =   "Linea|N|N|0|999|rbodfacturas_lineas|numlinea|000|N|"
         Text            =   "lin"
         Top             =   4320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdObserva 
         Enabled         =   0   'False
         Height          =   375
         Left            =   11625
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2130
         Visible         =   0   'False
         Width           =   375
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
         Left            =   4095
         MaxLength       =   12
         TabIndex        =   24
         Tag             =   "Precio|N|N|0|999999.0000|rbodfacturas_lineas|precioar|###,##0.0000|N|"
         Text            =   "Precio"
         Top             =   4320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
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
         Left            =   5760
         MaxLength       =   5
         TabIndex        =   25
         Tag             =   "Dto|N|N|0|99.90|rbodfacturas_lineas|dtolinea|##,##0.000|N|"
         Text            =   "Dto"
         Top             =   4320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
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
         Left            =   7080
         MaxLength       =   12
         TabIndex        =   28
         Tag             =   "Importe|N|N|0||rbodfacturas_lineas|importel|#,###,###,##0.00|N|"
         Text            =   "Importe"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmBodHcoFacturas.frx":0044
         Height          =   2025
         Left            =   240
         TabIndex        =   15
         Top             =   2670
         Width           =   11935
         _ExtentX        =   21061
         _ExtentY        =   3572
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmBodHcoFacturas.frx":0059
         Height          =   1995
         Left            =   225
         TabIndex        =   16
         Top             =   495
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   3519
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
         Caption         =   "Albaranes de la Factura"
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
      Begin VB.Frame FrameObserva 
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
         Height          =   2085
         Left            =   6330
         TabIndex        =   17
         Tag             =   "Observación 4|T|S|||scafac1|observa4||N|"
         Top             =   450
         Width           =   5845
         Begin VB.TextBox Text3 
            Height          =   1470
            Index           =   4
            Left            =   150
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   18
            Tag             =   "Observación 1|T|S|||rbodfacturas_alb|observac||N|"
            Top             =   360
            Width           =   5530
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   -74250
         MaxLength       =   15
         TabIndex        =   87
         Text            =   "Text1 7"
         Top             =   3495
         Width           =   1365
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   -72450
         MaxLength       =   15
         TabIndex        =   88
         Text            =   "Text1 7"
         Top             =   3495
         Width           =   1365
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
         Index           =   31
         Left            =   -72690
         TabIndex        =   92
         Top             =   3420
         Width           =   135
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
         Index           =   30
         Left            =   -74490
         TabIndex        =   91
         Top             =   3420
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "Imp. Dto Gn"
         Height          =   255
         Index           =   12
         Left            =   -72330
         TabIndex        =   90
         Top             =   3300
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Imp. Dto PP"
         Height          =   255
         Index           =   11
         Left            =   -74130
         TabIndex        =   89
         Top             =   3300
         Width           =   855
      End
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   4485
      MaxLength       =   30
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   2205
      Width           =   660
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   1
      Left            =   5145
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   41
      Text            =   "Text2"
      Top             =   2205
      Width           =   3525
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   4485
      MaxLength       =   30
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   1845
      Width           =   660
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   0
      Left            =   5145
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   39
      Text            =   "Text2"
      Top             =   1845
      Width           =   3525
   End
   Begin VB.Frame Frame2 
      Height          =   885
      Index           =   0
      Left            =   120
      TabIndex        =   31
      Top             =   780
      Width           =   12415
      Begin VB.CheckBox Check1 
         Caption         =   "Impreso"
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
         Left            =   11040
         TabIndex        =   8
         Tag             =   "Contabilizado|N|N|0|1|rbodfacturas|impreso||N|"
         Top             =   375
         Width           =   1185
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Aridoc"
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
         Left            =   10050
         TabIndex        =   7
         Tag             =   "Contabilizado|N|N|0|1|rbodfacturas|pasaridoc||N|"
         Top             =   375
         Width           =   1275
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
         Index           =   17
         Left            =   210
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Tipo|T|N|||rbodfacturas|codtipom||S|"
         Top             =   390
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
         Index           =   3
         Left            =   4530
         MaxLength       =   40
         TabIndex        =   4
         Tag             =   "Nombre Socio|T|N|||rbodfacturas|nomsocio||N|"
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   390
         Width           =   3750
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
         Left            =   3630
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Cod. socio|N|N|0|999999|rbodfacturas|codsocio|000000|S|"
         Text            =   "Text1"
         Top             =   390
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
         Left            =   2250
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Factura|F|N|||rbodfacturas|fecfactu|dd/mm/yyyy|S|"
         Top             =   390
         Width           =   1305
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
         Left            =   1020
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Nº Factura|N|N|||rbodfacturas|numfactu||S|"
         Text            =   "Text1 7"
         Top             =   390
         Width           =   1125
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Contabilizado"
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
         Left            =   8355
         TabIndex        =   6
         Tag             =   "Contabilizado|N|N|0|1|rbodfacturas|intconta||N|"
         Top             =   375
         Width           =   1725
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   2
         Left            =   210
         TabIndex        =   82
         Top             =   135
         Width           =   615
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
         Left            =   3630
         TabIndex        =   34
         Top             =   150
         Width           =   555
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   4215
         ToolTipText     =   "Buscar socio"
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "F. Factura"
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
         Left            =   2250
         TabIndex        =   33
         Top             =   135
         Width           =   1095
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
         Left            =   1050
         TabIndex        =   32
         Top             =   150
         Width           =   1305
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   13
      Left            =   3555
      MaxLength       =   4
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   900
      Width           =   540
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   13
      Left            =   4140
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   36
      Text            =   "Text2"
      Top             =   900
      Width           =   3285
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Height          =   315
      Index           =   17
      Left            =   7500
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   27
      Text            =   "ABCDKFJADKSFJAK"
      Top             =   5430
      Visible         =   0   'False
      Width           =   1725
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   480
      Top             =   5160
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
      Left            =   1920
      Top             =   5160
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
      Index           =   16
      Left            =   2385
      Locked          =   -1  'True
      MaxLength       =   60
      TabIndex        =   26
      Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
      Top             =   6780
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.Frame Frame1 
      Height          =   525
      Index           =   0
      Left            =   105
      TabIndex        =   11
      Top             =   6615
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
         Left            =   240
         TabIndex        =   12
         Top             =   180
         Width           =   1755
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
      Left            =   11505
      TabIndex        =   9
      Top             =   6705
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
      Left            =   10350
      TabIndex        =   5
      Top             =   6705
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
      Left            =   11490
      TabIndex        =   10
      Top             =   6705
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
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   12075
      TabIndex        =   98
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
   Begin VB.Image imgBuscar 
      Enabled         =   0   'False
      Height          =   240
      Index           =   2
      Left            =   4455
      Picture         =   "frmBodHcoFacturas.frx":006E
      ToolTipText     =   "Buscar población"
      Top             =   900
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   6
      Left            =   4125
      Picture         =   "frmBodHcoFacturas.frx":0170
      ToolTipText     =   "Buscar trabajador"
      Top             =   2220
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Trabajador Pedido"
      Height          =   255
      Index           =   9
      Left            =   2565
      TabIndex        =   44
      Top             =   2220
      Width           =   1425
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   5
      Left            =   4140
      Picture         =   "frmBodHcoFacturas.frx":0272
      ToolTipText     =   "Buscar trabajador"
      Top             =   1845
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Trabajador Albaran"
      Height          =   255
      Index           =   21
      Left            =   2565
      TabIndex        =   43
      Top             =   1890
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Trabajador"
      Height          =   255
      Index           =   1
      Left            =   2340
      TabIndex        =   38
      Top             =   900
      Width           =   855
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   4
      Left            =   3270
      Picture         =   "frmBodHcoFacturas.frx":0374
      ToolTipText     =   "Buscar trabajador"
      Top             =   930
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Lote"
      Height          =   255
      Index           =   3
      Left            =   7500
      TabIndex        =   35
      Top             =   5250
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Ampliación Línea"
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
      Index           =   35
      Left            =   2385
      TabIndex        =   14
      Top             =   6510
      Visible         =   0   'False
      Width           =   1335
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
      Begin VB.Menu mnLineas 
         Caption         =   "&Lineas"
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^L
         Visible         =   0   'False
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
Attribute VB_Name = "frmBodHcoFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)


Public Factura As String ' cuando venimos de documentos de proveedores
Public Tipo As Byte

'Si se llama de la busqueda en el frmAlmMovimArticulos se accede
'a las tablas del Albaran o de Facturas de movimiento seleccionado (solo consulta)
Public hcoCodMovim As String 'cod. movim
Public hcoFechaMovim As Date 'Fecha del Movim
Public hcoCodSocio As Integer 'Codigo de Socio

'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
'--monica
'Private WithEvents frmCP As frmCPostal 'Codigos Postales

Private WithEvents frmSoc As frmManSocios  'Form Mto socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmFPago As frmForpaConta 'Form Mto Formas de Pago de la conta de la seccion que toque
Attribute frmFPago.VB_VarHelpID = -1


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

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

'Si el cliente mostrado es de Varios o No
Dim EsDeVarios As Boolean


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
Private BuscaChekc As String

Dim vSeccion As CSeccion



Private Sub Check1_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "Check1(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "check1(" & Index & ")|"
    End If

End Sub

Private Sub Check1_GotFocus(Index As Integer)
    PonerFocoChk Me.Check1(Index)
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 4  'MODIFICAR
            If DatosOK Then
               If ModificarFactura Then
                    TerminaBloquear
'                    PosicionarData
               Else
                    '---- Laura 24/10/2006
                    'como no hemos modificado dejamos la fecha como estaba ya que ahora se puede modificar
                    Text1(1).Text = Me.Data1.Recordset!fecfactu
               End If
               PosicionarData
            End If
            
         Case 5 'InsertarModificar LINEAS
            'Actualizar el registro en la tabla de lineas 'slialb'
            If ModificaLineas = 1 Then 'INSERTAR lineas Albaran
'                PrimeraLin = False
'                If Data2.Recordset.EOF = True Then PrimeraLin = True
'                If InsertarLinea(NumLinea) Then
'                    'Comprobar si el Articulo tiene control de Nº de Serie
'                    ComprobarNSeriesLineas NumLinea
'                    If PrimeraLin Then
'                        CargaGrid DataGrid1, Data2, True
'                    Else
'                        CargaGrid2 DataGrid1, Data2
'                    End If
'                    BotonAnyadirLinea
'                End If
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                    TerminaBloquear
                    NumRegElim = Data2.Recordset.AbsolutePosition
                    
                    CargaGrid2 DataGrid1, Data2
                    ModificaLineas = 0
                    PonerBotonCabecera True
                    BloquearTxt Text2(16), True
                    BloquearTxt Text2(17), True
           
                    LLamaLineas Modo, 0, "DataGrid1"
                    PosicionarData
                    If (Not Data2.Recordset.EOF) And (Not Data2.Recordset.BOF) Then
                        SituarDataPosicion Data2, NumRegElim, ""
                    End If
                End If
                Me.DataGrid1.Enabled = True
                Me.DataGrid2.Enabled = True
            End If
    End Select
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 0, 1 'Busqueda, Insertar
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
            LLamaLineas Modo, 0, "DataGrid2"
            PonerFoco Text1(0)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
        Case 5 'Lineas Detalle
            TerminaBloquear
            BloquearTxt Text2(16), True
            BloquearTxt Text2(17), True
'            If ModificaLineas = 1 Then 'INSERTAR
'                ModificaLineas = 0
'                DataGrid1.AllowAddNew = False
'                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
'            End If
            ModificaLineas = 0
            LLamaLineas Modo, 0, "DataGrid1"
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
    End Select
End Sub


Private Sub BotonBuscar()
Dim anc As Single

    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PonerModo 1
        
        'poner los txtaux para buscar por lineas de albaran
        anc = DataGrid2.Top
        If DataGrid2.Row < 0 Then
            anc = anc + 440
        Else
            anc = anc + DataGrid2.RowTop(DataGrid2.Row) + 20
        End If
        LLamaLineas Modo, anc, "DataGrid2"
        
        'Si pasamos el control aqui lo ponemos en amarillo
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
Dim cad As String
    
    Select Case Tipo
        Case 0
            cad = "mid(rbodfacturas.codtipom,2,2) = 'ZA'"
        Case 1
            cad = "mid(rbodfacturas.codtipom,2,2) = 'AB'"
    End Select

    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia cad
    Else
        LimpiarCampos

        LimpiarDataGrids
        CadenaConsulta = "Select rbodfacturas.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla & " where " & cad & Ordenacion

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

    'solo se puede modificar la factura si no esta contabilizada
    '++monica:añadida la condicion de solo si hay contabilidad
    If vParamAplic.NumeroConta <> 0 Then
'        If Me.Check1.Value = 1 Then
'            TerminaBloquear
'            Exit Sub
'        End If
        
        If FactContabilizada Then
            TerminaBloquear
            Exit Sub
        End If
    End If
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFocoChk Me.Check1(0)
        
    'Si es proveedor de Varios no se pueden modificar sus datos
'--monica
'    DeVarios = EsProveedorVarios(Text1(2).Text)
    BloquearDatosSocio (DeVarios)
End Sub


Private Sub BotonModificarLinea()
'Modificar una linea
Dim vWhere As String
Dim anc As Single
Dim J As Byte
On Error GoTo eModificarLinea




    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then
        TerminaBloquear
        Exit Sub '1= Insertar
    End If
    
    If Data2.Recordset.EOF Then
        TerminaBloquear
        Exit Sub
    End If
    
    vWhere = ObtenerWhereCP(False)
    vWhere = vWhere & " AND numalbar=" & Data3.Recordset.Fields!numalbar & ""
    vWhere = vWhere & " and numlinea=" & Data2.Recordset!numlinea
    If Not BloqueaRegistro(NomTablaLineas, vWhere) Then
        TerminaBloquear
        Exit Sub
    End If

    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        J = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, J
        DataGrid1.Refresh
    End If
    
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 210
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 20
    End If

    For J = 0 To 2
        txtAux(J).Text = DataGrid1.Columns(J + 5).Text
    Next J
    Text2(16).Text = DataGrid1.Columns(J + 5).Text
    For J = J + 1 To 8
        txtAux(J - 1).Text = DataGrid1.Columns(J + 5).Text
    Next J
'--monica
'    Text2(17).Text = DataGrid1.Columns(14).Text
    
    ModificaLineas = 2 'Modificar
    LLamaLineas ModificaLineas, anc, "DataGrid1"
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR LINEAS"
    PonerBotonCabecera False
    BloquearTxt Text2(16), False 'Campo Ampliacion Linea
    BloquearTxt Text2(17), False 'Campo Ampliacion Linea
'    PonerFoco txtAux(4)
    PonerFoco Text2(16)
    Me.DataGrid1.Enabled = False

eModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub LLamaLineas(xModo As Byte, Optional alto As Single, Optional grid As String)
Dim jj As Integer
Dim B As Boolean
'
'    Select Case grid
'        Case "DataGrid1"
'            DeseleccionaGrid Me.DataGrid1
'            'PonerModo xModo + 1
'
'            b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Lineas
'
'            For jj = 0 To txtAux.Count - 1
'                If jj = 4 Or jj = 5 Or jj = 6 Or jj = 7 Then
'                    txtAux(jj).Height = DataGrid1.RowHeight
'                    txtAux(jj).Top = alto
'                    txtAux(jj).visible = b
'                End If
'            Next jj
'
        If grid = "DataGrid2" Then
            DeseleccionaGrid Me.DataGrid2
            B = (xModo = 1)
             For jj = 0 To txtAux3.Count - 1
                txtAux3(jj).Height = DataGrid2.RowHeight
                txtAux3(jj).Top = alto
                txtAux3(jj).visible = B
            Next jj
        End If
'    End Select
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Facturas (scafac)
' y los registros correspondientes de las tablas cab. albaranes (scafac1)
' y las lineas de la factura (slifac)
Dim cad As String
Dim NumPedElim As Long
On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    'solo se puede eliminar si no esta en la contabilidad
    If Me.Check1(0).Value = 1 Then Exit Sub
    
'    'solo se puede modificar la factura si no esta contabilizada
'    If FactContabilizada Then Exit Sub
    
    cad = "Cabecera de Facturas." & vbCrLf
    cad = cad & "-----------------------------------" & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar la Factura:            "
    cad = cad & vbCrLf & "Socio  :  " & Text1(2).Text & " - " & Text1(3).Text
    cad = cad & vbCrLf & "NºFact.:  " & Text1(0).Text
    cad = cad & vbCrLf & "Fecha  :  " & Format(Text1(1).Text, "dd/mm/yyyy")

    cad = cad & vbCrLf & vbCrLf & " ¿Desea Eliminarla? "
    
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        NumPedElim = Data1.Recordset.Fields(1).Value
        
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
    
EEliminar:
        Screen.MousePointer = vbDefault
        If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Factura", Err.Description
End Sub



Private Sub BloqueaText3()
Dim I As Byte
    'bloquear los Text3 que son las lineas de scafpa
    For I = 0 To 1
        BloquearTxt Text3(I), (Modo <> 4)
    Next I
    If Me.FrameObserva.visible Then
        For I = 4 To 8
            BloquearTxt Text3(I), (Modo <> 4)
        Next I
    End If
    'numpedpr, fecpedpr siempre bloqueados
    For I = 2 To 3
        BloquearTxt Text3(I), True
    Next I
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        DataGrid2.Enabled = True
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


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo Error1

    If Not Data2.Recordset.EOF Then
        If ModificaLineas <> 1 Then
            Text2(16).Text = DBLet(Data2.Recordset.Fields!ampliaci)
'            Text2(17).Text = DBLet(Data2.Recordset.Fields!numlotes)
        End If
    Else
        Text2(16).Text = ""
        Text2(17).Text = ""
    End If

Error1:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim I As Byte

    If Not Data3.Recordset.EOF Then
'        Text3(0).Text = DBLet(Data3.Recordset.Fields!codtrab2, "T")
'        Text3_LostFocus (0)
'        Text3(1).Text = DBLet(Data3.Recordset.Fields!codtrab1, "T")
'        Text3_LostFocus (1)

'--monica
'        Text3(2).Text = DBLet(Data3.Recordset.Fields!numpedpr, "N")
'        If Text3(2).Text <> "0" Then
'            FormateaCampo Text3(2)
'        Else
'            Text3(2).Text = ""
'        End If
'        Text3(3).Text = DBLet(Data3.Recordset.Fields!fecpedpr, "F")
        
        'Observaciones
        Text3(4).Text = DBLet(Data3.Recordset.Fields!Observac, "T")
'        Text3(5).Text = DBLet(Data3.Recordset.Fields!observa2, "T")
'        Text3(6).Text = DBLet(Data3.Recordset.Fields!observa3, "T")
'        Text3(7).Text = DBLet(Data3.Recordset.Fields!observa4, "T")
'        Text3(8).Text = DBLet(Data3.Recordset.Fields!observa5, "T")
        
        'Datos de la tabla
        CargaGrid DataGrid1, Data2, True
    Else
        
        Text3(4).Text = ""
        Text2(0).Text = ""
        Text2(1).Text = ""
        'Datos de la tabla slipre
        CargaGrid DataGrid1, Data2, False
    End If
    
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
    If hcoCodMovim <> "" And Not Data1.Recordset.EOF Then PonerCadenaBusqueda
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim I As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    ' ICONITOS DE LA BARRA
    btnPrimero = 15
'    With Me.Toolbar1
'        .HotImageList = frmPpal.imgListComun_OM
'        .DisabledImageList = frmPpal.imgListComun_BN
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1   'Botón Buscar
'        .Buttons(2).Image = 2   'Botón Todos
'        .Buttons(5).Image = 4   'Modificar
'        .Buttons(6).Image = 5   'Borrar
'        .Buttons(9).Image = 15 'Mto Lineas Ofertas
'        .Buttons(10).Image = 10 'Imprimir
'        .Buttons(12).Image = 11  'Salir
'        .Buttons(btnPrimero).Image = 6  'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 'Último
'    End With

    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'l'1 i el 2 son separadors
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2   'Totss
        'el 5 i el 6 son separadors
'        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
        'el 10  son separadors
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

    Me.SSTab1.Tab = 0
      
    'cargar IMAGES de busqueda
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
      
    Select Case Tipo
        Case 0 ' almazara
            ConexionConta vParamAplic.SeccionAlmaz
        Case 1 ' bodega
            ConexionConta vParamAplic.SeccionBodega
    End Select
      
    LimpiarCampos   'Limpia los campos TextBox
     
    'cargar icono de observaciones de los albaranes de factura
'    CargarICO Me.cmdObserva, "message.ico"
    Me.cmdObserva.Picture = frmPpal.imgListComun.ListImages(12).Picture '--monica antes 41
'    Me.FrameObserva.visible = False
    Me.cmdObserva.ToolTipText = "ver observaciones Factura"
    
    VieneDeBuscar = False
            
    '## A mano
    NombreTabla = "rbodfacturas"
    NomTablaLineas = "rbodfacturas_lineas" 'Tabla lineas de Facturacion
    Ordenacion = " ORDER BY rbodfacturas.fecfactu desc, rbodfacturas.numfactu "
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    CadenaConsulta = "Select * from " & NombreTabla
    
    If hcoCodMovim <> "" Then
        'Se llama desde Dobleclick en frmAlmMovimArticulos
'        CadenaConsulta = CadenaConsulta & " WHERE numalbar='" & hcoCodMovim & "' AND fechaalb= """ & Format(hcoFechaMovim, "yyyy-mm-dd") & """"
'        CadenaConsulta = CadenaConsulta & " AND codprove=" & hcoCodProve
        If Factura <> "" Then
            CadenaConsulta = CadenaConsulta & " WHERE codsocio = " & hcoCodSocio & " and numfactu = '" & hcoCodMovim & "' and "
            CadenaConsulta = CadenaConsulta & " fecfactu = '" & Format(hcoFechaMovim, "yyyy-mm-dd") & "'"
        Else
            CadenaConsulta = CadenaConsulta & ObtenerSelFactura
        End If
    Else
        CadenaConsulta = CadenaConsulta & " where numfactu=-1"
    End If
    
    Select Case Tipo
        Case 0 ' almazara
            CadenaConsulta = CadenaConsulta & " and mid(codtipom,2,2) = 'ZA'"
        Case 1 ' bodega
            CadenaConsulta = CadenaConsulta & " and mid(codtipom,2,2) = 'AB'"
    End Select
    
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    If hcoCodMovim = "" Then
        If DatosADevolverBusqueda = "" Then
            PonerModo 0
        Else
            PonerModo 1
        End If
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PrimeraVez = False
    Else
         PonerModo 0
    End If
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    
    Me.Check1(0).Value = 0
    Me.Check1(1).Value = 0
    Me.Check1(2).Value = 0
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Not vSeccion Is Nothing Then
        vSeccion.CerrarConta
        Set vSeccion = Nothing
    End If
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(17), CadenaDevuelta, 1)
        CadB = Aux
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 2)
        CadB = CadB & " and " & Aux
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 3)
        CadB = CadB & " and " & Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB
        CadenaConsulta = CadenaConsulta & " GROUP BY rbodfacturas.codtipom, rbodfacturas.numfactu, rbodfacturas.fecfactu "
        CadenaConsulta = CadenaConsulta & " " & Ordenacion
        PonerCadenaBusqueda
'        Text1(0).Text = RecuperaValor(CadenaDevuelta, 2)
    End If
    Screen.MousePointer = vbDefault
End Sub

'--monica
'Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
''Formulario Mantenimiento C. Postales
'Dim Indice As Byte
'Dim devuelve As String
'
'        Indice = 7
'        Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
'        Text1(Indice + 1).Text = ObtenerPoblacion(Text1(Indice).Text, devuelve)  'Poblacion
'        'provincia
'        Text1(Indice + 2).Text = devuelve
'End Sub

Private Sub frmFPago_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim Indice As Byte
    Indice = 10
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Pago
    Text2(10).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub


Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Scoios
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1)  'Cod socio
End Sub


Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Integer

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. socio
            PonerFoco Text1(2)
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|"
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            Indice = 2
            PonerFoco Text1(Indice)
      
         Case 3 'Forma de Pago
            Indice = 10
            AbrirFrmForpaConta Indice
            PonerFoco Text1(Indice)
            
    End Select
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub mnBuscar_Click()
    Me.SSTab1.Tab = 0
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
'Imprimir Factura
Dim frmBodRe As frmBodReimpre

    Set frmBodRe = New frmBodReimpre

    frmBodRe.Tipo = 0
    frmBodRe.Label1 = "Reimpresión de Facturas de Almazara"
    frmBodRe.txtCodigo(4) = Text1(0).Text
    frmBodRe.txtCodigo(5) = Text1(0).Text
    frmBodRe.txtCodigo(2) = Text1(1).Text
    frmBodRe.txtCodigo(3) = Text1(1).Text
    frmBodRe.txtCodigo(0) = Text1(2).Text
    frmBodRe.txtCodigo(1) = Text1(2).Text
    frmBodRe.txtNombre(0) = Text1(3).Text
    frmBodRe.txtNombre(1) = Text1(3).Text
    
    frmBodRe.Show vbModal
    
    Set frmBodRe = Nothing
    
End Sub


Private Sub mnLineas_Click()
    BotonMtoLineas 1, "Facturas"
End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
        'bloquea la tabla cabecera de factura: scafac
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: scafac1
            If BloqueaAlbxFac Then
                If BloqueaLineasFac Then BotonModificarLinea
            End If
        End If
         
    Else   'Modificar Pedido
        'bloquea la tabla cabecera de factura: scafpc
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: scafpa
            If BloqueaAlbxFac Then BotonModificar
        End If
    End If
End Sub


Private Function BloqueaAlbxFac() As Boolean
'bloquea todos los albaranes de la factura
Dim Sql As String
On Error GoTo EBloqueaAlb

    BloqueaAlbxFac = False
    'bloquear cabecera albaranes x factura
    Sql = "select * FROM rbodfacturas_alb "
    Sql = Sql & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute Sql, , adCmdText
    BloqueaAlbxFac = True

EBloqueaAlb:
    If Err.Number <> 0 Then BloqueaAlbxFac = False
End Function


Private Function BloqueaLineasFac() As Boolean
'bloquea TODAS las lineas de la factura
Dim Sql As String
    
    On Error GoTo EBloqueaLin

    BloqueaLineasFac = False
    
    'bloquear cabecera albaranes x factura
    Sql = "select * FROM rbodfacturas_lineas "
    Sql = Sql & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute Sql, , adCmdText
    BloqueaLineasFac = True

EBloqueaLin:
    If Err.Number <> 0 Then BloqueaLineasFac = False
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
    If Index = 9 Then HaCambiadoCP = False 'CPostal
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
        
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1 'Fecha factura
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
                
'--monica
'        Case 13 'Cod trabajador
'            Text2(Index).Text = PonerNombreDeCod(Text1(Index), cAgro, "straba", "nomtraba", "codtraba")

        Case 2 'Cod. socio
            If Modo = 1 Then 'Modo=1 Busqueda
                Text1(Index + 1).Text = PonerNombreDeCod(Text1(Index), "rsocios", "nomsocio")
            Else
                PonerDatosSocio (Text1(Index).Text)
            End If
        
        Case 4 'NIF
'            If Not EsDeVarios Then Exit Sub
'            If Modo = 4 Then 'Modificar
'                'si no se ha modificado el nif del cliente no hacer nada
'                If Text1(4).Text = DBLet(Data1.Recordset!nifSocio, "T") Then
'                    Exit Sub
'                End If
'            End If
'            PonerDatosProveVario (Text1(Index).Text)


'--monica
'        Case 7 'Cod. Postal
'            If Text1(Index).Locked Then Exit Sub
'            If Text1(Index).Text = "" Then
'                Text1(Index + 1).Text = ""
'                Text1(Index + 2).Text = ""
'                Exit Sub
'            End If
'            If (Not VieneDeBuscar) Or (VieneDeBuscar And HaCambiadoCP) Then
'                 Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, devuelve)
'                 Text1(Index + 2).Text = devuelve
'            End If
'            VieneDeBuscar = False
        
        
        Case 10 'Forma de Pago
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "forpago", "nomforpa")
            Else
                Text2(Index).Text = ""
            End If
            
'        Case 11, 12 'Descuentos
'            If Modo = 4 Then 'comprobar que el dato a cambiado
'                If Index = 11 Then
'                    If CCur(Text1(Index).Text) = CCur(Data1.Recordset!DtoPPago) Then Exit Sub
'                ElseIf Index = 12 Then
'                    If CCur(Text1(Index).Text) = CCur(Data1.Recordset!DtoGnral) Then Exit Sub
'                End If
'            End If
'
'            If Modo = 3 Or Modo = 4 Then
'                If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 4 'Tipo 4: Decimal(4,2)
'                If Not ActualizarDatosFactura Then
'                   If Index = 11 Then Text1(Index).Text = Data1.Recordset!DtoPPago
'                   If Index = 12 Then Text1(Index).Text = Data1.Recordset!DtoGnral
'                   FormateaCampo Text1(Index)
'                End If
'            End If
    End Select
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda(Me)
    If CadB <> "" Then CadB = CadB & " and "
    
    Select Case Tipo
        Case 0 'almazara
            CadB = CadB & "mid(rbodfacturas.codtipom,2,2) = 'ZA'"
        Case 1 'bodega
            CadB = CadB & "mid(rbodfacturas.codtipom,2,2) = 'AB'"
    End Select
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select " & NombreTabla & ".* from " & NombreTabla & " LEFT OUTER JOIN rbodfacturas_alb ON " & NombreTabla & ".codtipom=rbodfacturas_alb.codtipom AND " & NombreTabla & ".numfactu=rbodfacturas_alb.numfactu AND " & NombreTabla & ".fecfactu=rbodfacturas_alb.fecfactu "
        CadenaConsulta = CadenaConsulta & " WHERE " & CadB
        CadenaConsulta = CadenaConsulta & " GROUP BY rbodfacturas.codtipom, rbodfacturas.numfactu, rbodfacturas.fecfactu "
        CadenaConsulta = CadenaConsulta & " " & Ordenacion
        
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String
Dim devuelve As String
    
    'Llamamos a al form
    '##A mano
    cad = ""
        cad = cad & ParaGrid(Text1(17), 10, "Tipo Fac.")
        cad = cad & ParaGrid(Text1(0), 18, "Nº Factura")
        cad = cad & ParaGrid(Text1(1), 15, "Fecha Fac.")
        cad = cad & ParaGrid(Text1(2), 12, "Socio")
        cad = cad & ParaGrid(Text1(3), 45, "Nombre Socio")
        tabla = NombreTabla
        Titulo = "Facturas Retirada Bodega/Almazara"
        devuelve = "0|1|2|"
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vtabla = tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
'        frmB.vDevuelve = "0|1|"
        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vSelElem = 0
'--monica
'        frmB.vConexionGrid = conAri  'Conexión a BD: Ariges

'        If Not EsCabecera Then frmB.Label1.FontSize = 11
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
        'End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass

    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then PonerFoco Text1(0)
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        LLamaLineas Modo, 0, "DataGrid2"
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
'Carga el grid de los AlbaranesxFactura, es decir, la tabla scafpc de la factura seleccionada
On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass
    
    'Datos de la tabla albaranes x factura: scafpa
    CargaGrid DataGrid2, Data3, True
   
    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim BrutoFac As Single

    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    'Poner la base imponible (impbruto - dtoppago - dtognral
    BrutoFac = CSng(Text1(14).Text) - CSng(Text1(15).Text) - CSng(Text1(16).Text)
'    Text1(17).Text = Format(BrutoFac, FormatoImporte)
    
    'poner descripcion campos
    Text2(10).Text = PonerNombreDeCod(Text1(10), "forpago", "nomforpa")
'--monica
'    Text2(13).Text = PonerNombreDeCod(Text1(13), conAri, "straba", "nomtraba", "codtraba")
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    
'++monica
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
'++
    
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    If Err.Number <> 0 Then Err.Clear
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte, NumReg As Byte
Dim B As Boolean

    On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar
    '## No tiene el boton modificar y no utiliza la funcion general
    ActualizarToolbar Modo, Kmodo
    BuscaChekc = ""
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    B = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Or (Factura <> "") Then
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
    '---- laura 24/10/2006: si ponemos las claves de la tabla con ON UPDATE CASCADE
    'podemos permitir modificar la fecha de la factura que es clave primaria
'    If Modo = 4 Then BloquearTxt Text1(1), False
    
    For I = 0 To Check1.Count - 1
        Me.Check1(I).Enabled = (Modo = 1) '  Or Modo = 3 Or Modo = 4)
    Next I
    
    B = (Modo <> 1)
    'Campos Nº Factura bloqueado y en azul
    BloquearTxt Text1(0), B, True
    BloquearTxt Text1(17), B, True
    
    BloquearTxt Text1(3), B   'referencia
    
    'Importes siempre bloqueados
    For I = 14 To 30
        If I <> 17 Then BloquearTxt Text1(I), (Modo <> 1)
    Next I

    'Campo B.Imp y Imp. IVA siempre en azul
'    Text1(17).BackColor = &HFFFFC0
    Text1(27).BackColor = &HFFFFC0
    Text1(28).BackColor = &HFFFFC0
    Text1(29).BackColor = &HFFFFC0
    Text1(30).BackColor = &HC0C0FF
    
    'bloquear los Text3 que son las lineas de scafac1
    BloqueaText3
    
    'Si no es modo lineas Boquear los TxtAux
    For I = 0 To txtAux.Count - 1
        BloquearTxt txtAux(I), (Modo <> 5)
    Next I
    BloquearTxt txtAux(8), True
    
    'Si no es modo Busqueda Bloquear los TxtAux3 (son los txtaux de los albaranes de factura)
    For I = 0 To txtAux3.Count - 1
        BloquearTxt txtAux3(I), (Modo <> 1)
    Next I
    
    'ampliacion linea
    B = (Modo = 5) And Me.DataGrid1.visible
    'Modo Linea de Albaranes
    Me.Label1(35).visible = B
    Me.Label1(3).visible = B
    Me.Text2(16).visible = B
    Me.Text2(17).visible = B
    BloquearTxt Text2(16), (Modo <> 5) Or (Modo = 5 And ModificaLineas <> 1)
    BloquearTxt Text2(17), (Modo <> 5) Or (Modo = 5 And ModificaLineas <> 1)

    '---------------------------------------------
    B = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    
    
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Enabled = B
    Next I
    Me.imgBuscar(0).Enabled = (Modo = 1)
    Me.imgBuscar(1).visible = False
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
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

Private Function DatosOK() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim B As Boolean
On Error GoTo EDatosOK

    DatosOK = False
    
    'Para que no den errores los 0's de los importes de dtos
    ComprobarDatosTotales
        
    'comprobamos datos OK de la tabla scafac
    B = CompForm(Me) 'Comprobar formato datos ok de la cabecera: opcion=1
    If Not B Then Exit Function
       
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

    For I = 0 To txtAux.Count - 1
        If I = 4 Or I = 5 Or I = 6 Then
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


Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 16 And KeyCode = 40 Then 'campo Amliacion Linea y Flecha hacia abajo
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 17 And KeyAscii = 13 Then 'campo nº de lote y ENTER
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Text3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Text3_LostFocus(Index As Integer)
    Select Case Index
'--monica
'        Case 0, 1 'trabajador
'            Text2(Index).Text = PonerNombreDeCod(Text3(Index), conAri, "straba", "nomtraba", "codtraba", "Cod. Trabajador", "N")
        Case 8 'observa 5
            PonerFocoBtn Me.cmdAceptar
    End Select
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
        
        Case 5  'Buscar
            mnBuscar_Click
        Case 6  'Todos
            BotonVerTodos

'        Case 9  'Lineas
'            mnLineas_Click
        
        Case 8 'Imprimir Albaran
            mnImprimir_Click
'        Case 12    'Salir
'            mnSalir_Click
'        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
'            Desplazamiento (Button.Index - btnPrimero)
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
    
    
Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Albaran: slialb
Dim Sql As String
Dim vWhere As String
Dim B As Boolean

    On Error GoTo eModificarLinea

    ModificarLinea = False
    If Data2.Recordset.EOF Then Exit Function
    
    vWhere = ObtenerWhereCP(True)
    vWhere = vWhere & " AND numalbar='" & Data3.Recordset.Fields!numalbar & "'"
    vWhere = vWhere & " AND numlinea=" & Data2.Recordset.Fields!numlinea
    
    If DatosOkLinea() Then
        Sql = "UPDATE " & NomTablaLineas & " SET "
        Sql = Sql & " ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
        Sql = Sql & "precioar= " & DBSet(txtAux(4).Text, "N") & ", "
        Sql = Sql & "dtoline1= " & DBSet(txtAux(5).Text, "N") & ", dtoline2= " & DBSet(txtAux(6).Text, "N") & ", "
        Sql = Sql & "importel= " & DBSet(txtAux(7).Text, "N")
        Sql = Sql & ", numlotes=" & DBSet(Text2(17).Text, "T")
        Sql = Sql & vWhere
    End If
    
    If Sql <> "" Then
        'actualizar la factura y vencimientos
        B = ModificarFactura(Sql)
        ModificarLinea = B
    End If
    
eModificarLinea:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Modificar Lineas Factura" & vbCrLf & Err.Description
        B = False
    End If
    ModificarLinea = B
End Function


Private Sub PonerBotonCabecera(B As Boolean)
'Pone el boton de Regresar a la Cabecera si pasamos a MAntenimiento de Lineas
'o Pone los botones de Aceptar y cancelar en Insert,update o delete lineas
On Error Resume Next

    Me.cmdAceptar.visible = Not B
    Me.cmdCancelar.visible = Not B
    Me.cmdRegresar.visible = B
    Me.cmdRegresar.Caption = "Cabecera"
    If B Then
        Me.lblIndicador.Caption = "Líneas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
    End If
    'Habilitar las opciones correctas del menu segun Modo
'    PonerModoOpcionesMenu (Modo)
'    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
    DataGrid2.Enabled = Not B
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim B As Boolean
Dim Opcion As Byte
Dim Sql As String

    On Error GoTo ECargaGrid

'    b = DataGrid1.Enabled

    If vDataGrid.Name = "DataGrid1" Then
        Opcion = 1
    Else
        Opcion = 2
    End If
    
    Sql = MontaSQLCarga(enlaza, Opcion)
    CargaGridGnral vDataGrid, vData, Sql, PrimeraVez
    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
    
    B = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
    vDataGrid.Enabled = Not B
    PrimeraVez = False
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim tots As String

    On Error GoTo ECargaGrid
    
    vData.Refresh
    Select Case vDataGrid.Name
        Case "DataGrid1" 'Lineas de Albaran
            'SQL = "SELECT codtipom, numfactu, fecfactu, numparte, numlinea,
            'codalmac, codartic, nomartic, ampliaci, cantidad, preciove, dosishab, importel "
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux(0)|T|Linea|620|;S|txtAux(1)|T|Código|950|;S|txtAux(2)|T|Nombre Variedad|4450|;"
            tots = tots & "N||||0|;S|txtAux(3)|T|Cantidad|1050|;S|txtAux(4)|T|Precio|1300|;S|txtAux(5)|T|Dto|1300|;S|txtAux(6)|T|Importe|1650|;" 'N||||0|;"
            arregla tots, DataGrid1, Me
            DataGrid1.Columns(9).Alignment = dbgRight
            DataGrid1.Columns(10).Alignment = dbgRight
'            DataGrid1.Columns(12).Alignment = dbgRight
'            DataGrid1.Columns(13).Alignment = dbgRight
                       
         Case "DataGrid2" 'albaranes x articulo
            'SQL = "SELECT codtipom,numfactu,fecfactu,codtipoa,numalbar, fechaalb,"
            'numpedcl,fecpedcl,sementre,numofert,fecofert, referenc, codenvio,codtraba, codtrab1, codtrab2,observa1,observa2,observa3,observa4,observa5  "
            tots = "N||||0|;N||||0|;N||||0|;"
            tots = tots & "S|txtAux3(0)|T|Albarán|1200|;S|txtAux3(1)|T|Fecha|1450|;"
'            tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;" 'N||||0|;N||||0|;"
            tots = tots & "N||||0|;"
            arregla tots, DataGrid2, Me
        
            DataGrid2_RowColChange 1, 1
    End Select
    
    vDataGrid.HoldFields
    Exit Sub
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux_LostFocus(Index As Integer)

    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    Select Case Index
        Case 4 'Precio
            If txtAux(Index).Text <> "" Then
                PonerFormatoDecimal txtAux(Index), 2 'Tipo 2: Decimal(10,4)
            End If
            
        Case 5, 6 'Descuentos
            PonerFormatoDecimal txtAux(Index), 4 'Tipo 4: Decimal(4,2)
            If Index = 6 Then PonerFoco Me.Text2(16)
            
        Case 7 'Importe Linea
            PonerFormatoDecimal txtAux(Index), 1 'Tipo 3: Decimal(12,2)
    End Select
    
    If (Index = 3 Or Index = 4 Or Index = 6 Or Index = 7) Then 'Cant., Precio, Dto1, Dto2
        If txtAux(1).Text = "" Then Exit Sub
'        txtAux(7).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(5).Text, txtAux(6).Text, vParamAplic.TipoDtos)
        txtAux(7).Text = CalcularImporte(txtAux(3).Text, txtAux(4).Text, txtAux(5).Text, txtAux(6).Text, 0, 0)
        PonerFormatoDecimal txtAux(7), 1
    End If
End Sub


Private Sub BotonMtoLineas(numTab As Integer, cad As String)
    Me.SSTab1.Tab = numTab
    
    If Me.DataGrid1.visible Then 'Lineas de Albaranes
        If Me.Data2.Recordset.RecordCount < 1 Then
            MsgBox "La factura no tiene lineas.", vbInformation
            Exit Sub
        End If
        TituloLinea = cad
        
        ModificaLineas = 0
        PonerModo 5
        PonerBotonCabecera True
    End If
End Sub


Private Function Eliminar() As Boolean
Dim Sql As String
Dim cta As String
Dim B As Boolean
Dim vTipoMov As CTiposMov

    On Error GoTo FinEliminar

        B = False
        Eliminar = False
        If Data1.Recordset.EOF Then Exit Function
        
        conn.BeginTrans
        
        B = True
        
        'Eliminar en tablas de factura de Retirada de Bodega / Almazara: rbodfacturas, rbodfacturas_alb, rbodfacturas_lineas
        '---------------------------------------------------------------
        If B Then
            Sql = " " & ObtenerWhereCP(True)
        
            'Lineas de facturas (slifpc)
            conn.Execute "Delete from " & NomTablaLineas & Sql
        
            'Lineas de cabeceras de albaranes de la factura
            conn.Execute "Delete from rbodfacturas_alb " & Sql
            
            'Cabecera de facturas (scafpc)
            conn.Execute "Delete from " & NombreTabla & Sql
        End If
        
        'Eliminar los movimientos generados por el albaran que genero la factura
        '-----------------------------------------------------------------------
        If B Then
            'Decrementar contador si borramos el ultima factura
            Set vTipoMov = New CTiposMov
            vTipoMov.DevolverContador Text1(17).Text, Val(Text1(0).Text)
            Set vTipoMov = Nothing
        End If
        
        
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Factura", Err.Description
        B = False
    End If
    If Not B Then
        conn.RollbackTrans
'        ConnConta.RollbackTrans
    Else
        conn.CommitTrans
'        ConnConta.CommitTrans
    End If
    Eliminar = B
End Function


Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next
    CargaGrid DataGrid2, Data3, False
    CargaGrid DataGrid1, Data2, False
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         If SituarDataMULTI(Data1, vWhere, Indicador) Then
             If Modo <> 5 Then
                PonerModo 2
                PonerCampos
             End If
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
    Sql = "codtipom= '" & Text1(17).Text & "' and numfactu= " & Text1(0).Text & " and fecfactu='" & Format(Text1(1).Text, FormatoFecha) & "' "
    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
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
    
    If Opcion = 1 Then
        Sql = "SELECT codtipom, numfactu, fecfactu, numalbar, numlinea, rbodfacturas_lineas.codvarie, nomvarie, ampliaci, cantidad, rbodfacturas_lineas.precioar, dtolinea, importel "
        Sql = Sql & " FROM rbodfacturas_lineas inner join variedades on rbodfacturas_lineas.codvarie = variedades.codvarie " 'lineas de factura
    ElseIf Opcion = 2 Then
        Sql = "SELECT codtipom,numfactu,fecfactu,numalbar, fechaalb, observac  "
        Sql = Sql & " FROM rbodfacturas_alb " 'cabeceras albaranes de la factura
    End If
    
    If enlaza Then
        Sql = Sql & " " & ObtenerWhereCP(True)
        'lineas factura proveedor
        If Opcion = 1 Then Sql = Sql & " AND numalbar=" & Data3.Recordset.Fields!numalbar
    Else
        Sql = Sql & " WHERE numfactu = -1"
    End If
    Sql = Sql & " ORDER BY codtipom, numfactu, fecfactu,numalbar "
    If Opcion = 1 Then Sql = Sql & ", numlinea "
    MontaSQLCarga = Sql
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim B As Boolean

        B = ((Modo = 2) Or (Modo = 5 And ModificaLineas = 0)) And Me.Check1(0).Value = 0 And (Factura = "")
        
        'Modificar
        Toolbar1.Buttons(2).Enabled = B
        Me.mnModificar.Enabled = B
        'eliminar
        Toolbar1.Buttons(3).Enabled = B '(Modo = 2)
        Me.mnEliminar.Enabled = B '(Modo = 2)
            
'        b = (Modo = 2)
'        'Mantenimiento lineas
'        Toolbar1.Buttons(9).Enabled = b
'        Me.mnLineas.Enabled = b
        'Imprimir
'        Toolbar1.Buttons(10).Enabled = b
'        Me.mnImprimir.Enabled = b
        
        B = ((Modo >= 3) Or Modo = 1)
        'Buscar
        Toolbar1.Buttons(5).Enabled = Not B And (Factura = "")
        Me.mnBuscar.Enabled = Not B And (Factura = "")
        'Ver Todos
        Toolbar1.Buttons(6).Enabled = Not B And (Factura = "")
        Me.mnVerTodos.Enabled = Not B And (Factura = "")
End Sub


Private Sub PonerDatosSocio(Codsocio As String, Optional nifSocio As String)
Dim vSocio As cSocio
Dim Observaciones As String
    
    On Error GoTo EPonerDatos
    
    If Codsocio = "" Then
        LimpiarDatosSocio
        Exit Sub
    End If

    Set vSocio = New cSocio
    'si se ha modificado el proveedor volver a cargar los datos
    If vSocio.Existe(Codsocio) Then
        If vSocio.LeerDatos(Codsocio) Then
           
            EsDeVarios = False 'vProve.DeVarios
            BloquearDatosSocio (EsDeVarios)
        
            If Modo = 4 And EsDeVarios Then 'Modificar
                'si no se ha modificado el proveedor no hacer nada
                If CLng(Text1(2).Text) = CLng(Data1.Recordset!Codsocio) Then
                    Set vSocio = Nothing
                    Exit Sub
                End If
            End If
        
            Text1(2).Text = vSocio.Codigo
            FormateaCampo Text1(2)
            
            If (Modo = 3) Or (Modo = 4) Then
                Text1(3).Text = vSocio.Nombre  'Nom socio
                Text1(6).Text = vSocio.Direccion
                Text1(7).Text = vSocio.CPostal
                Text1(8).Text = vSocio.Poblacion
                Text1(9).Text = vSocio.Provincia
                Text1(4).Text = vSocio.nif
                Text1(5).Text = DBLet(vSocio.Tfno1, "T")
            End If
            
            Observaciones = DBLet(vSocio.Observaciones)
            If Observaciones <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del socio"
            End If
        End If
    Else
        LimpiarDatosSocio
    End If
    Set vSocio = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Socio", Err.Description
End Sub

'--monica
'Private Sub PonerDatosProveVario(nifProve As String)
''Poner el los campos Text el valor del proveedor
'Dim vProve As CProveedor
'Dim b As Boolean
'
'    If nifProve = "" Then Exit Sub
'
'    Set vProve = New CProveedor
'    b = vProve.LeerDatosProveVario(nifProve)
'    If b Then
'        Text1(3).Text = vProve.Nombre   'Nom proveedor
'        Text1(6).Text = vProve.Domicilio
'        Text1(7).Text = vProve.CPostal
'        Text1(8).Text = vProve.Poblacion
'        Text1(9).Text = vProve.Provincia
'        Text1(5).Text = DBLet(vProve.TfnoAdmon, "T")
'    End If
'    Set vProve = Nothing
'End Sub
'

Private Sub LimpiarDatosSocio()
Dim I As Byte

    For I = 3 To 9
        Text1(I).Text = ""
    Next I
End Sub
   

Private Function ModificaAlbxFac() As Boolean
Dim Sql As String
Dim B As Boolean
On Error GoTo EModificaAlb
    
    ModificaAlbxFac = False
    If Data3.Recordset.EOF Then Exit Function
    
    'comprobar datos OK de la scafac1
     B = CompForm(Me) 'Comprobar formato datos ok de la cabecera alb: opcion=2
    If Not B Then Exit Function
'--monica
'    SQL = "UPDATE scafpa SET codtrab2=" & DBSet(Text3(0).Text, "N", "S") & ", "
'    SQL = SQL & "codtrab1=" & DBSet(Text3(1).Text, "N", "S")
    If Me.FrameObserva.visible Then
        Sql = "UPDATE rbodfacturas_alb SET "
        Sql = Sql & " observac=" & DBSet(Text3(4).Text, "T")
'        SQL = SQL & ", observa2=" & DBSet(Text3(5).Text, "T")
'        SQL = SQL & ", observa3=" & DBSet(Text3(6).Text, "T")
'        SQL = SQL & ", observa4=" & DBSet(Text3(7).Text, "T")
'        SQL = SQL & ", observa5=" & DBSet(Text3(8).Text, "T")
        Sql = Sql & ObtenerWhereCP(True)
        Sql = Sql & " AND numalbar=" & Data3.Recordset.Fields!numalbar
        conn.Execute Sql
    End If
'--monica
'    SQL = SQL & ObtenerWhereCP(True)
'    SQL = SQL & " AND numalbar=" & Data3.Recordset.Fields!numalbar
'    Conn.Execute SQL
    ModificaAlbxFac = True
    
EModificaAlb:
If Err.Number <> 0 Then MuestraError Err.Number, "Modificar Albaranes de factura", Err.Description
End Function



Private Function ModificarFactura(Optional sqlLineas As String) As Boolean
'si se ha modificado la linea de slifac, añadir a la transaccion la modificación de la linea y recalcular
Dim bol As Boolean
Dim MenError As String
Dim Sql As String
Dim vFactuBOD As CFacturaBOD
On Error GoTo EModFact

    bol = False
    conn.BeginTrans
    
    
'    '++monica:añadida la condicion de solo si hay contabilidad
'    If vParamAplic.NumeroConta <> 0 Then ConnConta.BeginTrans

    If sqlLineas <> "" Then
        'actualizar el importe de la linea modificada
        MenError = "Modificando lineas de Factura."
        conn.Execute sqlLineas
    End If

'    'recalcular las bases imponibles x IVA
'    MenError = "Recalcular importes IVA"
'    bol = ActualizarDatosFactura
    bol = True
    
    If bol Then
        'modificamos la scafpc
        MenError = "Modificando cabecera de factura"
        bol = ModificaDesdeFormulario1(Me, 1)

'        If bol Then
'            'Si es proveedor de varios actualizar datos proveedor en tabla:sprvar
'            MenError = "Modificando datos socio varios"
'            bol = ActualizarProveVarios(Text1(2).Text, Text1(4).Text)
'        End If

        If bol Then
            MenError = "Modificando albaranes de factura"
            'modificar la tabla: scafpa
            bol = ModificaAlbxFac
'            '++monica:añadida la condicion de solo si hay contabilidad
'            If vParamAplic.NumeroConta <> 0 Then
'                If bol Then 'si se ha modificado la factura
'                    MenError = "Actualizando en Tesoreria"
'                    'y eliminar de tesoreria conta.spagop los registros de la factura
'
'                    'antes de Eliminar en las tablas de la Contabilidad
'                    Set vFactu = New CFacturaADV
'                    bol = vFactu.LeerDatos(Text1(2).Text, Text1(0).Text, Text1(1).Text)
'
'                    If bol Then
'                        'Eliminar de la spagop
'                        SQL = " ctaprove='" & vFactu.CtaProve & "' AND numfactu='" & Data1.Recordset.Fields!numfactu & "'"
'                        SQL = SQL & " AND fecfactu='" & Format(Data1.Recordset.Fields!fecfactu, FormatoFecha) & "'"
'                        ConnConta.Execute "Delete from spagop WHERE " & SQL
'
'                        'Volvemos a grabar en TESORERIA. Tabla de Contabilidad: sconta.spagop
'                        If bol Then
'                            bol = vFactu.InsertarEnTesoreria(MenError)
'                        End If
'                    End If
'                    Set vFactu = Nothing
'                End If
'            End If
        End If
    End If

EModFact:
     If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        '++monica:añadida la condicion de solo si hay contabilidad
'        If vParamAplic.NumeroConta <> 0 Then ConnConta.CommitTrans
        ModificarFactura = True
    Else
        conn.RollbackTrans
        '++monica:añadida la condicion de solo si hay contabilidad
'        If vParamAplic.NumeroConta <> 0 Then ConnConta.RollbackTrans
        ModificarFactura = False
        MenError = "Actualizando Factura." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
    End If
End Function



Private Function FactContabilizada() As Boolean
Dim cta As String, numasien As String
On Error GoTo EContab

    'comprabar que se puede modificar/eliminar la factura
    If Me.Check1(0).Value = 1 Then 'si esta contabilizada
    
'--monica: he sustituido lo de abajo por
'        'comprobar en la contabilidad si esta contabilizada
'        cta = DevuelveDesdeBDNew(cAgro, "rsocios_seccion", "codmaccli", "codsocio", Text1(2).Text, "N", , "codsecci", vParamAplic.SeccionADV, "N")
'        If cta <> "" Then
'            numasien = DevuelveDesdeBDNew(cConta, "cabfact", "numasien", "codmacta", cta, "T", , "codfaccl", Text1(0).Text, "T", "fecfaccl", Text1(1).Text, "F")
'            If numasien <> "" Then
'                FactContabilizada = True
'                MsgBox "La factura esta contabilizada y no se puede modificar ni eliminar.", vbInformation
'                Exit Function
'            Else
'                FactContabilizada = False
'            End If
'        Else
'            FactContabilizada = True
'            Exit Function
'        End If

'++monica:
        FactContabilizada = True
    Else
        FactContabilizada = False
    End If
EContab:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Factura contabilizada", Err.Description
End Function


Private Sub TxtAux3_GotFocus(Index As Integer)
    ConseguirFoco txtAux3(Index), Modo
End Sub

Private Sub TxtAux3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index <> 0 And KeyCode <> 38 Then KEYdown KeyCode
End Sub

Private Sub TxtAux3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub TxtAux3_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux3(Index), Modo) Then Exit Sub
End Sub


Private Sub BloquearDatosSocio(bol As Boolean)
Dim I As Byte

    'bloquear/desbloquear campos de datos segun sea de varios o no
    If Modo <> 5 Then
        Me.imgBuscar(1).visible = bol 'NIF
        Me.imgBuscar(1).Enabled = bol 'NIF
        Me.imgBuscar(2).Enabled = bol 'poblacion
        
        For I = 3 To 9 'si no es de varios no se pueden modificar los datos
            BloquearTxt Text1(I), Not bol
        Next I
    End If
End Sub


Private Function ActualizarProveVarios(Prove As String, nif As String) As Boolean
''Modifica los datos de la tabla de Proveedores Varios
'Dim vProve As CProveedor
'On Error GoTo EActualizarCV
'
'    ActualizarProveVarios = False
'
'    Set vProve = New CProveedor
'    If EsProveedorVarios(Prove) Then
'        vProve.NIF = NIF
'        vProve.Nombre = Text1(3).Text
'        vProve.Domicilio = Text1(6).Text
'        vProve.CPostal = Text1(7).Text
'        vProve.Poblacion = Text1(8).Text
'        vProve.Provincia = Text1(9).Text
'        vProve.TfnoAdmon = Text1(5).Text
'        vProve.ActualizarProveV (NIF)
'    End If
'    Set vProve = Nothing
'
'    ActualizarProveVarios = True
'
'EActualizarCV:
'    If Err.Number <> 0 Then
'        ActualizarProveVarios = False
'    Else
'        ActualizarProveVarios = True
'    End If
End Function


Private Function ObtenerSelFactura() As String
'Cuando venimos desde dobleClick en Movimientos de Articulos para Albaranes ya
'Facturados, abrimos este form pero cargando los datos de la factura
'correspendiente al albaran que se selecciono
Dim cad As String
Dim Rs As ADODB.Recordset
On Error Resume Next

    cad = "SELECT codsocio,numfactu,fecfactu FROM rbodfacturas_alb "
    cad = cad & " WHERE codsocio=" & DBSet(hcoCodSocio, "N") & " AND numalbar=" & DBSet(hcoCodMovim, "T")
    cad = cad & " AND fechaalb=" & DBSet(hcoFechaMovim, "F")

    Set Rs = New ADODB.Recordset
    Rs.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then 'where para la factura
        cad = " WHERE codsocio=" & Rs!Codsocio & " AND numfactu= " & Rs!numfactu & " AND fecfactu=" & DBSet(Rs!fecfactu, "F")
    Else
        cad = " where numfactu=-1"
    End If
    Rs.Close
    Set Rs = Nothing

    ObtenerSelFactura = cad
End Function



Private Function ActualizarDatosFactura() As Boolean
Dim vFactuBOD As CFacturaBOD
Dim cadSel As String
Dim vSocio As cSocio

'    Set vFactuADV = New CFacturaADV
'    cadSel = ObtenerWhereCP(False)
'    cadSel = "advfacturas_lineas." & cadSel
'    vFactuADV.DtoPPago = CCur(Text1(11).Text)
'    vFactuADV.DtoGnral = CCur(Text1(12).Text)
'
'    Set vSocio = New CSocio
'    If EsSocioDeSeccion(Text1(2).Text, vParamAplic.SeccionADV) Then
'        If vFactuADV.CalcularDatosFacturaADV(vSocio) Then
'            Text1(14).Text = vFactuADV.BrutoFac
'            Text1(15).Text = vFactuADV.ImpPPago
'            Text1(16).Text = vFactuADV.ImpGnral
'            Text1(18).Text = vFactuADV.TipoIVA1
'            Text1(19).Text = vFactuADV.TipoIVA2
'            Text1(20).Text = vFactuADV.TipoIVA3
'            Text1(21).Text = vFactuADV.PorceIVA1
'            Text1(22).Text = vFactuADV.PorceIVA2
'            Text1(23).Text = vFactuADV.PorceIVA3
'            Text1(24).Text = vFactuADV.BaseIVA1
'            Text1(25).Text = vFactuADV.BaseIVA2
'            Text1(26).Text = vFactuADV.BaseIVA3
'            Text1(27).Text = vFactuADV.ImpIVA1
'            Text1(28).Text = vFactuADV.ImpIVA2
'            Text1(29).Text = vFactuADV.ImpIVA3
'            Text1(30).Text = vFactuADV.TotalFac
'
'            FormatoDatosTotales
'
'            ActualizarDatosFactura = True
'        Else
'            ActualizarDatosFactura = False
'            MuestraError Err.Number, "Recalculando Factura", Err.Description
'        End If
'        Set vFactuADV = Nothing
'    End If
'
'    Set vSocio = Nothing
'
End Function


Private Sub FormatoDatosTotales()
Dim I As Byte

    For I = 14 To 16
'        Text1(I).Text = QuitarCero(Text1(I).Text)
        FormateaCampo Text1(I)
    Next I
    
    For I = 24 To 26
        If Text1(I).Text <> "" Then
            'Si la Base Imp. es 0
            If CSng(Text1(I).Text) = 0 Then
                Text1(I).Text = QuitarCero(Text1(I).Text)
                Text1(I - 3).Text = QuitarCero(Text1(I - 3).Text)
                Text1(I - 6).Text = QuitarCero(Text1(I - 6).Text)
                Text1(I + 3).Text = QuitarCero(Text1(I + 3).Text)
            Else
                FormateaCampo Text1(I)
                FormateaCampo Text1(I - 3)
                FormateaCampo Text1(I - 6)
                FormateaCampo Text1(I + 3)
            End If
        Else 'No hay Base Imponible
            Text1(I - 3).Text = QuitarCero(Text1(I - 3).Text)
            Text1(I - 6).Text = QuitarCero(Text1(I - 6).Text)
            Text1(I + 3).Text = ""
        End If
    Next I
End Sub



Private Sub ComprobarDatosTotales()
Dim I As Byte

    For I = 14 To 16
        Text1(I).Text = ComprobarCero(Text1(I).Text)
    Next I
End Sub


Private Sub AbrirFrmForpaConta(Indice As Integer)
    Set frmFPago = New frmForpaConta
    frmFPago.DatosADevolverBusqueda = "0|1|"
    frmFPago.CodigoActual = Text1(Indice)
    frmFPago.Show vbModal
    Set frmFPago = Nothing
End Sub


Private Sub ConexionConta(Seccion As String)
    
    If vSeccion Is Nothing Then
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(Seccion) Then
            vSeccion.AbrirConta
        End If
    Else
        ' si el objeto existia: cerramos la conexion y volvemos crearlo
        vSeccion.CerrarConta
        Set vSeccion = Nothing
        
        
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(Seccion) Then
            vSeccion.AbrirConta
        End If
    End If
End Sub
