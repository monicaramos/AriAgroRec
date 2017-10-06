VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmADVListArticulos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   9435
   Icon            =   "frmADVListArticulos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   8730
      Top             =   5580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameInfArticulos 
      Height          =   5880
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9435
      Begin VB.ComboBox combo1 
         BeginProperty Font 
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
         Left            =   1740
         TabIndex        =   31
         Tag             =   "Cod. Tipo Artículo|T|N|||advartic|tipoprod||N|"
         Text            =   "Combo2"
         Top             =   4785
         Width           =   1845
      End
      Begin VB.Frame FrameOrden 
         BorderStyle     =   0  'None
         Height          =   2025
         Left            =   6480
         TabIndex        =   25
         Top             =   1305
         Width           =   2655
         Begin VB.CommandButton cmdBajar 
            Height          =   510
            Left            =   2055
            Picture         =   "frmADVListArticulos.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   1080
            Width           =   510
         End
         Begin VB.CommandButton cmdSubir 
            Caption         =   " "
            Height          =   510
            Left            =   2055
            Picture         =   "frmADVListArticulos.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   375
            Width           =   510
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   1335
            Left            =   255
            TabIndex        =   28
            Top             =   255
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   2355
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Orden del Informe"
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
            Index           =   31
            Left            =   255
            TabIndex        =   29
            Top             =   15
            Width           =   1770
         End
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   63
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   1845
         Width           =   4635
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   62
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text5"
         Top             =   1440
         Width           =   4635
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   71
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   4035
         Width           =   4800
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   70
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   3630
         Width           =   4800
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   63
         Left            =   1755
         MaxLength       =   4
         TabIndex        =   4
         Top             =   1845
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   62
         Left            =   1755
         MaxLength       =   4
         TabIndex        =   3
         Top             =   1440
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   71
         Left            =   1755
         MaxLength       =   16
         TabIndex        =   8
         Top             =   4035
         Width           =   1815
      End
      Begin VB.TextBox txtCodigo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   70
         Left            =   1755
         MaxLength       =   16
         TabIndex        =   7
         Top             =   3630
         Width           =   1815
      End
      Begin VB.CommandButton cmdAceptarArtic 
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
         Left            =   6435
         TabIndex        =   9
         Top             =   5130
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   11
         Left            =   7515
         TabIndex        =   10
         Top             =   5115
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   66
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   5
         Top             =   2520
         Width           =   1005
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   67
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   6
         Top             =   2925
         Width           =   1005
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   66
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "Text5"
         Top             =   2520
         Width           =   4575
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   67
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Text5"
         Top             =   2925
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   "Informe de Artículos"
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
         Left            =   495
         TabIndex        =   30
         Top             =   495
         Width           =   6735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Artículo"
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
         Index           =   30
         Left            =   570
         TabIndex        =   24
         Top             =   4500
         Width           =   1560
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   20
         Left            =   1470
         ToolTipText     =   "Buscar familia"
         Top             =   1845
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   19
         Left            =   1470
         ToolTipText     =   "Buscar familia"
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Familia"
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
         Index           =   39
         Left            =   555
         TabIndex        =   23
         Top             =   1110
         Width           =   660
      End
      Begin VB.Label Label3 
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
         Index           =   56
         Left            =   780
         TabIndex        =   22
         Top             =   1845
         Width           =   600
      End
      Begin VB.Label Label3 
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
         Index           =   55
         Left            =   780
         TabIndex        =   21
         Top             =   1440
         Width           =   645
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   28
         Left            =   1470
         ToolTipText     =   "Buscar artículo"
         Top             =   4035
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   27
         Left            =   1470
         ToolTipText     =   "Buscar artículo"
         Top             =   3630
         Width           =   240
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
         Index           =   38
         Left            =   555
         TabIndex        =   20
         Top             =   3270
         Width           =   750
      End
      Begin VB.Label Label3 
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
         Index           =   54
         Left            =   780
         TabIndex        =   19
         Top             =   4035
         Width           =   600
      End
      Begin VB.Label Label3 
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
         Index           =   51
         Left            =   780
         TabIndex        =   18
         Top             =   3630
         Width           =   645
      End
      Begin VB.Label Label3 
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
         Index           =   50
         Left            =   780
         TabIndex        =   17
         Top             =   2520
         Width           =   645
      End
      Begin VB.Label Label3 
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
         Index           =   48
         Left            =   780
         TabIndex        =   16
         Top             =   2925
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
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
         Index           =   37
         Left            =   555
         TabIndex        =   15
         Top             =   2145
         Width           =   990
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   23
         Left            =   1470
         ToolTipText     =   "Buscar proveedor"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   24
         Left            =   1470
         ToolTipText     =   "Buscar proveedor"
         Top             =   2955
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmADVListArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmArt As frmADVArticulos
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents frmFam As frmComercial
Attribute frmFam.VB_VarHelpID = -1
Private WithEvents frmPro As frmComercial
Attribute frmPro.VB_VarHelpID = -1

Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1

'---- Variables para el INFORME ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para el frmImprimir
Private conSubRPT As Boolean 'Si el informe tiene subreports
Private cadNombreRPT As String 'Nombre del informe
'-----------------------------------

Dim TipCod As String
Dim indCodigo As Integer 'indice para txtCodigo

Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean
Dim indFrame As Single


Private Sub KEYpress(KeyAscii As Integer)
    Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptarArtic_Click()
'Listado de Articulos
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim campo As String
Dim Opcion As Byte, numOp As Byte

    InicializarVbles
    
    cadNombreRPT = "rADVListArticulos.rpt"  'Nombre fichero .rpt a Imprimir
    cadTabla = "advartic"
    
    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H FAMILIA
    '--------------------------------------------
    cDesde = Trim(txtCodigo(62).Text)
    cHasta = Trim(txtCodigo(63).Text)
    nDesde = txtNombre(62).Text
    nHasta = txtNombre(63).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & cadTabla & ".codfamia}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFamilia= """) Then Exit Sub
    End If

    
    'Cadena para seleccion D/H PROVEEDOR
    '--------------------------------------------
    cDesde = Trim(txtCodigo(66).Text)
    cHasta = Trim(txtCodigo(67).Text)
    nDesde = txtNombre(66).Text
    nHasta = txtNombre(67).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & cadTabla & ".codprove}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHProveedor= """) Then Exit Sub
    End If
    
    'Cadena para seleccion TIPO ARTICULO
    '--------------------------------------------
    If Combo1(0).ListIndex <> 3 Then
        If AnyadirAFormula(cadFormula, "{advartic.tipoprod}=" & Combo1(0).ListIndex) = False Then Exit Sub
        If AnyadirAFormula(cadSelect, "{advartic.tipoprod}=" & Combo1(0).ListIndex) = False Then Exit Sub
    End If
    
    'Cadena para seleccion D/H ARTICULO
    '--------------------------------------------
    cDesde = Trim(txtCodigo(70).Text)
    cHasta = Trim(txtCodigo(71).Text)
    nDesde = txtNombre(70).Text
    nHasta = txtNombre(71).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & cadTabla & ".codartic}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHArticulo= """) Then Exit Sub
    End If

    
    'Obtener el parametro con el Orden del Informe
    '---------------------------------------------
        numOp = PonerGrupo(1, ListView2.ListItems(1).Text)
        If numOp <> 0 Then Opcion = numOp
        numOp = PonerGrupo(2, ListView2.ListItems(2).Text)
        If numOp <> 0 Then Opcion = numOp
        numOp = PonerGrupo(3, ListView2.ListItems(3).Text)
        If numOp <> 0 Then Opcion = numOp
        Opcion = Opcion - 1
    
        Select Case Opcion
            Case 1 'El group2 es el Proveedor
                campo = "pTitulo1=""" & ListView2.ListItems(3).Text & """" '3
                CadParam = CadParam & campo & "|"
                numParam = numParam + 1
                
                campo = "pTitulo2=""" & ListView2.ListItems(3).Text & """" '4
                CadParam = CadParam & campo & "|"
                numParam = numParam + 1
            Case 2 'El Group3 es el Proveedor
                campo = "pTitulo1=""" & ListView2.ListItems(2).Text & """" '2
                CadParam = CadParam & campo & "|"
                numParam = numParam + 1
                
                campo = "pTitulo2=""" & ListView2.ListItems(3).Text & """" '4
                CadParam = CadParam & campo & "|"
                numParam = numParam + 1
            Case 3, 0 'El Group4 es el Proveedor
                      '0 'El Group1 es el Proveedor
                campo = "pTitulo1=""" & ListView2.ListItems(2).Text & """" '2
                CadParam = CadParam & campo & "|"
                numParam = numParam + 1
                
                campo = "pTitulo2=""" & ListView2.ListItems(3).Text & """" '3
                CadParam = CadParam & campo & "|"
                numParam = numParam + 1
                
                If Opcion = 0 Then
                    campo = "pTitulo3=""" & ListView2.ListItems(3).Text & """" '4
                    CadParam = CadParam & campo & "|"
                    numParam = numParam + 1
                End If
        End Select
   
    'Parametro Orden del Informe
    campo = "pOrden=" & Opcion
    CadParam = CadParam & campo & "|"
    numParam = numParam + 1
    
    If HayRegParaInforme(cadTabla, cadSelect) Then
       LlamarImprimir
    End If
    
End Sub

Private Sub cmdBajar_Click()
'Bajar el item seleccionado del listview2
    BajarItemList Me.ListView2
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdSubir_Click()
'Subir el item seleccionado del listview2 una posicion
    SubirItemList Me.ListView2
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(62)
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim I As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon


    PrimeraVez = True
    limpiar Me

    For I = 19 To 20
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    For I = 23 To 24
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    For I = 27 To 28
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I

    'Ocultar todos los Frames de Formulario
    Me.FrameInfArticulos.visible = False
    
    CommitConexion
    
    cadTitulo = ""
    cadNombreRPT = ""
    
    ListadosAlmacen H, W
    
    CargaCombo
    
    Combo1(0).ListIndex = 3
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub



Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Articulos
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmFam_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Familias
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMarcas_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Marcas
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmPro_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Proveedores
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTArt_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Tipo de articulos
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
'Buscar general: cada index llama a una tabla
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 19, 20 'cod. FAMILIA
            indCodigo = Index + 43
            Set frmFam = New frmComercial
            
            AyudaFamiliasADV frmFam, txtCodigo(indCodigo).Text
            
            Set frmFam = Nothing
            
  
        Case 27, 28 'cod. ARTICULO
            indCodigo = Index + 43
            Set frmArt = New frmADVArticulos
            frmArt.DatosADevolverBusqueda = "0|1|" 'Abrimos en Modo Busqueda
            frmArt.Show vbModal
            Set frmArt = Nothing
            

        Case 23, 24 'cod. PROVEEDOR
            indCodigo = Index + 43
            
            Set frmPro = New frmComercial
            
            AyudaProveedoresCom frmPro, txtCodigo(indCodigo).Text
            
            Set frmPro = Nothing
            
    End Select
    PonerFoco txtCodigo(indCodigo)
    Screen.MousePointer = vbDefault
End Sub



Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim tabla As String
Dim codcampo As String, nomCampo As String
Dim TipCampo As String, Formato As String
Dim Titulo As String
Dim EsNomCod As Boolean 'Si es campo Cod-Descripcion llama a PonerNombreDeCod


    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    EsNomCod = False
        
    Select Case Index
        Case 70, 71  'Cod. ARTICULO
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "advartic", "nomartic", "codartic", "T")
        
        Case 62, 63 'Cod. FAMILIA
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "advfamia", "nomfamia", "codfamia", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
        
        Case 66, 67 'PROVEEDOR
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "proveedor", "nomprove")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
    End Select
    
End Sub

Private Sub ponerFrameArticulosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el informe de Articulos, de tabla: sartic
Dim B As Boolean

    B = True
    H = 6620
    W = 9435
    
    PonerFrameVisible Me.FrameInfArticulos, visible, H, W

    If visible = True Then
        'visible orden campos si opcionlistado=6
        Me.FrameOrden.visible = B
    End If
End Sub


Private Sub CargarListViewOrden()
'Carga el List View del frame: frameInfArticulos
'para establecer el orden en que se van a mostrar los datos en el Informe
'Orden: Familia, MArca, Proveedor, Tipo de Articulo, Articulo
Dim ItmX As ListItem

    'Los encabezados
    ListView2.ColumnHeaders.Clear
    ListView2.ColumnHeaders.Add , , "Campo", 1600
    Set ItmX = ListView2.ListItems.Add
    ItmX.Text = "Familia"
    Set ItmX = ListView2.ListItems.Add
    ItmX.Text = "Proveedor"
    Set ItmX = ListView2.ListItems.Add
    ItmX.Text = "Tipo Articulo"
End Sub


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    CadParam = ""
    numParam = 0
    conSubRPT = False
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
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        devuelve2 = CadenaDesdeHastaBD(codD, codH, Codigo, TipCod)
        If devuelve2 = "Error" Then Exit Function
        If Not AnyadirAFormula(cadSelect, devuelve2) Then Exit Function
    End If
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            CadParam = CadParam & AnyadirParametroDH(param, codD, codH, nomD, nomH)
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function


Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = CadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .Titulo = cadTitulo
        .EnvioEMail = False
        .NombreRPT = cadNombreRPT
        .Opcion = 0 'Opcion
        .Show vbModal
    End With
End Sub


Private Function PonerGrupo(numGrupo As Byte, cadgrupo As String) As Byte
Dim campo As String
Dim nomCampo As String
Dim NomTipo As String

    campo = "pGroup" & numGrupo & "="
    nomCampo = "pGroup" & numGrupo & "Name="
    PonerGrupo = 0
    
    Select Case cadgrupo
        Case "Familia"
            CadParam = CadParam & campo & "{advartic.codfamia}" & "|"
            If numGrupo = 1 Then
                CadParam = CadParam & nomCampo & " ""FAMILIA: "" & " & " totext({advartic.codfamia},""0000"") & " & """  """ & " & {advfamia.nomfamia}" & "|"
            Else
                CadParam = CadParam & nomCampo & " totext({advartic.codfamia},""0000"") & " & """ """ & " & {advfamia.nomfamia}" & "|"
            End If
            numParam = numParam + 1
        Case "Proveedor"
            CadParam = CadParam & campo & "{advartic.codprove}" & "|"
            If numGrupo = 1 Then
                CadParam = CadParam & nomCampo & " ""PROVEEDOR: "" & " & " totext({advartic.codprove},""000000"") & " & """  """ & " & {proveedor.nomprove}" & "|"
            Else
                CadParam = CadParam & nomCampo & " totext({advartic.codprove},""000000"") & " & """ """ & " & {proveedor.nomprove}" & "|"
            End If
            numParam = numParam + 1
            PonerGrupo = numGrupo
        Case "Tipo Articulo"
            CadParam = CadParam & campo & "{advartic.tipoprod}" & "|"
            Select Case Combo1(0).ListIndex
                Case 0
                    NomTipo = "Producto"
                Case 1
                    NomTipo = "Trabajo"
                Case 2
                    NomTipo = "Varios"
                Case 3
                    NomTipo = "Todos"
            End Select
            
            If numGrupo = 1 Then
                CadParam = CadParam & nomCampo & " ""TIPO ARTICULO: "" & {@nomTipo}" & "|"
            Else
                CadParam = CadParam & nomCampo & "{@nomTipo}" & "|"
            End If
            numParam = numParam + 1
    End Select

End Function


Private Sub ListadosAlmacen(H As Integer, W As Integer)
   'Listado de Artículo
    ponerFrameArticulosVisible True, H, W
    CargarListViewOrden
    Codigo = "{advartic"
    indFrame = 11
    cadTitulo = "Listado de Artículos ADV"
End Sub


Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim I As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    Combo1(0).Clear

    Combo1(0).AddItem "Producto"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Trabajo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Varios"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    Combo1(0).AddItem "Todos"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 3

End Sub

