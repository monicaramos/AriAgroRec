VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTercListFact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6840
   Icon            =   "frmTercListFact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCobros 
      Height          =   5925
      Left            =   60
      TabIndex        =   10
      Top             =   60
      Width           =   6555
      Begin VB.CheckBox Check1 
         Caption         =   "Imprimir Resumen"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   510
         TabIndex        =   28
         Top             =   5070
         Width           =   2175
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
         Index           =   7
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Factura|T|S|||clientes|codposta|||"
         Top             =   2610
         Width           =   1050
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
         Index           =   6
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Factura|T|S|||clientes|codposta|||"
         Top             =   2220
         Width           =   1050
      End
      Begin VB.TextBox txtNombre 
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
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Text5"
         Top             =   4575
         Width           =   3360
      End
      Begin VB.TextBox txtNombre 
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
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Text5"
         Top             =   4200
         Width           =   3360
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
         Index           =   5
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   7
         Top             =   4590
         Width           =   830
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
         Index           =   4
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   6
         Top             =   4200
         Width           =   830
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
         Index           =   3
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1605
         Width           =   1350
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
         Index           =   2
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1245
         Width           =   1350
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
         Left            =   4965
         TabIndex        =   9
         Top             =   5265
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
         Left            =   3780
         TabIndex        =   8
         Top             =   5265
         Width           =   1065
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
         Index           =   0
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   4
         Top             =   3180
         Width           =   830
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
         Index           =   1
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   5
         Top             =   3555
         Width           =   830
      End
      Begin VB.TextBox txtNombre 
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
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   3180
         Width           =   3360
      End
      Begin VB.TextBox txtNombre 
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
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   3555
         Width           =   3360
      End
      Begin VB.Label Label4 
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
         Index           =   5
         Left            =   825
         TabIndex        =   27
         Top             =   2640
         Width           =   645
      End
      Begin VB.Label Label4 
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
         Index           =   4
         Left            =   825
         TabIndex        =   26
         Top             =   2280
         Width           =   690
      End
      Begin VB.Label Label4 
         Caption         =   "Número Factura"
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
         Index           =   3
         Left            =   510
         TabIndex        =   25
         Top             =   1935
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Retenciones de Terceros"
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
         Left            =   510
         TabIndex        =   24
         Top             =   360
         Width           =   5070
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1530
         MouseIcon       =   "frmTercListFact.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   4620
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1530
         MouseIcon       =   "frmTercListFact.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   4200
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00972E0B&
         Height          =   240
         Index           =   2
         Left            =   495
         TabIndex        =   23
         Top             =   3915
         Width           =   855
      End
      Begin VB.Label Label4 
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
         Index           =   1
         Left            =   810
         TabIndex        =   22
         Top             =   4620
         Width           =   645
      End
      Begin VB.Label Label4 
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
         Index           =   0
         Left            =   810
         TabIndex        =   21
         Top             =   4200
         Width           =   690
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   510
         TabIndex        =   18
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label Label4 
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
         Index           =   15
         Left            =   825
         TabIndex        =   17
         Top             =   1245
         Width           =   690
      End
      Begin VB.Label Label4 
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
         Index           =   14
         Left            =   825
         TabIndex        =   16
         Top             =   1605
         Width           =   645
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1530
         Picture         =   "frmTercListFact.frx":02B0
         ToolTipText     =   "Buscar fecha"
         Top             =   1245
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1530
         Picture         =   "frmTercListFact.frx":033B
         ToolTipText     =   "Buscar fecha"
         Top             =   1605
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   13
         Left            =   825
         TabIndex        =   15
         Top             =   3180
         Width           =   690
      End
      Begin VB.Label Label4 
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
         Index           =   12
         Left            =   825
         TabIndex        =   14
         Top             =   3555
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00972E0B&
         Height          =   240
         Index           =   11
         Left            =   510
         TabIndex        =   13
         Top             =   2895
         Width           =   540
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1530
         MouseIcon       =   "frmTercListFact.frx":03C6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   3180
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1530
         MouseIcon       =   "frmTercListFact.frx":0518
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   3555
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmTercListFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmVar As frmComVar 'Variedad
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'Socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me

End Sub

Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim I As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim Tipo As Byte

InicializarVbles
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    CadParam = CadParam & "|pUsu=" & vUsu.Codigo & "|"
    numParam = numParam + 1
    
    
     '======== FORMULA  ====================================
    'Seleccionar registros de la empresa conectada
'    Codigo = "{" & tabla & ".codempre}=" & vEmpresa.codEmpre
'    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
'    If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
    
    
    'D/H Socio Tercero
    cDesde = Trim(txtCodigo(0).Text)
    cHasta = Trim(txtCodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rcafter.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H Variedad
    cDesde = Trim(txtCodigo(4).Text)
    cHasta = Trim(txtCodigo(5).Text)
    nDesde = txtNombre(4).Text
    nHasta = txtNombre(5).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rlifter.codvarie}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad= """) Then Exit Sub
    End If
    
    'D/H Fecha factura
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rcafter.fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    'D/H Numer de factura
    cDesde = Trim(txtCodigo(6).Text)
    cHasta = Trim(txtCodigo(7).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rcafter.numfactu}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFactura= """) Then Exit Sub
    End If
        
    Tipo = 0
    If Check1.Value Then Tipo = 1
    
    CadParam = CadParam & "pResumen=" & Tipo & "|"
    numParam = numParam + 1


    cadTabla = "rcafter INNER JOIN rlifter ON rcafter.codsocio = rlifter.codsocio and rcafter.numfactu = rlifter.numfactu and rcafter.fecfactu = rlifter.fecfactu "
    
    '[Monica]20/01/2015: cargamos la tabla temporal con las facturas de todas las campañas
    If CargarFacturas(cadTabla, cadSelect) Then
        If HayRegParaInforme("tmprcafter", "tmprcafter.codusu=" & vUsu.Codigo) Then 'nTabla, cadSelect) Then
              'Nombre fichero .rpt a Imprimir
              cadFormula = "{tmprcafter.codusu} = " & vUsu.Codigo
              cadTitulo = "Listado de Facturas de Terceros"
              cadNombreRPT = "rInfFactTerc.rpt"
              LlamarImprimir
        End If
    End If
End Sub

Private Function CargarFacturas(cTabla As String, cSelect As String) As Boolean
Dim Sql As String
Dim vCampAnt As CCampAnt
Dim ctabla1 As String
Dim Sql2 As String
Dim Sql4 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim CadValues As String
Dim BaseIva As Currency
Dim ImpoIva As Currency
Dim SocioAnt As Long
Dim DifBase As Currency
Dim DifRete As Currency
Dim TipoIRPF As Byte
Dim ImpoReten As Currency
Dim BaseReten As Currency
Dim Producto As Long

Dim BDatos As Dictionary
Dim SqlBd As String
Dim RsBd As ADODB.Recordset

Dim cTablaAnticip As String


    On Error GoTo eCargarFacturas
    
    Screen.MousePointer = vbHourglass
    
    Set BDatos = New Dictionary
    
    ' quitamos las llaves de la tabla y where
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    If cSelect <> "" Then
        cSelect = QuitarCaracterACadena(cSelect, "{")
        cSelect = QuitarCaracterACadena(cSelect, "}")
        cSelect = QuitarCaracterACadena(cSelect, "_1")
    End If
    
    ' borramos las tablas temporales donde insertaremos las facturas para los listados
    Sql = "delete from tmprcafter where codusu= " & vUsu.Codigo
    conn.Execute Sql
    Sql = "delete from tmprlifter where codusu= " & vUsu.Codigo
    conn.Execute Sql
    
    ' insertamos las facturas correspondientes a la campaña actual
    Sql = "insert into tmprcafter (codusu,codsocio,numfactu,fecrecep,nomsocio,domsocio,codpobla,pobsocio,prosocio,nifsocio,"
    Sql = Sql & "telsocio,codforpa,brutofac,dtoppago,dtognral,impppago,impgnral,baseiva1,baseiva2,baseiva3,tipoiva1,tipoiva2,tipoiva3,"
    Sql = Sql & "porciva1,porciva2,porciva3,porcrec1,porcrec2,porcrec3,impoiva1,impoiva2,impoiva3,imporec1,imporec2,imporec3,"
    Sql = Sql & "retfacpr,basereten,trefacpr,totalfac,intconta,intracom,esanticipo,porccorredor,concepcargo,impcargo) "
    Sql = Sql & "select " & vUsu.Codigo & ",rcafter.codsocio,rcafter.numfactu,rcafter.fecrecep,rcafter.nomsocio,rcafter.domsocio,rcafter.codpobla,rcafter.pobsocio,rcafter.prosocio,rcafter.nifsocio,rcafter.telsocio,rcafter.codforpa, "
    Sql = Sql & "brutofac,dtoppago,dtognral,impppago,impgnral,baseiva1,baseiva2,baseiva3,tipoiva1,tipoiva2,tipoiva3,porciva1,porciva2,porciva3,"
    Sql = Sql & "porcrec1,porcrec2,porcrec3,impoiva1,impoiva2,impoiva3,imporec1,imporec2,imporec3,retfacpr,basereten,trefacpr,totalfac,"
    Sql = Sql & "intconta,intracom,esanticipo,porccorredor,concepcargo,impcargo from "
    Sql = Sql & cTabla
    If cSelect <> "" Then Sql = Sql & " where " & cSelect
    Sql = Sql & " group by 1,2,3,4 "
    
    conn.Execute Sql
    
    ' insertamos las lineas de las facturas
    Sql = "insert into tmprlifter (codusu,codsocio,numfactu,fecfactu,numalbar,fechaalb,codvarie,kilosnet,importel,observa1,observa2,observa3,observa4,observa5,prestimado,descontado) "
    Sql = Sql & " select " & vUsu.Codigo & ",rlifter.codsocio,rlifter.numfactu,rlifter.fecfactu,rlifter.numalbar,rlifter.fechaalb,rlifter.codvarie,rlifter.kilosnet,rlifter.importel,observa1,observa2,observa3,observa4,observa5,prestimado,descontado "
    Sql = Sql & " from " & cTabla
    If cSelect <> "" Then Sql = Sql & " where " & cSelect
    Sql = Sql & " group by 1,2,3,4 "
    
    conn.Execute Sql
  
' TODAS LAS CAMPAÑAS ANTERIORES

'[Monica]17/10/2013: Si y solo si no es Montifrut que tiene en otro ariagro otra cosa
    '[Monica]21/01/2015: añado el caso de Natural
If vParamAplic.Cooperativa <> 12 And vParamAplic.Cooperativa <> 9 Then

    SqlBd = "SHOW DATABASES like 'ariagro%' "
    Set RsBd = New ADODB.Recordset
    RsBd.Open SqlBd, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RsBd.EOF
        If Trim(DBLet(RsBd.Fields(0).Value)) <> vEmpresa.BDAriagro And Trim(DBLet(RsBd.Fields(0).Value)) <> "" And InStr(1, DBLet(RsBd.Fields(0).Value), "ariagroutil") = 0 Then
        
            ' borramos la tabla temporal de la campaña anterior
            Sql = "delete from " & Trim(RsBd.Fields(0).Value) & ".tmprcafter where codusu= " & vUsu.Codigo
            conn.Execute Sql
            Sql = "delete from " & Trim(RsBd.Fields(0).Value) & ".tmprlifter where codusu= " & vUsu.Codigo
            conn.Execute Sql
            
            Sql = "insert into " & Trim(RsBd.Fields(0).Value) & ".tmprcafter (codusu,codsocio,numfactu,fecrecep,nomsocio,domsocio,codpobla,pobsocio,prosocio,nifsocio,"
            Sql = Sql & "telsocio,codforpa,brutofac,dtoppago,dtognral,impppago,impgnral,baseiva1,baseiva2,baseiva3,tipoiva1,tipoiva2,tipoiva3,"
            Sql = Sql & "porciva1,porciva2,porciva3,porcrec1,porcrec2,porcrec3,impoiva1,impoiva2,impoiva3,imporec1,imporec2,imporec3,"
            Sql = Sql & "retfacpr,basereten,trefacpr,totalfac,intconta,intracom,esanticipo,porccorredor,concepcargo,impcargo) "
            Sql = Sql & "select " & vUsu.Codigo & ",rcafter.codsocio,rcafter.numfactu,rcafter.fecrecep,rcafter.nomsocio,rcafter.domsocio,rcafter.codpobla,rcafter.pobsocio,rcafter.prosocio,rcafter.nifsocio,rcafter.telsocio,rcafter.codforpa, "
            Sql = Sql & "brutofac,dtoppago,dtognral,impppago,impgnral,baseiva1,baseiva2,baseiva3,tipoiva1,tipoiva2,tipoiva3,porciva1,porciva2,porciva3,"
            Sql = Sql & "porcrec1,porcrec2,porcrec3,impoiva1,impoiva2,impoiva3,imporec1,imporec2,imporec3,retfacpr,basereten,trefacpr,totalfac,"
            Sql = Sql & "intconta,intracom,esanticipo,porccorredor,concepcargo,impcargo"
            Sql = Sql & " from " & Replace(cTabla, vEmpresa.BDAriagro, RsBd.Fields(0).Value)
            If cSelect <> "" Then Sql = Sql & " where " & cSelect
            Sql = Sql & " group by 1,2,3,4 "
            conn.Execute Sql
            
            ' insertamos las lineas de las facturas
            Sql = "insert into " & Trim(RsBd.Fields(0).Value) & ".tmprlifter (codusu,codsocio,numfactu,fecfactu,numalbar,fechaalb,codvarie,kilosnet,importel,observa1,observa2,observa3,observa4,observa5,prestimado,descontado) "
            Sql = Sql & " select " & vUsu.Codigo & ",rlifter.codsocio,rlifter.numfactu,rlifter.fecfactu,rlifter.numalbar,rlifter.fechaalb,rlifter.codvarie,rlifter.kilosnet,rlifter.importel,observa1,observa2,observa3,observa4,observa5,prestimado,descontado "
            Sql = Sql & " from " & Replace(cTabla, vEmpresa.BDAriagro, RsBd.Fields(0).Value)
            If cSelect <> "" Then Sql = Sql & " where " & cSelect
            Sql = Sql & " group by 1,2,3,4 "
            
            conn.Execute Sql
            
            ' introducimos las facturas de la campaña anterior en la temporal de la
            ' campaña actual
            Sql = "insert into tmprcafter select * from " & Trim(RsBd.Fields(0).Value) & ".tmprcafter "
            Sql = Sql & " where codusu = " & vUsu.Codigo
            
            conn.Execute Sql
            
            Sql = "insert into tmprlifter select * from " & Trim(RsBd.Fields(0).Value) & ".tmprlifter "
            Sql = Sql & " where codusu = " & vUsu.Codigo
            
            conn.Execute Sql
            
            
        End If
    
        RsBd.MoveNext
    Wend
  
    Set RsBd = Nothing
End If ' Por Montifrut que tiene la otra en el ariagro2
  
    
    
    CargarFacturas = True
    Screen.MousePointer = vbDefault
    Exit Function
    
eCargarFacturas:
    MuestraError Err.Number, "Cargar Facturas", Err.Description
    Screen.MousePointer = vbDefault
    CargarFacturas = False
End Function




Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(2)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
     For H = 0 To 1
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Next H
     For H = 4 To 5
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Next H

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, H, W
    indFrame = 5
    tabla = "rcafter"
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de socios
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Variedades
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgFec_Click(Index As Integer)
'FEchas
    Dim esq, dalt As Long
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
       
    ' es desplega dalt i cap a la esquerra
    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + 420 + 30

    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
    imgFec(2).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index).Text <> "" Then frmC.NovaData = txtCodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(2).Tag))
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 'Socios Terceros
            AbrirFrmTerceros (Index)
        
        Case 4, 5 'VARIEDADES
            AbrirFrmVariedades (Index)
        
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub Optcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub OptNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'14/02/2007 antes
'    KEYpress KeyAscii
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'socio desde
            Case 1: KEYBusqueda KeyAscii, 1 'socio hasta
            Case 4: KEYBusqueda KeyAscii, 4 'variedad desde
            Case 5: KEYBusqueda KeyAscii, 5 'variedad hasta
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFec_Click (Indice)
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 0, 1 'socio tercero
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
        Case 2, 3 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 4, 5 'VARIEDAD
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
            
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 5925
        Me.FrameCobros.Width = 6555
        W = Me.FrameCobros.Width
        H = Me.FrameCobros.Height
    End If
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    CadParam = ""
    numParam = 0
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
        .NombreRPT = cadNombreRPT
        .ConSubInforme = False
        .EnvioEMail = False
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmTerceros(Indice As Integer)
    indCodigo = Indice
    Set frmSoc = New frmManSocios
    frmSoc.DatosADevolverBusqueda = "0|1|"
    frmSoc.Show vbModal
    Set frmSoc = Nothing
End Sub

Private Sub AbrirFrmVariedades(Indice As Integer)
    indCodigo = Indice
    Set frmVar = New frmComVar
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.CodigoActual = txtCodigo(indCodigo)
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub

Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
        '.SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
        .OtrosParametros = CadParam
        .NumeroParametros = numParam
        '##descomen
'        .MostrarTree = MostrarTree
'        .Informe = MIPATH & Nombre
'        .InfConta = InfConta
        '##
        
'        If NombreSubRptConta <> "" Then
'            .SubInformeConta = NombreSubRptConta
'        Else
'            .SubInformeConta = ""
'        End If
        '##descomen
'        .ConSubInforme = ConSubInforme
        '##
        .Opcion = ""
'        .ExportarPDF = (chkEMAIL.Value = 1)
        .Show vbModal
    End With
    
'    If Me.chkEMAIL.Value = 1 Then
'    '####Descomentar
'        If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
'    End If
    Unload Me
End Sub

Private Sub AbrirEMail()
    If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
End Sub


Private Function HayRegistros(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
Dim Rs As ADODB.Recordset

    Sql = "Select * FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Rs.EOF Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        HayRegistros = False
    Else
        HayRegistros = True
    End If

End Function

Private Function ProcesarCambios(cadWHERE As String) As Boolean
Dim Sql As String
Dim Sql1 As String
Dim I As Integer
Dim HayReg As Integer
Dim B As Boolean

On Error GoTo eProcesarCambios

    HayReg = 0
    
    conn.Execute "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
        
    If cadWHERE <> "" Then
        cadWHERE = QuitarCaracterACadena(cadWHERE, "{")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "}")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "_1")
    End If
        
    Sql = "insert into tmpinformes (codusu, codigo1) select " & DBSet(vUsu.Codigo, "N")
    Sql = Sql & ", rhisfruta.numalbar from rhisfruta, rsocios where  numalbar not in (select numalbar from rlifter) "
    Sql = Sql & " and rsocios.tipoprod = 1 and rhisfruta.codsocio = rsocios.codsocio "
    
    If cadWHERE <> "" Then Sql = Sql & " and " & cadWHERE
    
    
    conn.Execute Sql
        
    ProcesarCambios = HayRegistros("tmpinformes", "codusu = " & vUsu.Codigo)

eProcesarCambios:
    If Err.Number <> 0 Then
        ProcesarCambios = False
    End If
End Function


Private Sub InsertaLineaEnTemporal(ByRef ItmX As ListItem)
Dim Sql As String
Dim Codmacta As String
Dim Rs As ADODB.Recordset
Dim Sql1 As String

        Sql1 = "insert into tmpinformes(codusu, codigo1) values ("
        Sql1 = Sql1 & DBSet(vUsu.Codigo, "N") & "," & DBSet(ItmX.Text, "N") & ")"

        conn.Execute Sql1
    
End Sub

