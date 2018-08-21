VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTRAInfHcoEntradas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6705
   Icon            =   "frmTRAInfHcoEntradas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameFacturar 
      Height          =   5070
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6615
      Begin VB.CheckBox Check1 
         Caption         =   "No mostrar socio "
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
         Left            =   390
         TabIndex        =   25
         Top             =   3840
         Width           =   2205
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   4
         Top             =   2880
         Width           =   1350
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   5
         Top             =   3285
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
         Left            =   5145
         TabIndex        =   7
         Top             =   4275
         Width           =   1065
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   0
         Top             =   1095
         Width           =   1350
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1455
         Width           =   1350
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
         Index           =   12
         Left            =   3015
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   1095
         Width           =   3285
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
         Index           =   13
         Left            =   3015
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text5"
         Top             =   1455
         Width           =   3285
      End
      Begin VB.CommandButton Command3 
         Height          =   440
         Left            =   7860
         Picture         =   "frmTRAInfHcoEntradas.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command4 
         Height          =   440
         Left            =   7860
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   2
         Top             =   2040
         Width           =   870
      End
      Begin VB.TextBox txtcodigo 
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
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   3
         Top             =   2400
         Width           =   870
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
         Index           =   20
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text5"
         Top             =   2040
         Width           =   3825
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
         Index           =   21
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text5"
         Top             =   2400
         Width           =   3825
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
         Left            =   3975
         TabIndex        =   6
         Top             =   4290
         Width           =   1065
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1350
         Picture         =   "frmTRAInfHcoEntradas.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   3300
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00972E0B&
         Height          =   240
         Index           =   19
         Left            =   405
         TabIndex        =   24
         Top             =   2700
         Width           =   600
      End
      Begin VB.Label Label2 
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
         Index           =   20
         Left            =   645
         TabIndex        =   23
         Top             =   2940
         Width           =   690
      End
      Begin VB.Label Label2 
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
         Index           =   21
         Left            =   645
         TabIndex        =   22
         Top             =   3285
         Width           =   645
      End
      Begin VB.Label Label2 
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
         Index           =   22
         Left            =   645
         TabIndex        =   21
         Top             =   1095
         Width           =   690
      End
      Begin VB.Label Label2 
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
         Index           =   23
         Left            =   645
         TabIndex        =   20
         Top             =   1455
         Width           =   645
      End
      Begin VB.Label Label3 
         Caption         =   "Informe de Entradas de Transporte"
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
         Left            =   375
         TabIndex        =   19
         Top             =   345
         Width           =   5925
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00972E0B&
         Height          =   240
         Index           =   27
         Left            =   390
         TabIndex        =   18
         Top             =   810
         Width           =   1320
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   1350
         MouseIcon       =   "frmTRAInfHcoEntradas.frx":0122
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar transportista"
         Top             =   1095
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   1350
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar transportista"
         Top             =   1485
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1350
         Picture         =   "frmTRAInfHcoEntradas.frx":0126
         ToolTipText     =   "Buscar fecha"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clase"
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
         Left            =   390
         TabIndex        =   17
         Top             =   1785
         Width           =   525
      End
      Begin VB.Label Label2 
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
         Index           =   18
         Left            =   645
         TabIndex        =   16
         Top             =   2040
         Width           =   690
      End
      Begin VB.Label Label2 
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
         Index           =   28
         Left            =   645
         TabIndex        =   15
         Top             =   2430
         Width           =   645
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   20
         Left            =   1350
         MouseIcon       =   "frmTRAInfHcoEntradas.frx":01B1
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   21
         Left            =   1350
         MouseIcon       =   "frmTRAInfHcoEntradas.frx":0303
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2430
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7800
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTRAInfHcoEntradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DatoSeleccionado(CadenaSeleccion As String)

Public OpcionListado As Integer
'(ver opciones en frmListado)

      
'Alguna opcion mas
'                   1000.-  Es cuando paso pedido a albaran y este a factura en el mismo proceso
'                   1001.-  Facturar un unico albaran
      
      
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir


Public CodClien As String 'Para seleccionar inicialmente las ofertas del Proveedor

'#Laura 14/11/2006 Recuperar facturas Alzira
Public EstaRecupFact As Boolean ' si esta recuperando facturas (para albaranes de mostrador)


Private WithEvents frmCla As frmBasico2 'Ayuda de Clases de comercial
Attribute frmCla.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'Mensajes
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1


'Private HaDevueltoDatos As Boolean
Private NomTabla As String
Private NomTablaLin As String

'Private WithEvents frmMtoCartasOfe As frmFacCartasOferta
Private WithEvents frmTra As frmManTranspor
Attribute frmTra.VB_VarHelpID = -1
'Private WithEvents frmB As frmBuscaGrid  'Busquedas
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1

'----- Variables para el INforme ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String
Private numParam As Byte
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private Titulo As String 'Titulo informe que se pasa a frmImprimir
Private nomRPT As String 'nombre del fichero .rpt a imprimir
Private conSubRPT As Boolean 'si tiene subinformes para enlazarlos a las tablas correctas
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe
Private ConSubInforme As Boolean
'-------------------------------------



Dim Tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String

Dim Indice As Integer



Dim indCodigo As Integer 'indice para txtCodigo

Dim PrimeraVez As Boolean

Dim vSeccion As CSeccion

Dim cadSelect1 As String


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub






Private Function DatosOk() As Boolean
Dim B As Boolean
Dim SQL As String
Dim Sql2 As String
Dim vClien As cSocio
' añadido
Dim Mens As String
Dim numfactu As String
Dim numser As String
Dim Fecha As Date
Dim vCont As CTiposMov
Dim tipoMov As String

    B = True
    
'            If txtcodigo(6).Text = "" Then
'                MsgBox "Debe introducir obligatoriamente la Fecha desde.", vbExclamation
'                b = False
'                PonerFoco txtcodigo(6)
'            End If
    DatosOk = B

End Function


Private Sub cmdAceptar_Click()
'Facturacion de Albaranes
Dim campo As String, cad As String
Dim cadFrom As String
Dim cadSQL As String 'Para seleccionar los Albaranes del rango seleccion
                      'que no se van a facturar
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim cadDesde As Date
Dim cadhasta As Date

Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim nTabla As String
Dim Tabla1 As String


Dim Nregs As Long
Dim FecFac As Date
Dim TipoPrec As Byte ' 0 anticipos
                     ' 1 liquidaciones
Dim B As Boolean
Dim Sql2 As String



    InicializarVbles
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If DatosOk Then
        '======== FORMULA  ====================================
        'D/H TRANSPORTISTA
        cDesde = Trim(txtCodigo(12).Text)
        cHasta = Trim(txtCodigo(13).Text)
        nDesde = txtNombre(12).Text
        nHasta = txtNombre(13).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{rclasifica.codtrans}"
            TipCod = "T"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
        End If
        
        'D/H CLASE
        cDesde = Trim(txtCodigo(20).Text)
        cHasta = Trim(txtCodigo(21).Text)
        nDesde = txtNombre(20).Text
        nHasta = txtNombre(21).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{variedades.codclase}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
        End If
        
        Sql2 = ""
        If txtCodigo(20).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase >=" & DBSet(txtCodigo(20).Text, "N")
        If txtCodigo(21).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase <=" & DBSet(txtCodigo(21).Text, "N")
        
        
        'D/H fecha
        cDesde = Trim(txtCodigo(6).Text)
        cHasta = Trim(txtCodigo(7).Text)
        nDesde = ""
        nHasta = ""
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{rclasifica.fechaent}"
            TipCod = "F"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
        End If
        
        nTabla = "(((rclasifica "
        nTabla = nTabla & " INNER JOIN rtransporte ON rclasifica.codtrans = rtransporte.codtrans) "
        nTabla = nTabla & " INNER JOIN variedades ON rclasifica.codvarie = variedades.codvarie) "
        nTabla = nTabla & " INNER JOIN productos ON variedades.codprodu = productos.codprodu) "
        nTabla = nTabla & " INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        nTabla = nTabla & " and grupopro.codgrupo <> 5 and grupopro.codgrupo <> 6 " ' grupo no puede ser 5=almazara ni 6=bodega
        
        If Not AnyadirAFormula(cadSelect, "{rclasifica.transportadopor} = 0 ") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rclasifica.transportadopor} = 0 ") Then Exit Sub
        
        If Not AnyadirAFormula(cadSelect, "{rclasifica.tipoentr} <> 1 ") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rclasifica.tipoentr} <> 1 ") Then Exit Sub
        
        cadSelect1 = Replace(Replace(Replace(cadSelect, "rclasifica", "rhisfruta_entradas"), "rhisfruta_entradas.transportadopor", "rhisfruta.transportadopor"), "rhisfruta_entradas.tipoentr", "rhisfruta.tipoentr")
        Tabla1 = "((((rhisfruta INNER JOIN rhisfruta_entradas on rhisfruta.numalbar = rhisfruta_entradas.numalbar) "
        Tabla1 = Tabla1 & " INNER JOIN rtransporte ON rhisfruta_entradas.codtrans = rtransporte.codtrans) "
        Tabla1 = Tabla1 & " INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie) "
        Tabla1 = Tabla1 & " INNER JOIN productos ON variedades.codprodu = productos.codprodu) "
        Tabla1 = Tabla1 & " INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        Tabla1 = Tabla1 & " and grupopro.codgrupo <> 5 and grupopro.codgrupo <> 6 " ' grupo no puede ser 5=almazara ni 6=bodega
        
        
        cadDesde = "01/01/1900"
        cadhasta = "31/12/2500"
        
        If txtCodigo(6).Text <> "" Then cadDesde = CDate(txtCodigo(6).Text)
        If txtCodigo(7).Text <> "" Then cadhasta = CDate(txtCodigo(7).Text)
        
        CadParam = CadParam & "pFecDesde= Date(" & Year(cadDesde) & "," & Month(cadDesde) & "," & Day(cadDesde) & ")" & "|" 'txtcodigo(6).Text & """|"
        CadParam = CadParam & "pFecHasta= Date(" & Year(cadhasta) & "," & Month(cadhasta) & "," & Day(cadhasta) & ")" & "|" 'txtcodigo(7).Text & """|"
        numParam = numParam + 2
        
        
'        If Not AnyadirAFormula(cadSelect, "{rhisfruta.numalbar} not in (select numalbar from rfacttra_albaran) ") Then Exit Sub
'
        '[Monica]04/11/2013: no sacamos el nombre del socio si es Catadau y lo marcan
        CadParam = CadParam & "pQuitarSocio=" & Check1.Value & "|"
        numParam = numParam + 1
        
        cadTitulo = "Informe Entradas de Transporte"
                
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 16
        frmMens.cadWHERE = Sql2
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        conSubRPT = True
        'Comprobar si hay registros a Mostrar antes de abrir el Informe
'        If HayRegParaInforme(nTabla, cadSelect) Then
'            LlamarImprimir
'        End If
        
        If ProcesoEntradasTransportista(nTabla, cadSelect, Tabla1, cadSelect1) Then
            If HayRegParaInforme("tmpinformes", "{tmpinformes.codusu} = " & vUsu.Codigo) Then
                cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo

                'Nombre fichero .rpt a Imprimir
                indRPT = 58 ' informe de entradas por transportista
                
                If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                
                frmImprimir.NombreRPT = nomDocu
                
                cadNombreRPT = nomDocu '"rInfEntradasTrans.rpt"
                
                ConSubInforme = True
                LlamarImprimir
            Else
                MsgBox "No hay registros entre esos límites.", vbExclamation
            End If
        End If
        
    End If
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(12)
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim i As Integer
Dim indFrame As Single

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me

    'Ocultar todos los Frames de Formulario
    Me.FrameFacturar.visible = False
    
    
    For i = 12 To 13
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 20 To 21
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i

    ' Necesitamos la conexion a la contabilidad de la seccion de adv
    ' para sacar los porcentajes de iva de los articulos y calcular
    ' los datos de la factura
    
    
    NomTabla = "rhisfruta"
    NomTablaLin = "rhisfruta_entradas"
        
'    OpcionListado = 52
    
    PonerFrameFacVisible True, H, W
    txtCodigo(7).Text = Format(Now - 1, "dd/mm/yyyy")
    indFrame = 6
    
    
    '[Monica]04/11/2013: no mostramos el socio cuando el informe es para el transportista (solo catadau)
    Me.Check1.visible = (vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19)
    Me.Check1.Enabled = (vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19)
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not vSeccion Is Nothing Then
        vSeccion.CerrarConta
        Set vSeccion = Nothing
    End If
End Sub


Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtCodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmCla_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") ' codigo de clase
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) ' descripcion
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim SQL As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        SQL = " {variedades.codvarie} in (" & CadenaSeleccion & ")"
        Sql2 = " {variedades.codvarie} in [" & CadenaSeleccion & "]"
    Else
        SQL = " {variedades.codvarie} = -1 "
    End If
    If Not AnyadirAFormula(cadSelect, SQL) Then Exit Sub
    If Not AnyadirAFormula(cadSelect1, SQL) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub



Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub



Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Articulos
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFPa_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Formas de Pabo
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscarOfer_Click(Index As Integer)
    Select Case Index
            
        Case 14, 15, 20, 21 'Cod. Socio
'            Select Case Index
'                Case 11, 12: indCodigo = Index + 9
'                Case 14, 15: indCodigo = Index + 14
'                Case 20, 21: indCodigo = Index + 20
'                Case 27, 28: indCodigo = Index + 21
'                Case 32: indCodigo = 8
'            End Select
'            Set frmSoc = New frmManSocios
'            frmSoc.DatosADevolverBusqueda = "0|2|"
'            If Not IsNumeric(txtCodigo(indCodigo).Text) Then txtCodigo(indCodigo).Text = ""
'            frmSoc.Show vbModal
'            Set frmSoc = Nothing
            
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub


Private Sub imgFecha_Click(Index As Integer)
   
'++monica

   '++monica
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object

    Set frmF = New frmCal
    
    esq = imgFec(Index).Left
    dalt = imgFec(Index).Top
    
    Set obj = imgFec(Index).Container

    While imgFec(Index).Parent.Name <> obj.Name
        esq = esq + obj.Left
        dalt = dalt + obj.Top
        Set obj = obj.Container
    Wend
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmF.Left = esq + imgFec(Index).Parent.Left + 30
    frmF.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40
   
   frmF.NovaData = Now
   
   Select Case Index
        Case 0 'FramePreFacturar
            indCodigo = 6
        Case 1 'FramePreFacturar
            indCodigo = 7
        Case 15 'Frame Factura
            indCodigo = 15
   End Select
   
   PonerFormatoFecha txtCodigo(indCodigo)
   If txtCodigo(indCodigo).Text <> "" Then frmF.NovaData = CDate(txtCodigo(indCodigo).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtCodigo(indCodigo)


'++
'
'
'
'
'
'   Screen.MousePointer = vbHourglass
'   Set frmF = New frmCal
'   frmF.Fecha = Now
'
'
'   Select Case Index
'        Case 10 'FramePreFacturar
'            indCodigo = 26
'        Case 11 'FramePreFacturar
'            indCodigo = 27
'        Case 12 'Frame Factura
'            indCodigo = 38
'        Case 13 'Frame Factura
'            indCodigo = 39
'        Case 14 'FrameFactura
'            indCodigo = 34
'   End Select
'
'   PonerFormatoFecha txtCodigo(indCodigo)
'   If txtCodigo(indCodigo).Text <> "" Then frmF.Fecha = CDate(txtCodigo(indCodigo).Text)
'
'   Screen.MousePointer = vbDefault
'   frmF.Show vbModal
'   Set frmF = Nothing
'   PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub OptTipoInf_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub OptDetalle_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 18, 19, 20, 21, 28, 29 'Clases
            AbrirFrmClase (Index)
        
        Case 0, 1, 12, 13, 16, 17, 24, 25 'transportistas
            AbrirFrmTransportistas (Index)
        
    End Select
    PonerFoco txtCodigo(indCodigo)
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
            Indice = 6
        Case 1
            Indice = 7
        Case 2
            Indice = 15
        Case 3, 4
            Indice = Index - 1
    End Select

    imgFec(0).Tag = Indice '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtCodigo(Indice).Text <> "" Then frmC.NovaData = txtCodigo(Indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtCodigo(CByte(imgFec(0).Tag)) '<===
    ' ********************************************


End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'15/02/2007
'    KEYpress KeyAscii
'ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'transportista desde
            Case 1: KEYBusqueda KeyAscii, 1 'transportista hasta
            Case 12: KEYBusqueda KeyAscii, 12 'transportista desde
            Case 13: KEYBusqueda KeyAscii, 13 'transportista hasta
            Case 20: KEYBusqueda KeyAscii, 20 'clase desde
            Case 21: KEYBusqueda KeyAscii, 21 'clase hasta
            
            Case 2: KEYFecha KeyAscii, 3 'fecha desde
            Case 3: KEYFecha KeyAscii, 4 'fecha hasta
            Case 6: KEYFecha KeyAscii, 0 'fecha desde
            Case 7: KEYFecha KeyAscii, 1 'fecha hasta
            Case 15: KEYFecha KeyAscii, 2 'fecha hasta
            
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
Dim devuelve As String
Dim codCampo As String, nomCampo As String
Dim Tabla As String
      
    Select Case Index
        Case 0, 1, 12, 13 'transportistas
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rtransporte", "nomtrans", "codtrans", "T")
    
    
        'FECHA Desde Hasta
        Case 2, 3, 6, 7, 15
            If txtCodigo(Index).Text <> "" Then
                PonerFormatoFecha txtCodigo(Index)
            End If
        
        Case 4, 5
            PonerFormatoEntero txtCodigo(Index)
        
        
        Case 36, 37  'Nº de Parte
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
            End If
            
        Case 20, 21
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "clases", "nomclase", "codclase", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
        
        
        Case 40, 41  'Cod. Socio
            If PonerFormatoEntero(txtCodigo(Index)) Then
                nomCampo = "nomsocio"
                Tabla = "rsocios"
                codCampo = "codsocio"
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), Tabla, nomCampo, codCampo, "N")
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
            Else
                txtNombre(Index).Text = ""
            End If
            
        Case 42  'Cod. Formas de PAGO de comercial
            If PonerFormatoEntero(txtCodigo(Index)) Then
'                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "forpago", "nomforpa", "codforpa", "N")
'[Monica] 09/02/2010 no es de comercial sino de la contabilidad de adv
                If vParamAplic.ContabilidadNueva Then
                    txtNombre(Index).Text = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtCodigo(Index), "N")
                Else
                    txtNombre(Index).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtCodigo(Index), "N")
                End If
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            Else
                txtNombre(Index).Text = ""
            End If
        
    End Select
End Sub



Private Sub PonerFrameFacVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Facturacion de Albaran Visible y Ajustado al Formulario, y visualiza los controles
Dim cad As String

    H = 5070
    W = 6735
    
    
    PonerFrameVisible Me.FrameFacturar, visible, H, W
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


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    CadParam = ""
    numParam = 0
End Sub


Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = CadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 0
        .Titulo = cadTitulo
        .ConSubInforme = conSubRPT
        .NombreRPT = cadNombreRPT  'nombre del informe
        .Show vbModal
    End With
End Sub

Private Sub txtCodigo_Validate(Index As Integer, Cancel As Boolean)
'    Select Case Index
'           Case 15, 16 'ARTICULO
'            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "sartic", "nomartic", "codartic", "Articulo", "T")
'            If txtNombre(Index).Text = "" And txtcodigo(Index) <> "" Then Cancel = True
'    End Select
End Sub

Private Function ObtenerClientes(cadW As String, Importe As String) As String
Dim SQL As String
Dim Rs As ADODB.Recordset

    On Error GoTo EClientes
    
    cadW = Replace(cadW, "{", "")
    cadW = Replace(cadW, "}", "")
    
    SQL = "select codclien,nomclien,sum(baseimp1),sum(baseimp2),sum(baseimp3),sum(baseimp1)+ sum(if(isnull(baseimp2),0,baseimp2))+ sum(if(isnull(baseimp3),0,baseimp3)) as BaseImp"
    SQL = SQL & " From scafac "
    If cadW <> "" Then SQL = SQL & " where " & cadW
    SQL = SQL & " group by codclien "
    If Importe <> "" Then SQL = SQL & "having baseimp>" & Importe
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    While Not Rs.EOF
'        If RS!BaseImp >= CCur(Importe) Then
            SQL = SQL & Rs!CodClien & ","
'        End If
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    If SQL <> "" Then
        SQL = Mid(SQL, 1, Len(SQL) - 1)
        SQL = "( {scafac.codclien} IN [" & SQL & "] )"
    End If
    ObtenerClientes = SQL
    
EClientes:
   If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
End Function




Private Function ActualizarRegistrosFac(cTabla As String, cWhere As String) As Boolean
'Actualizar la marca de impreso
Dim SQL As String

    On Error GoTo eActualizarRegistros

    ActualizarRegistrosFac = False
    SQL = "update " & cTabla & ", usuarios.stipom set impreso = 1 "
    SQL = SQL & " where usuarios.stipom.codtipom = rfactsoc.codtipom "
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " and " & cWhere
    End If
    
    conn.Execute SQL
    
    ActualizarRegistrosFac = True
    Exit Function
    
eActualizarRegistros:
    MuestraError Err.Number, "Actualizando registros", Err.Description
End Function



Private Sub AbrirFrmClase(Indice As Integer)
    indCodigo = Indice
    Set frmCla = New frmBasico2
    
    AyudaClasesCom frmCla, txtCodigo(Indice).Text
    
    Set frmCla = Nothing
End Sub

Private Sub AbrirFrmTransportistas(Indice As Integer)
    indCodigo = Indice
    Set frmTra = New frmManTranspor
    frmTra.DatosADevolverBusqueda = "0|1|"
    frmTra.Show vbModal
    Set frmTra = Nothing
End Sub


Private Function ProcesoEntradasTransportista(cTabla As String, cWhere As String, ctabla1 As String, cwhere1 As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset

    On Error GoTo eProcesoEntradasTransportista
    
    Screen.MousePointer = vbHourglass
    
    ProcesoEntradasTransportista = False

    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    ctabla1 = QuitarCaracterACadena(ctabla1, "{")
    ctabla1 = QuitarCaracterACadena(ctabla1, "}")
    
    cWhere = QuitarCaracterACadena(cWhere, "{")
    cWhere = QuitarCaracterACadena(cWhere, "}")
    
    cwhere1 = QuitarCaracterACadena(cwhere1, "{")
    cwhere1 = QuitarCaracterACadena(cwhere1, "}")
    
    SQL = "select " & vUsu.Codigo & ", rclasifica.codtrans,rclasifica.codvarie,rclasifica.codsocio,rclasifica.codcampo,rclasifica.fechaent,rclasifica.kilosnet,rclasifica.kilostra,rclasifica.impacarr, rclasifica.codtarif, rclasifica.numnotac from " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        SQL = SQL & " WHERE " & cWhere
    End If
    SQL = SQL & " union "
    SQL = SQL & "select " & vUsu.Codigo & ", rhisfruta_entradas.codtrans,rhisfruta.codvarie,rhisfruta.codsocio,rhisfruta.codcampo,rhisfruta_entradas.fechaent,rhisfruta_entradas.kilosnet,rhisfruta_entradas.kilostra,rhisfruta_entradas.impacarr, rhisfruta_entradas.codtarif, rhisfruta_entradas.numnotac from " & QuitarCaracterACadena(ctabla1, "_1")
    If cwhere1 <> "" Then
        SQL = SQL & " WHERE " & cwhere1
    End If
    SQL = SQL & " order by 1, 2, 3 "
                                           'transpor,codvarie, codsocio, codcampo, fechaent, kilosnet, kilostra, impacarr, codtarif,  numnotac
    Sql2 = "insert into tmpinformes (codusu, nombre1, codigo1, importe1, importe2, fecha1,  importe3,  importe4, importe5, importeb1, importeb2) "
    Sql2 = Sql2 & SQL
    
    conn.Execute Sql2
    
    Screen.MousePointer = vbDefault
    
    ProcesoEntradasTransportista = True
    Exit Function
    
eProcesoEntradasTransportista:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Proceso de Entradas Transportista", Err.Description
End Function

