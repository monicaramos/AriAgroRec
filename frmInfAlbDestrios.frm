VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInfAlbDestrios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6705
   Icon            =   "frmInfAlbDestrios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7800
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameFacturar 
      Height          =   5940
      Left            =   0
      TabIndex        =   10
      Top             =   -30
      Width           =   6615
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "Text5"
         Top             =   4140
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "Text5"
         Top             =   3780
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   7
         Top             =   4140
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   6
         Top             =   3780
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Resumen"
         Height          =   225
         Left            =   420
         TabIndex        =   27
         Top             =   4620
         Width           =   3105
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   4
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   5
         Top             =   3285
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5190
         TabIndex        =   9
         Top             =   5325
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   1650
         MaxLength       =   6
         TabIndex        =   0
         Top             =   1095
         Width           =   990
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   1650
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1455
         Width           =   990
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Text5"
         Top             =   1095
         Width           =   3195
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Text5"
         Top             =   1455
         Width           =   3195
      End
      Begin VB.CommandButton Command3 
         Height          =   440
         Left            =   7860
         Picture         =   "frmInfAlbDestrios.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command4 
         Height          =   440
         Left            =   7860
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   20
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   2
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   21
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   3
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   20
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   21
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   2400
         Width           =   3375
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4110
         TabIndex        =   8
         Top             =   5340
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   390
         TabIndex        =   33
         Top             =   4980
         Visible         =   0   'False
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Cargando Temporal"
         Height          =   195
         Index           =   24
         Left            =   420
         TabIndex        =   34
         Top             =   5280
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1350
         MouseIcon       =   "frmInfAlbDestrios.frx":0097
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar plaga"
         Top             =   4170
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1350
         MouseIcon       =   "frmInfAlbDestrios.frx":01E9
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar plaga"
         Top             =   3780
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   2
         Left            =   765
         TabIndex        =   32
         Top             =   4215
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   1
         Left            =   765
         TabIndex        =   31
         Top             =   3825
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Plaga"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   30
         Top             =   3570
         Width           =   405
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1350
         Picture         =   "frmInfAlbDestrios.frx":033B
         ToolTipText     =   "Buscar fecha"
         Top             =   3300
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   19
         Left            =   405
         TabIndex        =   26
         Top             =   2700
         Width           =   450
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   20
         Left            =   735
         TabIndex        =   25
         Top             =   2940
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   21
         Left            =   735
         TabIndex        =   24
         Top             =   3285
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   22
         Left            =   720
         TabIndex        =   23
         Top             =   1140
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   23
         Left            =   720
         TabIndex        =   22
         Top             =   1500
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Informe de Destrios Varios"
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
         Left            =   420
         TabIndex        =   21
         Top             =   345
         Width           =   5925
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   27
         Left            =   390
         TabIndex        =   20
         Top             =   900
         Width           =   405
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   1350
         MouseIcon       =   "frmInfAlbDestrios.frx":03C6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1095
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   1350
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1485
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1350
         Picture         =   "frmInfAlbDestrios.frx":03CA
         ToolTipText     =   "Buscar fecha"
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clase"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   11
         Left            =   390
         TabIndex        =   19
         Top             =   1830
         Width           =   390
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   18
         Left            =   765
         TabIndex        =   18
         Top             =   2085
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   28
         Left            =   765
         TabIndex        =   17
         Top             =   2475
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   20
         Left            =   1350
         MouseIcon       =   "frmInfAlbDestrios.frx":0455
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   21
         Left            =   1350
         MouseIcon       =   "frmInfAlbDestrios.frx":05A7
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2430
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmInfAlbDestrios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DatoSeleccionado(CadenaSeleccion As String)
      
Private WithEvents frmCla As frmComercial 'Ayuda de Clases de comercial
Attribute frmCla.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'Mensajes
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmMens1 As frmMensajes 'Mensajes
Attribute frmMens1.VB_VarHelpID = -1
Private WithEvents frmInc As frmManInciden ' incidencias (plagas)
Attribute frmInc.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1


'Private HaDevueltoDatos As Boolean
Private NomTabla As String
Private NomTablaLin As String

Private WithEvents frmSoc As frmManSocios
Attribute frmSoc.VB_VarHelpID = -1
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

Dim cadSelect1 As String
Dim cadSelect2 As String
Dim cadSelectBorra As String

Dim Tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String

Dim indice As Integer



Dim indCodigo As Integer 'indice para txtCodigo

Dim PrimeraVez As Boolean

Dim vSeccion As CSeccion

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
    
    b = True
    
    DatosOk = b

End Function


Private Sub cmdAceptar_Click()
'Facturacion de Albaranes
Dim campo As String, Cad As String
Dim cadFrom As String
Dim cadSQL As String 'Para seleccionar los Albaranes del rango seleccion
                      'que no se van a facturar
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim cadDesde As Date
Dim cadhasta As Date

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim nTabla As String
Dim Tabla1 As String
Dim nTabla2 As String

Dim Nregs As Long
Dim FecFac As Date

Dim b As Boolean
Dim Sql2 As String
Dim Sql3 As String

Dim CadFechas As String

    InicializarVbles
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If DatosOk Then
        '======== FORMULA  ====================================
        'D/H Socios
        cDesde = Trim(txtcodigo(12).Text)
        cHasta = Trim(txtcodigo(13).Text)
        nDesde = txtNombre(12).Text
        nHasta = txtNombre(13).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{rhisfruta.codsocio}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
        End If
        
        'D/H CLASE
        cDesde = Trim(txtcodigo(20).Text)
        cHasta = Trim(txtcodigo(21).Text)
        nDesde = txtNombre(20).Text
        nHasta = txtNombre(21).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{variedades.codclase}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
        End If
        
        Sql2 = ""
        If txtcodigo(20).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase >=" & DBSet(txtcodigo(20).Text, "N")
        If txtcodigo(21).Text <> "" Then Sql2 = Sql2 & " and variedades.codclase <=" & DBSet(txtcodigo(21).Text, "N")
        
        
        'D/H fecha
        cDesde = Trim(txtcodigo(6).Text)
        cHasta = Trim(txtcodigo(7).Text)
        nDesde = ""
        nHasta = ""
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{rhisfruta.fecalbar}"
            TipCod = "F"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
        End If
        
        'D/H incidencia
        cDesde = Trim(txtcodigo(0).Text)
        cHasta = Trim(txtcodigo(1).Text)
        nDesde = txtNombre(0).Text
        nHasta = txtNombre(1).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{rhisfruta_incidencia.codincid}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHIncidencia=""") Then Exit Sub
        End If
        
        Sql3 = ""
        If txtcodigo(0).Text <> "" Then Sql3 = Sql3 & " and rincidencia.codincid >=" & DBSet(txtcodigo(0).Text, "N")
        If txtcodigo(1).Text <> "" Then Sql3 = Sql3 & " and rincidencia.codincid <=" & DBSet(txtcodigo(1).Text, "N")
        
        
        
        nTabla = "((((((rhisfruta "
        nTabla = nTabla & " INNER JOIN rhisfruta_entradas ON rhisfruta.numalbar = rhisfruta_entradas.numalbar) "
        nTabla = nTabla & " INNER JOIN rhisfruta_incidencia ON rhisfruta_entradas.numalbar = rhisfruta_incidencia.numalbar and rhisfruta_entradas.numnotac = rhisfruta_incidencia.numnotac) "
        nTabla = nTabla & " INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio) "
        nTabla = nTabla & " INNER JOIN rincidencia ON rhisfruta_incidencia.codincid = rincidencia.codincid) "
        nTabla = nTabla & " INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie) "
        nTabla = nTabla & " INNER JOIN productos ON variedades.codprodu = productos.codprodu) "
        nTabla = nTabla & " INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        nTabla = nTabla & " and grupopro.codgrupo <> 5 and grupopro.codgrupo <> 6 " ' grupo no puede ser 5=almazara ni 6=bodega
        
        
        cadDesde = "01/01/1900"
        cadhasta = "31/12/2500"
        
        If txtcodigo(6).Text <> "" Then cadDesde = CDate(txtcodigo(6).Text)
        If txtcodigo(7).Text <> "" Then cadhasta = CDate(txtcodigo(7).Text)
        
        CadParam = CadParam & "pFecDesde= Date(" & Year(cadDesde) & "," & Month(cadDesde) & "," & Day(cadDesde) & ")" & "|" 'txtcodigo(6).Text & """|"
        CadParam = CadParam & "pFecHasta= Date(" & Year(cadhasta) & "," & Month(cadhasta) & "," & Day(cadhasta) & ")" & "|" 'txtcodigo(7).Text & """|"
        numParam = numParam + 2
        
        
        CadParam = CadParam & "pResumen=" & Check1.Value & "|"
        numParam = numParam + 1
        
        
        cadTitulo = "Informe Destrios Varios"
                
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 16
        frmMens.cadWHERE = Sql2
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        
        
        Set frmMens1 = New frmMensajes
        
        frmMens1.OpcionMensaje = 27
        frmMens1.cadWHERE = Sql3
        frmMens1.Show vbModal
        
        Set frmMens1 = Nothing
        
        
        
        conSubRPT = True
        
        If ProcesoEntradas(nTabla, cadSelect) Then
            If HayRegParaInforme("tmpinformes", "{tmpinformes.codusu} = " & vUsu.Codigo) Then
                cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo

                
                'Nombre fichero .rpt a Imprimir
                indRPT = 73 ' informe de entradas por transportista

                If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub

                frmImprimir.NombreRPT = nomDocu
                cadNombreRPT = nomDocu '"rInfAlbDestrios.rpt"
                
                If Check1.Value = 0 Then
                    cadNombreRPT = Replace(cadNombreRPT, ".rpt", "1.rpt") '"rInfAlbDestrios1.rpt"
                Else
                    cadNombreRPT = cadNombreRPT
                End If
                
                
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
        PonerFoco txtcodigo(12)
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
    
    For i = 0 To 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 12 To 13
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 20 To 21
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    NomTabla = "rhisfruta"
    NomTablaLin = "rhisfruta_entradas"
    
    PonerFrameFacVisible True, H, W
    indFrame = 6
    
    Me.Pb1.visible = False
    Label2(24).visible = False
    
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
    txtcodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmCla_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") ' codigo de clase
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) ' descripcion
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " {variedades.codvarie} in (" & CadenaSeleccion & ")"
        Sql2 = " {variedades.codvarie} in [" & CadenaSeleccion & "]"
    Else
        Sql = " {variedades.codvarie} = -1 "
    End If
    If Not AnyadirAFormula(cadSelect, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadSelect1, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadSelect2, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub



Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtcodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub



Private Sub imgFecha_Click(Index As Integer)
   
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
   
   PonerFormatoFecha txtcodigo(indCodigo)
   If txtcodigo(indCodigo).Text <> "" Then frmF.NovaData = CDate(txtcodigo(indCodigo).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtcodigo(indCodigo)
End Sub


Private Sub frmMens1_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    cadSelectBorra = ""
    If CadenaSeleccion <> "" Then
        Sql = " {rincidencia.codincid} in (" & CadenaSeleccion & ")"
        Sql2 = " {rincidencia.codincid} in [" & CadenaSeleccion & "]"
        
        cadSelectBorra = "tmpinformes.campo1 not in (" & CadenaSeleccion & ")"
        
    Else
        Sql = " {rincidencia.codincid} = -1 "
    End If
    If Not AnyadirAFormula(cadSelect, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 20, 21  'Clases
            AbrirFrmClase (Index)
        
        Case 12, 13  'socios
            AbrirFrmSocios (Index)
        
        Case 0, 1 ' Incidencias (plagas)
            AbrirFrmIncidencias (Index)
        
        
    End Select
    PonerFoco txtcodigo(indCodigo)
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
            indice = 6
        Case 1
            indice = 7
        Case 2
            indice = 15
        Case 3, 4
            indice = Index - 1
    End Select

    imgFec(0).Tag = indice '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtcodigo(indice).Text <> "" Then frmC.NovaData = txtcodigo(indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtcodigo(CByte(imgFec(0).Tag)) '<===
    ' ********************************************


End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtcodigo(Index), 3
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
            Case 12: KEYBusqueda KeyAscii, 12 'socio desde
            Case 13: KEYBusqueda KeyAscii, 13 'socio hasta
            Case 20: KEYBusqueda KeyAscii, 20 'clase desde
            Case 21: KEYBusqueda KeyAscii, 21 'clase hasta
            
            Case 0: KEYBusqueda KeyAscii, 20 'plaga desde
            Case 1: KEYBusqueda KeyAscii, 21 'plaga hasta
            
            Case 6: KEYFecha KeyAscii, 0 'fecha desde
            Case 7: KEYFecha KeyAscii, 1 'fecha hasta
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
Dim devuelve As String
Dim codcampo As String, nomCampo As String
Dim Tabla As String
      
    Select Case Index
        'FECHA Desde Hasta
        Case 6, 7
            If txtcodigo(Index).Text <> "" Then
                PonerFormatoFecha txtcodigo(Index)
            End If
        
            
        Case 0, 1 'plagas
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rincidencia", "nomincid", "codincid", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "00")
            
        Case 20, 21
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "clases", "nomclase", "codclase", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000")
        
        
        Case 12, 13  'Cod. Socio
            If PonerFormatoEntero(txtcodigo(Index)) Then
                nomCampo = "nomsocio"
                Tabla = "rsocios"
                codcampo = "codsocio"
                txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), Tabla, nomCampo, codcampo, "N")
                If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
            Else
                txtNombre(Index).Text = ""
            End If
            
    End Select
End Sub



Private Sub PonerFrameFacVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Facturacion de Albaran Visible y Ajustado al Formulario, y visualiza los controles
Dim Cad As String

    H = 5940
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



Private Sub AbrirFrmClase(indice As Integer)
    indCodigo = indice
    Set frmCla = New frmComercial
    
    AyudaClasesCom frmCla, txtcodigo(indice).Text
    
    Set frmCla = Nothing
End Sub

Private Sub AbrirFrmSocios(indice As Integer)
    indCodigo = indice
    Set frmSoc = New frmManSocios
    frmSoc.DatosADevolverBusqueda = "0|1|"
    frmSoc.Show vbModal
    Set frmSoc = Nothing
End Sub


Private Sub AbrirFrmIncidencias(indice As Integer)
    indCodigo = indice
    Set frmInc = New frmManInciden
    frmInc.DatosADevolverBusqueda = "0|1|"
    frmInc.Show vbModal
    Set frmInc = Nothing
End Sub




Private Function ProcesoEntradas(cTabla As String, cWhere As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset

    On Error GoTo eProcesoEntradas
    
    Screen.MousePointer = vbHourglass
    
    ProcesoEntradas = False

    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    cWhere = QuitarCaracterACadena(cWhere, "{")
    cWhere = QuitarCaracterACadena(cWhere, "}")
    
    If Check1.Value = 0 Then
        Sql = "select " & vUsu.Codigo & ",rhisfruta.codsocio,rhisfruta.codvarie,rhisfruta.codcampo,rhisfruta_incidencia.codincid,rhisfruta.fecalbar,rhisfruta.numalbar from " & QuitarCaracterACadena(cTabla, "_1")
        If cWhere <> "" Then
            Sql = Sql & " WHERE " & cWhere
        End If
        
        Sql = Sql & " order by 1, 2, 3, 5, 6 "
                                               'codsocio, codvarie,  codcampo, codincid, fecalbar, numalbar
        Sql2 = "insert into tmpinformes (codusu, importe1,  codigo1, importe2, importe3,  fecha1, importe4  ) "
        Sql2 = Sql2 & Sql
        
        conn.Execute Sql2
    Else
        If Not InsertaPlagasClasAuto(cWhere, cTabla) Then
            ProcesoEntradas = False
            Exit Function
        End If
        Pb1.visible = False
        Label2(24).visible = False
    End If
    
    Screen.MousePointer = vbDefault
    
    ProcesoEntradas = True
    Exit Function
    
eProcesoEntradas:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Proceso de Carga Albaranes Destrios Varios", Err.Description
End Function


Private Function InsertaPlagasClasAuto(vWhere As String, vtabla As String) As Boolean
Dim Sql As String
Dim KilosTot As Long
Dim i As Integer
Dim Porcen As Currency
Dim Rs As ADODB.Recordset
Dim Nregs As Long
    
    
    On Error GoTo eInsertaPlagasClasAuto
    
    InsertaPlagasClasAuto = False

    '[Monica]16/11/2010: las plagas se calculan agrupando las clasificaciones y sacando la media
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    Sql = "select rhisfruta.codsocio, rhisfruta.codvarie, rcampos.nrocampo from "
    Sql = Sql & "(" & vtabla & ") inner join rcampos on rhisfruta.codcampo = rcampos.codcampo "
    Sql = Sql & " where " & vWhere
    Sql = Sql & " group by 1,2,3 "
    Sql = Sql & " order by 1,2,3 "
    
    Nregs = TotalRegistrosConsulta(Sql)
    CargarProgres Pb1, CInt(Nregs)
    Pb1.visible = True
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    While Not Rs.EOF
        IncrementarProgres Pb1, 1
        
        Sql = "select sum(kilosplaga1+kilosplaga2+kilosplaga3+kilosplaga4+kilosplaga5+kilosplaga6+kilosplaga7+kilosplaga8+kilosplaga9+kilosplaga10+kilosplaga11) total "
        Sql = Sql & " from rcontrol_plagas "
        Sql = Sql & " where codsocio = " & DBSet(Rs!Codsocio, "N")
        Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
        Sql = Sql & " and codcampo = " & DBSet(Rs!NroCampo, "N")
        Sql = Sql & " and idplaga <> 2 "
        If txtcodigo(6).Text <> "" Then Sql = Sql & " and fechacla >= " & DBSet(txtcodigo(6).Text, "F")
        If txtcodigo(7).Text <> "" Then Sql = Sql & " and fechacla <= " & DBSet(txtcodigo(7).Text, "F")
        
        KilosTot = DevuelveValor(Sql)


        For i = 3 To 13
            Sql = "SELECT "
            If KilosTot <> 0 Then
                Sql = Sql & " round((sum(kilosplaga1)+sum(kilosplaga2)+sum(kilosplaga3)+sum(kilosplaga4)+sum(kilosplaga5)+sum(kilosplaga6)+sum(kilosplaga7)+sum(kilosplaga8)+sum(kilosplaga9)+sum(kilosplaga10)+sum(kilosplaga11)) * 100 / " & DBSet(KilosTot, "N") & ",2) "
            Else
                Sql = Sql & "0 "
            End If
            Sql = Sql & " from rcontrol_plagas "
            Sql = Sql & " where codsocio = " & DBSet(Rs!Codsocio, "N")
            Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
            Sql = Sql & " and codcampo = " & DBSet(Rs!NroCampo, "N")
            Sql = Sql & " and rcontrol_plagas.idplaga = " & DBSet(i, "N")
            If txtcodigo(6).Text <> "" Then Sql = Sql & " and fechacla >= " & DBSet(txtcodigo(6).Text, "F")
            If txtcodigo(7).Text <> "" Then Sql = Sql & " and fechacla <= " & DBSet(txtcodigo(7).Text, "F")
        
            Porcen = DevuelveValor(Sql)
            
            Select Case i
                Case 3 ' piojo gris
                    Select Case Porcen
                        Case 1 To 5
                            Sql = "select count(*) from tmpinformes "
                            Sql = Sql & " where importe1 = " & DBSet(Rs!Codsocio, "N")
                            Sql = Sql & " and codigo1 = " & DBSet(Rs!codvarie, "N")
                            Sql = Sql & " and importe2 = " & DBSet(Rs!NroCampo, "N")
                            Sql = Sql & " and campo1 = 1"
                            If TotalRegistros(Sql) = 0 Then
                                                                       'codvarie,codsocio, codcampo, codplaga
                                Sql = "insert into tmpinformes (codusu, codigo1, importe1, importe2, campo1) values ( "
                                Sql = Sql & vUsu.Codigo & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Codsocio, "N") & ","
                                Sql = Sql & DBSet(Rs!NroCampo, "N") & ",1) "
                                
                                conn.Execute Sql
                            End If
                            
                        Case 5.01 To 15
                            Sql = "select count(*) from tmpinformes "
                            Sql = Sql & " where importe1 = " & DBSet(Rs!Codsocio, "N")
                            Sql = Sql & " and codigo1 = " & DBSet(Rs!codvarie, "N")
                            Sql = Sql & " and importe2 = " & DBSet(Rs!NroCampo, "N")
                            Sql = Sql & " and campo1 = 2"
                            If TotalRegistros(Sql) = 0 Then
                                                                       'codvarie,codsocio, codcampo, codplaga
                                Sql = "insert into tmpinformes (codusu, codigo1, importe1, importe2, campo1) values ( "
                                Sql = Sql & vUsu.Codigo & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Codsocio, "N") & ","
                                Sql = Sql & DBSet(Rs!NroCampo, "N") & ",2)"
                                
                                conn.Execute Sql
                            End If
            
                        Case Is > 15
                            Sql = "select count(*) from tmpinformes "
                            Sql = Sql & " where importe1 = " & DBSet(Rs!Codsocio, "N")
                            Sql = Sql & " and codigo1 = " & DBSet(Rs!codvarie, "N")
                            Sql = Sql & " and importe2 = " & DBSet(Rs!NroCampo, "N")
                            Sql = Sql & " and campo1 = 3"
                            If TotalRegistros(Sql) = 0 Then
                                                                       'codvarie,codsocio, codcampo, codplaga
                                Sql = "insert into tmpinformes (codusu, codigo1, importe1, importe2, campo1) values ( "
                                Sql = Sql & vUsu.Codigo & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Codsocio, "N") & ","
                                Sql = Sql & DBSet(Rs!NroCampo, "N") & ",3)"
                                
                                conn.Execute Sql
                            End If
                    End Select
                
                
                Case 4 ' piojo rojo
                    Select Case Porcen
                        Case 1 To 5
                            Sql = "select count(*) from tmpinformes "
                            Sql = Sql & " where importe1 = " & DBSet(Rs!Codsocio, "N")
                            Sql = Sql & " and codigo1 = " & DBSet(Rs!codvarie, "N")
                            Sql = Sql & " and importe2 = " & DBSet(Rs!NroCampo, "N")
                            Sql = Sql & " and campo1 = 4"
                            If TotalRegistros(Sql) = 0 Then
                                                                       'codvarie,codsocio, codcampo, codplaga
                                Sql = "insert into tmpinformes (codusu, codigo1, importe1, importe2, campo1) values ( "
                                Sql = Sql & vUsu.Codigo & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Codsocio, "N") & ","
                                Sql = Sql & DBSet(Rs!NroCampo, "N") & ",4)"
                                
                                conn.Execute Sql
                            End If
                            
                        Case 5.01 To 15
                            Sql = "select count(*) from tmpinformes "
                            Sql = Sql & " where importe1 = " & DBSet(Rs!Codsocio, "N")
                            Sql = Sql & " and codigo1 = " & DBSet(Rs!codvarie, "N")
                            Sql = Sql & " and importe2 = " & DBSet(Rs!NroCampo, "N")
                            Sql = Sql & " and campo1 = 5"
                            If TotalRegistros(Sql) = 0 Then
                                                                       'codvarie,codsocio, codcampo, codplaga
                                Sql = "insert into tmpinformes (codusu, codigo1, importe1, importe2, campo1) values ( "
                                Sql = Sql & vUsu.Codigo & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Codsocio, "N") & ","
                                Sql = Sql & DBSet(Rs!NroCampo, "N") & ",5)"
                                
                                conn.Execute Sql
                            End If
            
                        Case Is > 15
                            Sql = "select count(*) from tmpinformes "
                            Sql = Sql & " where importe1 = " & DBSet(Rs!Codsocio, "N")
                            Sql = Sql & " and codigo1 = " & DBSet(Rs!codvarie, "N")
                            Sql = Sql & " and importe2 = " & DBSet(Rs!NroCampo, "N")
                            Sql = Sql & " and campo1 = 6"
                            If TotalRegistros(Sql) = 0 Then
                                                                       'codvarie,codsocio, codcampo, codplaga
                                Sql = "insert into tmpinformes (codusu, codigo1, importe1, importe2, campo1) values ( "
                                Sql = Sql & vUsu.Codigo & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Codsocio, "N") & ","
                                Sql = Sql & DBSet(Rs!NroCampo, "N") & ",6)"
                                
                                conn.Execute Sql
                            End If
                    End Select
                
                
                Case 5 ' serpeta
                    Select Case Porcen
                        Case 1 To 5
                            Sql = "select count(*) from tmpinformes "
                            Sql = Sql & " where importe1 = " & DBSet(Rs!Codsocio, "N")
                            Sql = Sql & " and codigo1 = " & DBSet(Rs!codvarie, "N")
                            Sql = Sql & " and importe2 = " & DBSet(Rs!NroCampo, "N")
                            Sql = Sql & " and campo1 = 7"
                            If TotalRegistros(Sql) = 0 Then
                                                                       'codvarie,codsocio, codcampo, codplaga
                                Sql = "insert into tmpinformes (codusu, codigo1, importe1, importe2, campo1) values ( "
                                Sql = Sql & vUsu.Codigo & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Codsocio, "N") & ","
                                Sql = Sql & DBSet(Rs!NroCampo, "N") & ",7)"
                                
                                conn.Execute Sql
                            End If
                            
                        Case 5.01 To 15
                            Sql = "select count(*) from tmpinformes "
                            Sql = Sql & " where importe1 = " & DBSet(Rs!Codsocio, "N")
                            Sql = Sql & " and codigo1 = " & DBSet(Rs!codvarie, "N")
                            Sql = Sql & " and importe2 = " & DBSet(Rs!NroCampo, "N")
                            Sql = Sql & " and campo1 = 8"
                            If TotalRegistros(Sql) = 0 Then
                                                                       'codvarie,codsocio, codcampo, codplaga
                                Sql = "insert into tmpinformes (codusu, codigo1, importe1, importe2, campo1) values ( "
                                Sql = Sql & vUsu.Codigo & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Codsocio, "N") & ","
                                Sql = Sql & DBSet(Rs!NroCampo, "N") & ",8)"
                                
                                conn.Execute Sql
                            End If
            
                        Case Is > 15
                            Sql = "select count(*) from tmpinformes "
                            Sql = Sql & " where importe1 = " & DBSet(Rs!Codsocio, "N")
                            Sql = Sql & " and codigo1 = " & DBSet(Rs!codvarie, "N")
                            Sql = Sql & " and importe2 = " & DBSet(Rs!NroCampo, "N")
                            Sql = Sql & " and campo1 = 9"
                            If TotalRegistros(Sql) = 0 Then
                                                                       'codvarie,codsocio, codcampo, codplaga
                                Sql = "insert into tmpinformes (codusu, codigo1, importe1, importe2, campo1) values ( "
                                Sql = Sql & vUsu.Codigo & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Codsocio, "N") & ","
                                Sql = Sql & DBSet(Rs!NroCampo, "N") & ",9)"
                                
                                conn.Execute Sql
                            End If
                    End Select
                
                
                Case 6 ' araña
                    Select Case Porcen
                        Case 1 To 5
                            Sql = "select count(*) from tmpinformes "
                            Sql = Sql & " where importe1 = " & DBSet(Rs!Codsocio, "N")
                            Sql = Sql & " and codigo1 = " & DBSet(Rs!codvarie, "N")
                            Sql = Sql & " and importe2 = " & DBSet(Rs!NroCampo, "N")
                            Sql = Sql & " and campo1 = 16"
                            If TotalRegistros(Sql) = 0 Then
                                                                       'codvarie,codsocio, codcampo, codplaga
                                Sql = "insert into tmpinformes (codusu, codigo1, importe1, importe2, campo1) values ( "
                                Sql = Sql & vUsu.Codigo & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Codsocio, "N") & ","
                                Sql = Sql & DBSet(Rs!NroCampo, "N") & ",16)"
                                
                                conn.Execute Sql
                            End If
                            
                        Case 5.01 To 15
                            Sql = "select count(*) from tmpinformes "
                            Sql = Sql & " where importe1 = " & DBSet(Rs!Codsocio, "N")
                            Sql = Sql & " and codigo1 = " & DBSet(Rs!codvarie, "N")
                            Sql = Sql & " and importe2 = " & DBSet(Rs!NroCampo, "N")
                            Sql = Sql & " and campo1 = 17"
                            If TotalRegistros(Sql) = 0 Then
                                                                       'codvarie,codsocio, codcampo, codplaga
                                Sql = "insert into tmpinformes (codusu, codigo1, importe1, importe2, campo1) values ( "
                                Sql = Sql & vUsu.Codigo & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Codsocio, "N") & ","
                                Sql = Sql & DBSet(Rs!NroCampo, "N") & ",17)"
                                
                                conn.Execute Sql
                            End If
            
                        Case Is > 15
                            Sql = "select count(*) from tmpinformes "
                            Sql = Sql & " where importe1 = " & DBSet(Rs!Codsocio, "N")
                            Sql = Sql & " and codigo1 = " & DBSet(Rs!codvarie, "N")
                            Sql = Sql & " and importe2 = " & DBSet(Rs!NroCampo, "N")
                            Sql = Sql & " and campo1 = 18"
                            If TotalRegistros(Sql) = 0 Then
                                                                       'codvarie,codsocio, codcampo, codplaga
                                Sql = "insert into tmpinformes (codusu, codigo1, importe1, importe2, campo1) values ( "
                                Sql = Sql & vUsu.Codigo & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Codsocio, "N") & ","
                                Sql = Sql & DBSet(Rs!NroCampo, "N") & ",18)"
                                
                                conn.Execute Sql
                            End If
                    End Select
                
                
                Case 7 ' %piedra
                    If Porcen > 1 Then
                        Sql = "select count(*) from tmpinformes "
                        Sql = Sql & " where importe1 = " & DBSet(Rs!Codsocio, "N")
                        Sql = Sql & " and codigo1 = " & DBSet(Rs!codvarie, "N")
                        Sql = Sql & " and importe2 = " & DBSet(Rs!NroCampo, "N")
                        Sql = Sql & " and campo1 = 22"
                        If TotalRegistros(Sql) = 0 Then
                                                                   'codvarie,codsocio, codcampo, codplaga
                            Sql = "insert into tmpinformes (codusu, codigo1, importe1, importe2, campo1) values ( "
                            Sql = Sql & vUsu.Codigo & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Codsocio, "N") & ","
                            Sql = Sql & DBSet(Rs!NroCampo, "N") & ",22)"
                            
                            conn.Execute Sql
                        End If
                    End If
                    
                Case 8 ' negrita
                    Select Case Porcen
                        Case 1 To 5
                            Sql = "select count(*) from tmpinformes "
                            Sql = Sql & " where importe1 = " & DBSet(Rs!Codsocio, "N")
                            Sql = Sql & " and codigo1 = " & DBSet(Rs!codvarie, "N")
                            Sql = Sql & " and importe2 = " & DBSet(Rs!NroCampo, "N")
                            Sql = Sql & " and campo1 = 19"
                            If TotalRegistros(Sql) = 0 Then
                                                                       'codvarie,codsocio, codcampo, codplaga
                                Sql = "insert into tmpinformes (codusu, codigo1, importe1, importe2, campo1) values ( "
                                Sql = Sql & vUsu.Codigo & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Codsocio, "N") & ","
                                Sql = Sql & DBSet(Rs!NroCampo, "N") & ",19)"
                                
                                conn.Execute Sql
                            End If
                            
                        Case 5.01 To 15
                            Sql = "select count(*) from tmpinformes "
                            Sql = Sql & " where importe1 = " & DBSet(Rs!Codsocio, "N")
                            Sql = Sql & " and codigo1 = " & DBSet(Rs!codvarie, "N")
                            Sql = Sql & " and importe2 = " & DBSet(Rs!NroCampo, "N")
                            Sql = Sql & " and campo1 = 20"
                            If TotalRegistros(Sql) = 0 Then
                                                                       'codvarie,codsocio, codcampo, codplaga
                                Sql = "insert into tmpinformes (codusu, codigo1, importe1, importe2, campo1) values ( "
                                Sql = Sql & vUsu.Codigo & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Codsocio, "N") & ","
                                Sql = Sql & DBSet(Rs!NroCampo, "N") & ",20)"
                                
                                conn.Execute Sql
                            End If
            
                        Case Is > 15
                            Sql = "select count(*) from tmpinformes "
                            Sql = Sql & " where importe1 = " & DBSet(Rs!Codsocio, "N")
                            Sql = Sql & " and codigo1 = " & DBSet(Rs!codvarie, "N")
                            Sql = Sql & " and importe2 = " & DBSet(Rs!NroCampo, "N")
                            Sql = Sql & " and campo1 = 21"
                            If TotalRegistros(Sql) = 0 Then
                                                                       'codvarie,codsocio, codcampo, codplaga
                                Sql = "insert into tmpinformes (codusu, codigo1, importe1, importe2, campo1) values ( "
                                Sql = Sql & vUsu.Codigo & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Codsocio, "N") & ","
                                Sql = Sql & DBSet(Rs!NroCampo, "N") & ",21)"
                                
                                conn.Execute Sql
                            End If
                    End Select
                
                
                Case 13 ' mosca
                    Select Case Porcen
                        Case 1 To 5
                            Sql = "select count(*) from tmpinformes "
                            Sql = Sql & " where importe1 = " & DBSet(Rs!Codsocio, "N")
                            Sql = Sql & " and codigo1 = " & DBSet(Rs!codvarie, "N")
                            Sql = Sql & " and importe2 = " & DBSet(Rs!NroCampo, "N")
                            Sql = Sql & " and campo1 = 10"
                            If TotalRegistros(Sql) = 0 Then
                                                                       'codvarie,codsocio, codcampo, codplaga
                                Sql = "insert into tmpinformes (codusu, codigo1, importe1, importe2, campo1) values ( "
                                Sql = Sql & vUsu.Codigo & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Codsocio, "N") & ","
                                Sql = Sql & DBSet(Rs!NroCampo, "N") & ",10)"
                                
                                conn.Execute Sql
                            End If
                            
                        Case 5.01 To 15
                            Sql = "select count(*) from tmpinformes "
                            Sql = Sql & " where importe1 = " & DBSet(Rs!Codsocio, "N")
                            Sql = Sql & " and codigo1 = " & DBSet(Rs!codvarie, "N")
                            Sql = Sql & " and importe2 = " & DBSet(Rs!NroCampo, "N")
                            Sql = Sql & " and campo1 = 11"
                            If TotalRegistros(Sql) = 0 Then
                                                                       'codvarie,codsocio, codcampo, codplaga
                                Sql = "insert into tmpinformes (codusu, codigo1, importe1, importe2, campo1) values ( "
                                Sql = Sql & vUsu.Codigo & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Codsocio, "N") & ","
                                Sql = Sql & DBSet(Rs!NroCampo, "N") & ",11)"
                                
                                conn.Execute Sql
                            End If
            
                        Case Is > 15
                            Sql = "select count(*) from tmpinformes "
                            Sql = Sql & " where importe1 = " & DBSet(Rs!Codsocio, "N")
                            Sql = Sql & " and codigo1 = " & DBSet(Rs!codvarie, "N")
                            Sql = Sql & " and importe2 = " & DBSet(Rs!NroCampo, "N")
                            Sql = Sql & " and campo1 = 12"
                            If TotalRegistros(Sql) = 0 Then
                                                                       'codvarie,codsocio, codcampo, codplaga
                                Sql = "insert into tmpinformes (codusu, codigo1, importe1, importe2, campo1) values ( "
                                Sql = Sql & vUsu.Codigo & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Codsocio, "N") & ","
                                Sql = Sql & DBSet(Rs!NroCampo, "N") & ",12)"
                                
                                conn.Execute Sql
                            End If
                    End Select
            End Select
        Next i

        Rs.MoveNext
    Wend
    
    Set Rs = Nothing

    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    If cadSelectBorra <> "" Then Sql = Sql & " and " & cadSelectBorra
    conn.Execute Sql
    
    InsertaPlagasClasAuto = True
    Exit Function


eInsertaPlagasClasAuto:
    MuestraError Err.Number, "Inserta Plagas Clasificacion", Err.Description
End Function
