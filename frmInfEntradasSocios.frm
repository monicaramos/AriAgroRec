VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInfEntradasSocios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6435
   Icon            =   "frmInfEntradasSocios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   6435
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
      Height          =   6150
      Left            =   0
      TabIndex        =   9
      Top             =   -60
      Width           =   6435
      Begin VB.CheckBox Check6 
         Caption         =   "Agrupar por GlobalGap"
         Height          =   225
         Left            =   3270
         TabIndex        =   35
         Top             =   3510
         Width           =   2745
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "Tipo Campo|N|N|0|1|rcampos|tipocampo||N|"
         Top             =   3780
         Width           =   1380
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Incluir Mermas"
         Height          =   225
         Left            =   3270
         TabIndex        =   33
         Top             =   5250
         Width           =   2745
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Incluir Entradas Facturas"
         Height          =   225
         Left            =   3270
         TabIndex        =   32
         Top             =   4890
         Width           =   2745
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Por Tipo de Entrada"
         Height          =   225
         Left            =   3270
         TabIndex        =   31
         Top             =   4560
         Width           =   2745
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Sólo resumen"
         Height          =   225
         Left            =   3270
         TabIndex        =   30
         Top             =   4200
         Width           =   2745
      End
      Begin VB.Frame Frame1 
         Caption         =   "Agrupado por"
         ForeColor       =   &H00972E0B&
         Height          =   915
         Left            =   330
         TabIndex        =   27
         Top             =   4350
         Width           =   2385
         Begin VB.OptionButton Option1 
            Caption         =   "Variedad"
            Height          =   255
            Index           =   1
            Left            =   1290
            TabIndex        =   29
            Top             =   330
            Width           =   945
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Socio"
            Height          =   255
            Index           =   0
            Left            =   210
            TabIndex        =   28
            Top             =   330
            Width           =   825
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Salta página por Socio"
         Height          =   225
         Left            =   3270
         TabIndex        =   26
         Top             =   3840
         Width           =   2745
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
         Left            =   4890
         TabIndex        =   8
         Top             =   5565
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
         TabIndex        =   15
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
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   1455
         Width           =   3195
      End
      Begin VB.CommandButton Command3 
         Height          =   440
         Left            =   7860
         Picture         =   "frmInfEntradasSocios.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command4 
         Height          =   440
         Left            =   7860
         Style           =   1  'Graphical
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
         Text            =   "Text5"
         Top             =   2400
         Width           =   3375
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3810
         TabIndex        =   7
         Top             =   5580
         Width           =   975
      End
      Begin VB.Label Label34 
         Caption         =   "Tipo Campo"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   3750
         Width           =   1245
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1350
         Picture         =   "frmInfEntradasSocios.frx":0097
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
         TabIndex        =   25
         Top             =   2700
         Width           =   450
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   20
         Left            =   735
         TabIndex        =   24
         Top             =   2940
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   21
         Left            =   735
         TabIndex        =   23
         Top             =   3285
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   22
         Left            =   720
         TabIndex        =   22
         Top             =   1140
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   23
         Left            =   720
         TabIndex        =   21
         Top             =   1500
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Informe de Entradas"
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   900
         Width           =   405
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   1350
         MouseIcon       =   "frmInfEntradasSocios.frx":0122
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
         Picture         =   "frmInfEntradasSocios.frx":0126
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
         TabIndex        =   18
         Top             =   1830
         Width           =   390
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   18
         Left            =   765
         TabIndex        =   17
         Top             =   2085
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   28
         Left            =   765
         TabIndex        =   16
         Top             =   2475
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   20
         Left            =   1350
         MouseIcon       =   "frmInfEntradasSocios.frx":01B1
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   21
         Left            =   1350
         MouseIcon       =   "frmInfEntradasSocios.frx":0303
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2430
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmInfEntradasSocios"
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
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1


'Private HaDevueltoDatos As Boolean
Private NomTabla As String
Private NomTablaLin As String

Private WithEvents frmSoc As frmManSocios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmMens8 As frmMensajes
Attribute frmMens8.VB_VarHelpID = -1

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


Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String

Dim Indice As Integer

Dim indCodigo As Integer 'indice para txtCodigo

Dim PrimeraVez As Boolean

Dim vSeccion As CSeccion
Dim Contratos As String

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub

Private Function DatosOK() As Boolean
Dim B As Boolean
    B = True
    
    DatosOK = B

End Function

Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Check2_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Check3_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Check4_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Check5_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

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
Dim nTabla2 As String


Dim Nregs As Long
Dim FecFac As Date
Dim TipoPrec As Byte ' 0 anticipos
                     ' 1 liquidaciones
Dim B As Boolean
Dim Sql2 As String
Dim vcad As String


    InicializarVbles
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If DatosOK Then
        '======== FORMULA  ====================================
        'D/H Socios
        cDesde = Trim(txtcodigo(12).Text)
        cHasta = Trim(txtcodigo(13).Text)
        nDesde = txtNombre(12).Text
        nHasta = txtNombre(13).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{rclasifica.codsocio}"
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
            Codigo = "{rclasifica.fechaent}"
            TipCod = "F"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
        End If
        
        
        If vParamAplic.Cooperativa = 16 Then
            If Combo1(0).ListIndex = 0 Then
                ' todos
            Else
                If Not AnyadirAFormula(cadSelect, "rcampos.codigoggap = " & Combo1(0).ListIndex) Then Exit Sub
            End If
        
        Else
            Select Case Combo1(0).ListIndex
                Case 0 ' todos
                                
                Case 1
                    If Not AnyadirAFormula(cadSelect, "rcampos.codigoggap >= '1' and not rcampos.codigoggap is null") Then Exit Sub
                Case 2
                    If Not AnyadirAFormula(cadSelect, "rcampos.esnaturane = 1") Then Exit Sub
            End Select
        End If
        
        nTabla = "((((rclasifica "
        nTabla = nTabla & " INNER JOIN rsocios ON rclasifica.codsocio = rsocios.codsocio) "
        nTabla = nTabla & " INNER JOIN variedades ON rclasifica.codvarie = variedades.codvarie) "
        nTabla = nTabla & " INNER JOIN productos ON variedades.codprodu = productos.codprodu) "
        '[Monica]22/12/2016: añadido el combo de tipo de campo (todos, globalgap o naturane)
        nTabla = nTabla & " INNER JOIN rcampos ON rclasifica.codcampo = rcampos.codcampo) "
        nTabla = nTabla & " INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        
        '[Monica]03/11/2011: en Quatretonda podemos mostrar las variedades de almazara
        If vParamAplic.Cooperativa <> 7 Then
            nTabla = nTabla & " and grupopro.codgrupo <> 5 and grupopro.codgrupo <> 6 " ' grupo no puede ser 5=almazara ni 6=bodega
        Else
            nTabla = nTabla & " and grupopro.codgrupo <> 6 " ' grupo no puede ser 6=bodega
        End If
        
        
        nTabla2 = "((((rentradas "
        nTabla2 = nTabla2 & " INNER JOIN rsocios ON rentradas.codsocio = rsocios.codsocio) "
        nTabla2 = nTabla2 & " INNER JOIN variedades ON rentradas.codvarie = variedades.codvarie) "
        nTabla2 = nTabla2 & " INNER JOIN productos ON variedades.codprodu = productos.codprodu) "
        '[Monica]22/12/2016: añadido el combo de tipo de campo (todos, globalgap o naturane)
        nTabla2 = nTabla2 & " INNER JOIN rcampos ON rentradas.codcampo = rcampos.codcampo) "
        nTabla2 = nTabla2 & " INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        
        '[Monica]03/11/2011: en Quatretonda podemos mostrar las variedades de almazara
        If vParamAplic.Cooperativa <> 7 Then
            nTabla2 = nTabla2 & " and grupopro.codgrupo <> 5 and grupopro.codgrupo <> 6 " ' grupo no puede ser 5=almazara ni 6=bodega
        Else
            nTabla2 = nTabla2 & " and grupopro.codgrupo <> 6 " ' grupo no puede ser 6=bodega
        End If
        
        cadSelect2 = Replace(cadSelect, "rclasifica", "rentradas")
        
'        If Not AnyadirAFormula(cadSelect, "{rclasifica.transportadopor} = 0 ") Then Exit Sub
'        If Not AnyadirAFormula(cadFormula, "{rclasifica.transportadopor} = 0 ") Then Exit Sub
'
'        If Not AnyadirAFormula(cadSelect, "{rclasifica.tipoentr} <> 1 ") Then Exit Sub
'        If Not AnyadirAFormula(cadFormula, "{rclasifica.tipoentr} <> 1 ") Then Exit Sub
        
        cadSelect1 = Replace(Replace(cadSelect, "rclasifica", "rhisfruta_entradas"), "rhisfruta_entradas.codsocio", "rhisfruta.codsocio")
        
        Tabla1 = "(((((rhisfruta INNER JOIN rhisfruta_entradas on rhisfruta.numalbar = rhisfruta_entradas.numalbar) "
        Tabla1 = Tabla1 & " INNER JOIN rsocios ON rhisfruta.codsocio = rsocios.codsocio) "
        Tabla1 = Tabla1 & " INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie) "
        Tabla1 = Tabla1 & " INNER JOIN productos ON variedades.codprodu = productos.codprodu) "
        '[Monica]22/12/2016: añadido el combo de tipo de campo (todos, globalgap o naturane)
        Tabla1 = Tabla1 & " INNER JOIN rcampos ON rhisfruta.codcampo = rcampos.codcampo) "
        Tabla1 = Tabla1 & " INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        '[Monica]03/11/2011: en Quatretonda podemos mostrar las variedades de almazara
        If vParamAplic.Cooperativa <> 7 Then
            Tabla1 = Tabla1 & " and grupopro.codgrupo <> 5 and grupopro.codgrupo <> 6 " ' grupo no puede ser 5=almazara ni 6=bodega
        Else
            Tabla1 = Tabla1 & " and grupopro.codgrupo <> 6 " ' grupo no puede ser 6=bodega
        End If
        
        
        cadDesde = "01/01/1900"
        cadhasta = "31/12/2500"
        
        If txtcodigo(6).Text <> "" Then cadDesde = CDate(txtcodigo(6).Text)
        If txtcodigo(7).Text <> "" Then cadhasta = CDate(txtcodigo(7).Text)
        
        CadParam = CadParam & "pFecDesde= Date(" & Year(cadDesde) & "," & Month(cadDesde) & "," & Day(cadDesde) & ")" & "|" 'txtcodigo(6).Text & """|"
        CadParam = CadParam & "pFecHasta= Date(" & Year(cadhasta) & "," & Month(cadhasta) & "," & Day(cadhasta) & ")" & "|" 'txtcodigo(7).Text & """|"
        numParam = numParam + 2
        
        
        CadParam = CadParam & "pSaltoSocio=" & Check1.Value & "|"
        numParam = numParam + 1
        
        If Me.Option1(0).Value Then
            cadTitulo = "Informe Entradas por Socio"
        Else
            cadTitulo = "Informe Entradas por Variedad"
        End If
                
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 16
        frmMens.cadWHERE = Sql2
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        
        
        '[Monica]30/12/2016: incluimos los contratos para poder seleccionar
        Contratos = ""
        If vParamAplic.Cooperativa = 16 Then
        
           Set frmMens8 = New frmMensajes
           
           frmMens8.desdeHco = True
           frmMens8.OpcionMensaje = 64
           frmMens8.Show vbModal
           
           Set frmMens8 = Nothing

            If Contratos <> "" Then
                ' rentradas
                If InStr(UCase(Contratos), "NULL") <> 0 Then
                    vcad = "(rentradas.contrato is null or rentradas.contrato in (" & Contratos & "))"
                Else
                    vcad = "(rentradas.contrato in (" & Contratos & "))"
                End If
                If Not AnyadirAFormula(cadSelect2, vcad) Then Exit Sub
            
                ' rclasifica
                If InStr(UCase(Contratos), "NULL") <> 0 Then
                    vcad = "(rclasifica.contrato is null or rclasifica.contrato in (" & Contratos & "))"
                Else
                    vcad = "(rclasifica.contrato in (" & Contratos & "))"
                End If
                If Not AnyadirAFormula(cadSelect, vcad) Then Exit Sub
                
                ' rhsifruta
                If InStr(UCase(Contratos), "NULL") <> 0 Then
                    vcad = "(rhisfruta.contrato is null or rhisfruta.contrato in (" & Contratos & "))"
                Else
                    vcad = "(rhisfruta.contrato in (" & Contratos & "))"
                End If
                If Not AnyadirAFormula(cadSelect1, vcad) Then Exit Sub
            Else
                ' rentradas
                vcad = "rentradas.contrato = '-1'"
                If Not AnyadirAFormula(cadSelect2, vcad) Then Exit Sub
                ' rclasifica
                vcad = "rclasifica.contrato = '-1'"
                If Not AnyadirAFormula(cadSelect, vcad) Then Exit Sub
                ' rhsifruta
                vcad = "rhisfruta.contrato = '-1'"
                If Not AnyadirAFormula(cadSelect1, vcad) Then Exit Sub
            End If
        End If
        
        conSubRPT = True
        
        If ProcesoEntradasSocio(nTabla, cadSelect, Tabla1, cadSelect1, nTabla2, cadSelect2) Then
            If HayRegParaInforme("tmpinformes", "{tmpinformes.codusu} = " & vUsu.Codigo) Then
                cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo

                'Nombre fichero .rpt a Imprimir
                indRPT = 59 ' informe de entradas por transportista
                
                If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                
                If Me.Option1(1).Value Then
                    If Check3.Value Then
                        nomDocu = Replace(nomDocu, "EntradasSocios.rpt", "EntradasSocios2.rpt")
                        '[Monica]26/08/2011: Resumen por variedad
                        If Check2.Value Then
                            CadParam = CadParam & "pResumen=1|"
                            numParam = numParam + 1
                        End If
                    Else
                        nomDocu = Replace(nomDocu, "EntradasSocios.rpt", "EntradasSocios1.rpt")
                        '[Monica]26/08/2011: Resumen por variedad
                        If Check2.Value Then
                            CadParam = CadParam & "pResumen=1|"
                            numParam = numParam + 1
                        End If
                    End If
                    '[Monica]26/08/2014: calculo con mermas para Quatretonda
                    If vParamAplic.Cooperativa = 7 Then
                        CadParam = CadParam & "pMerma=" & Check5.Value & "|"
                        numParam = numParam + 1
                    End If
                End If
                
                '[Monica]18/01/2012: permitimos seleccionar el solo resumen cuando es por socio
                '                    añado el siguiente if
                If Me.Option1(0).Value Then
                    If Check2.Value Then
                        nomDocu = Replace(nomDocu, "EntradasSocios.rpt", "EntradasSocios3.rpt")
                        CadParam = CadParam & "pResumen=1|"
                        numParam = numParam + 1
                    End If
                End If
                
                frmImprimir.NombreRPT = nomDocu
                
                cadNombreRPT = nomDocu '"rInfEntradasTrans.rpt"
                
                ConSubInforme = True
                LlamarImprimir
'            Else
'                MsgBox "No hay registros entre esos límites.", vbExclamation
            End If
        End If
        
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub Combo1_Click(Index As Integer)
    If Index = 0 Then
        Check6.Enabled = (Combo1(0).ListIndex = 1) And vParamAplic.Cooperativa = 0
        If Not Check6.Enabled Then Check6.Value = 0
    End If
End Sub

Private Sub Combo1_Validate(Index As Integer, Cancel As Boolean)
    If Index = 0 Then
        Check6.Enabled = (Combo1(0).ListIndex = 1) And vParamAplic.Cooperativa = 0
        If Not Check6.Enabled Then Check6.Value = 0
    End If
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtcodigo(12)
        Combo1(0).ListIndex = 0
        Check6.Enabled = (vParamAplic.Cooperativa = 0 And Combo1(0).ListIndex = 1)
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim I As Integer
Dim indFrame As Single

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me

    'Ocultar todos los Frames de Formulario
    Me.FrameFacturar.visible = False
    
    
    For I = 12 To 13
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    For I = 20 To 21
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I

    
    NomTabla = "rhisfruta"
    NomTablaLin = "rhisfruta_entradas"
        
    PonerFrameFacVisible True, H, W
    txtcodigo(7).Text = Format(Now - 1, "dd/mm/yyyy")
    indFrame = 6
    
    '[Monica]28/10/2014: Incluyen o no mermas en el listado (solo Quatretonda)
    Check5.Enabled = (vParamAplic.Cooperativa = 7)
    Check5.visible = (vParamAplic.Cooperativa = 7)
    
    
    
    Me.Option1(0).Value = True
    Option1_Click (0)
    
    CargaCombo
    
    '[Monica]14/03/2017: solo en el caso de catadau agrupamos por codigo productor (codigoggap)
    Check6.Enabled = (vParamAplic.Cooperativa = 0 And Combo1(0).ListIndex = 1)
    Check6.visible = (vParamAplic.Cooperativa = 0)
    
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
        
End Sub

Private Sub CargaCombo()
Dim I As Integer
Dim Rs As ADODB.Recordset
Dim SQL As String

   ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For I = 0 To Combo1.Count - 1
        Combo1(I).Clear
    Next I
    
    If vParamAplic.Cooperativa = 16 Then
        Combo1(0).AddItem "Todos"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 0
        
        SQL = "select * from rglobalgap order by codigo"
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs.EOF
            Combo1(0).ItemData(Combo1(0).NewIndex) = DBLet(Rs!Codigo, "N")
            Combo1(0).AddItem DBLet(Rs!Descripcion)
                    
            Rs.MoveNext
        Wend
        Set Rs = Nothing
    
    Else
        'tipo de campo
        Combo1(0).AddItem "Todos"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 0
        Combo1(0).AddItem "Globalgap"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 1
        Combo1(0).AddItem "Naturane"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    End If
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
    If Not AnyadirAFormula(cadSelect2, SQL) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub



Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtcodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub



Private Sub imgFecha_Click(Index As Integer)
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


Private Sub frmMens8_datoseleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        Contratos = CadenaSeleccion
    Else
        Contratos = ""
    End If
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
    If txtcodigo(Indice).Text <> "" Then frmC.NovaData = txtcodigo(Indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtcodigo(CByte(imgFec(0).Tag)) '<===
    ' ********************************************


End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0 ' por socio
            Label3.Caption = "Informe de Entradas por Socio"
            Check1.Enabled = True
'[Monica]18/01/2012: dejamos seleccionar el solo resumen
'            Check2.Value = 0
'            Check2.Enabled = False
            Check3.Value = 0
            Check3.Enabled = False
            '[Monica]28/10/2014: para el cálculo seleccionar o no mermas (Quatretonda)
            Check5.Enabled = False
        
        Case 1 ' por variedad
            Label3.Caption = "Informe de Entradas por Variedad"
            Check1.Value = 0
            Check1.Enabled = False
            Check2.Enabled = True
            Check3.Enabled = True
            
            '[Monica]28/10/2014: para el cálculo seleccionar o no mermas (Quatretonda)
            Check5.Enabled = (vParamAplic.Cooperativa = 7)
            
    End Select
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtcodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 12: KEYBusqueda KeyAscii, 12 'socio desde
            Case 13: KEYBusqueda KeyAscii, 13 'socio hasta
            Case 20: KEYBusqueda KeyAscii, 20 'clase desde
            Case 21: KEYBusqueda KeyAscii, 21 'clase hasta
            
            Case 6: KEYFecha KeyAscii, 0 'fecha desde
            Case 7: KEYFecha KeyAscii, 1 'fecha hasta
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
Dim codcampo As String, nomCampo As String
Dim tabla As String
      
    Select Case Index
        'FECHA Desde Hasta
        Case 6, 7
            If txtcodigo(Index).Text <> "" Then
                PonerFormatoFecha txtcodigo(Index)
            End If
        
            
        Case 20, 21
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "clases", "nomclase", "codclase", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000")
        
        
        Case 12, 13  'Cod. Socio
            If PonerFormatoEntero(txtcodigo(Index)) Then
                nomCampo = "nomsocio"
                tabla = "rsocios"
                codcampo = "codsocio"
                txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), tabla, nomCampo, codcampo, "N")
                If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
            Else
                txtNombre(Index).Text = ""
            End If
            
    End Select
End Sub



Private Sub PonerFrameFacVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Facturacion de Albaran Visible y Ajustado al Formulario, y visualiza los controles
Dim cad As String

    H = 6150 '5550
    W = 6435
    
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


Private Sub AbrirFrmClase(Indice As Integer)
    indCodigo = Indice
    Set frmCla = New frmComercial
    
    AyudaClasesCom frmCla, txtcodigo(Indice).Text
    
    Set frmCla = Nothing
End Sub

Private Sub AbrirFrmSocios(Indice As Integer)
    indCodigo = Indice
    Set frmSoc = New frmManSocios
    frmSoc.DatosADevolverBusqueda = "0|1|"
    frmSoc.Show vbModal
    Set frmSoc = Nothing
End Sub


Private Function ProcesoEntradasSocio(cTabla As String, cWhere As String, ctabla1 As String, cwhere1 As String, cTabla2 As String, cWhere2 As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset

    On Error GoTo eProcesoEntradasSocio
    
    Screen.MousePointer = vbHourglass
    
    ProcesoEntradasSocio = False

    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    ctabla1 = QuitarCaracterACadena(ctabla1, "{")
    ctabla1 = QuitarCaracterACadena(ctabla1, "}")
    
    cTabla2 = QuitarCaracterACadena(cTabla2, "{")
    cTabla2 = QuitarCaracterACadena(cTabla2, "}")
    
    
    cWhere = QuitarCaracterACadena(cWhere, "{")
    cWhere = QuitarCaracterACadena(cWhere, "}")
    
    cwhere1 = QuitarCaracterACadena(cwhere1, "{")
    cwhere1 = QuitarCaracterACadena(cwhere1, "}")
    
    cWhere2 = QuitarCaracterACadena(cWhere2, "{")
    cWhere2 = QuitarCaracterACadena(cWhere2, "}")
    
    
    SQL = "select " & vUsu.Codigo & ", rentradas.codsocio,rentradas.codvarie,rentradas.codcampo,rentradas.numnotac,rentradas.fechaent,rentradas.kilosnet,rentradas.kilostra,rentradas.kilosbru,rentradas.recolect, rentradas.tipoentr, "
    '[Monica]19/04/2013: añadidas las cajas por Montifrut (aunque solo tienen registros en rhisfruta)
    SQL = SQL & " (if(numcajo1 is null, 0,numcajo1) + if(numcajo2 is null, 0,numcajo2) + if(numcajo3 is null, 0,numcajo3) + if(numcajo4 is null, 0,numcajo4) + if(numcajo5 is null, 0,numcajo5)) numcajon "
    
    '[Monica]14/03/2017: añadido el codigo de ggap por catadau para el caso en que se quiere agrupar por productor(ggap)
    If Check6.Value Then
        SQL = SQL & ", rcampos.codigoggap "
    Else
        SQL = SQL & "," & ValorNulo
    End If
    
    SQL = SQL & " from " & QuitarCaracterACadena(cTabla2, "_1")
    If cWhere2 <> "" Then
        SQL = SQL & " WHERE " & cWhere2
    End If
    SQL = SQL & " union "
    SQL = SQL & "select " & vUsu.Codigo & ", rclasifica.codsocio,rclasifica.codvarie,rclasifica.codcampo,rclasifica.numnotac,rclasifica.fechaent,rclasifica.kilosnet,rclasifica.kilostra, rclasifica.kilosbru, rclasifica.recolect, rclasifica.tipoentr, rclasifica.numcajon "
    '[Monica]14/03/2017: añadido el codigo de ggap por catadau para el caso en que se quiere agrupar por productor(ggap)
    If Check6.Value Then
        SQL = SQL & ", rcampos.codigoggap "
    Else
        SQL = SQL & "," & ValorNulo
    End If
    
    SQL = SQL & " from " & QuitarCaracterACadena(cTabla, "_1")
    
    
    If cWhere <> "" Then
        SQL = SQL & " WHERE " & cWhere
    End If
    SQL = SQL & " union "
    SQL = SQL & "select " & vUsu.Codigo & ", rhisfruta.codsocio,rhisfruta.codvarie,rhisfruta.codcampo,rhisfruta_entradas.numnotac,rhisfruta_entradas.fechaent,rhisfruta_entradas.kilosnet,rhisfruta_entradas.kilostra,rhisfruta_entradas.kilosbru, rhisfruta.recolect, rhisfruta.tipoentr, rhisfruta.numcajon "
    
    '[Monica]14/03/2017: añadido el codigo de ggap por catadau para el caso en que se quiere agrupar por productor(ggap)
    If Check6.Value Then
        SQL = SQL & ", rcampos.codigoggap "
    Else
        SQL = SQL & "," & ValorNulo
    End If
    
    SQL = SQL & " from " & QuitarCaracterACadena(ctabla1, "_1")
    
    
    If cwhere1 <> "" Then
        SQL = SQL & " WHERE " & cwhere1
    End If
    
    '[Monica]03/05/2013: incluimos las entradas que sean de las facturas de siniestro
    If Check4.Value = 1 Then
        SQL = SQL & " union "
        SQL = SQL & "select " & vUsu.Codigo & ", rhisfrutasin.codsocio,rhisfrutasin.codvarie,rhisfrutasin.codcampo,rhisfrutasin_entradas.numnotac,rhisfrutasin_entradas.fechaent,rhisfrutasin_entradas.kilosnet,rhisfrutasin_entradas.kilostra,rhisfrutasin_entradas.kilosbru, rhisfrutasin.recolect, rhisfrutasin.tipoentr, rhisfrutasin.numcajon "
        
        '[Monica]14/03/2017: añadido el codigo de ggap por catadau para el caso en que se quiere agrupar por productor(ggap)
        If Check6.Value Then
            SQL = SQL & ", rcampos.codigoggap "
        Else
            SQL = SQL & "," & ValorNulo
        End If
        
        SQL = SQL & " from " & Replace(QuitarCaracterACadena(ctabla1, "_1"), "rhisfruta", "rhisfrutasin")
        
        
        If cwhere1 <> "" Then
            SQL = SQL & " WHERE " & Replace(cwhere1, "rhisfruta", "rhisfrutasin")
        End If
    End If
    
    
    SQL = SQL & " order by 1, 2, 3 "
                                           'codsocio, codvarie,  codcampo, numnotac, fechaent, kilosnet,  kilostra, kilosbru,  recolect,  tipoentr,  numcajon  codigoggap
    Sql2 = "insert into tmpinformes (codusu, importe1,  codigo1, importe2, importeb2, fecha1,  importe3,  importe4, importeb1, importeb3, importeb4, importeb5, nombre1) "
    Sql2 = Sql2 & SQL
    
    conn.Execute Sql2
    
    
    '[Monica]28/10/2014: añado las mermas en la columna importe5
    If Check5.Value = 1 Then
    
        '1º Taras = PB - (PN / 0.96)
        '2º merma = (PB - Taras) * 0.04
        '3º resultado = PN + merma
    
        SQL = "update tmpinformes tt, variedades vv set "
        SQL = SQL & " tt.importe3 =  tt.importe3 + round((tt.importeb1 - round(tt.importeb1 - (tt.importe3 / round(1-(vv.porcmerm / 100),2)),0)) * round(vv.porcmerm / 100,2) ,0) "
        SQL = SQL & " where tt.codusu = " & vUsu.Codigo
        SQL = SQL & " and tt.codigo1 = vv.codvarie "
        
        conn.Execute SQL
    End If
    
    Screen.MousePointer = vbDefault
    
    ProcesoEntradasSocio = True
    Exit Function
    
eProcesoEntradasSocio:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Proceso de Entradas Socio", Err.Description
End Function

