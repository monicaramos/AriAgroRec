VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDocAltaBaja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6345
   Icon            =   "frmDocAltaBaja.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCalidades 
      Height          =   5940
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   6165
      Begin VB.Frame Frame1 
         ForeColor       =   &H00972E0B&
         Height          =   660
         Left            =   450
         TabIndex        =   7
         Top             =   960
         Width           =   3000
         Begin VB.OptionButton Opcion1 
            Caption         =   "Transmisi�n"
            Height          =   255
            Index           =   2
            Left            =   1680
            TabIndex        =   36
            Top             =   270
            Width           =   1155
         End
         Begin VB.OptionButton Opcion1 
            Caption         =   "Baja"
            Height          =   255
            Index           =   1
            Left            =   870
            TabIndex        =   35
            Top             =   270
            Width           =   975
         End
         Begin VB.OptionButton Opcion1 
            Caption         =   "Alta"
            Height          =   345
            Index           =   0
            Left            =   150
            TabIndex        =   34
            Top             =   225
            Width           =   1290
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   4650
         TabIndex        =   5
         Top             =   5310
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   2
         Left            =   3570
         TabIndex        =   4
         Top             =   5295
         Width           =   975
      End
      Begin VB.CommandButton Command8 
         Height          =   440
         Left            =   7860
         Picture         =   "frmDocAltaBaja.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command7 
         Height          =   440
         Left            =   7860
         Picture         =   "frmDocAltaBaja.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.Frame Frame2 
         ForeColor       =   &H00972E0B&
         Height          =   660
         Left            =   3570
         TabIndex        =   1
         Top             =   960
         Width           =   2235
         Begin VB.OptionButton Opcion 
            Caption         =   "Socio"
            Height          =   345
            Index           =   0
            Left            =   270
            TabIndex        =   32
            Top             =   225
            Width           =   885
         End
         Begin VB.OptionButton Opcion 
            Caption         =   "Campo"
            Height          =   255
            Index           =   1
            Left            =   1170
            TabIndex        =   33
            Top             =   270
            Width           =   885
         End
      End
      Begin VB.Frame FrameAltaSocio 
         Caption         =   "Alta de Socios"
         ForeColor       =   &H00972E0B&
         Height          =   3180
         Left            =   495
         TabIndex        =   17
         Top             =   1800
         Width           =   5340
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   5
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   19
            Top             =   900
            Width           =   1035
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   18
            Top             =   405
            Width           =   1050
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Cargo en Banco"
            ForeColor       =   &H00972E0B&
            Height          =   330
            Index           =   0
            Left            =   270
            TabIndex        =   21
            Top             =   2610
            Width           =   2310
         End
         Begin VB.TextBox txtCodigo 
            Height          =   1095
            Index           =   3
            Left            =   1800
            MaxLength       =   300
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
            Top             =   1395
            Width           =   3225
         End
         Begin VB.Label Label2 
            Caption         =   "Cuota de Entrada"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   3
            Left            =   315
            TabIndex        =   24
            Top             =   945
            Width           =   1545
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Carga"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   0
            Left            =   315
            TabIndex        =   23
            Top             =   420
            Width           =   960
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   1
            Left            =   1350
            Picture         =   "frmDocAltaBaja.frx":0620
            ToolTipText     =   "Buscar fecha"
            Top             =   405
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Observaciones"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   2
            Left            =   315
            TabIndex        =   22
            Top             =   1440
            Width           =   1545
         End
      End
      Begin VB.Frame FrameAltaCampo 
         Caption         =   "Alta de Campos"
         ForeColor       =   &H00972E0B&
         Height          =   3180
         Left            =   480
         TabIndex        =   8
         Top             =   1800
         Width           =   5340
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   9
            Left            =   1620
            MaxLength       =   10
            TabIndex        =   10
            Top             =   1305
            Width           =   1050
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   1665
            MaxLength       =   10
            TabIndex        =   9
            Top             =   405
            Width           =   1050
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Acuerdo Consejo Rector"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   2
            Left            =   315
            TabIndex        =   25
            Top             =   960
            Width           =   3075
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   3
            Left            =   1350
            Picture         =   "frmDocAltaBaja.frx":06AB
            ToolTipText     =   "Buscar fecha"
            Top             =   1305
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   0
            Left            =   1395
            Picture         =   "frmDocAltaBaja.frx":0736
            ToolTipText     =   "Buscar fecha"
            Top             =   405
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Carga"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   16
            Left            =   315
            TabIndex        =   11
            Top             =   420
            Width           =   960
         End
      End
      Begin VB.Frame FrameTransmision 
         Caption         =   "Transmisi�n de Campos"
         ForeColor       =   &H00972E0B&
         Height          =   3180
         Left            =   450
         TabIndex        =   26
         Top             =   1800
         Width           =   5370
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   1410
            MaxLength       =   40
            TabIndex        =   31
            Top             =   1260
            Width           =   3705
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   1410
            MaxLength       =   10
            TabIndex        =   27
            Top             =   405
            Width           =   1050
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   330
            MaxLength       =   10
            TabIndex        =   28
            Top             =   1260
            Width           =   1050
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   0
            Left            =   1110
            ToolTipText     =   "Buscar Socio"
            Top             =   960
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   4
            Left            =   315
            TabIndex        =   30
            Top             =   420
            Width           =   570
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   4
            Left            =   1110
            Picture         =   "frmDocAltaBaja.frx":07C1
            ToolTipText     =   "Buscar fecha"
            Top             =   420
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Receptor"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   3
            Left            =   315
            TabIndex        =   29
            Top             =   960
            Width           =   735
         End
      End
      Begin VB.Frame FrameBajaSocio 
         Caption         =   "Baja de Socios"
         ForeColor       =   &H00972E0B&
         Height          =   3180
         Left            =   495
         TabIndex        =   12
         Top             =   1800
         Width           =   5340
         Begin VB.TextBox txtCodigo 
            Height          =   1950
            Index           =   8
            Left            =   1800
            MaxLength       =   800
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   855
            Width           =   3225
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   7
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   13
            Top             =   405
            Width           =   1050
         End
         Begin VB.Label Label2 
            Caption         =   "Motivos de Baja"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   5
            Left            =   270
            TabIndex        =   16
            Top             =   900
            Width           =   1545
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   2
            Left            =   1485
            Picture         =   "frmDocAltaBaja.frx":084C
            ToolTipText     =   "Buscar fecha"
            Top             =   405
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Baja"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   1
            Left            =   315
            TabIndex        =   15
            Top             =   420
            Width           =   960
         End
      End
      Begin VB.Label Label8 
         Caption         =   "Documentos Alta Baja Transmisi�n"
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
         Left            =   570
         TabIndex        =   6
         Top             =   450
         Width           =   5025
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5895
      Top             =   5265
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmDocAltaBaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: LAURA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public OpcionListado As Byte
    '==== Listados BASICOS ====
    '=============================
    ' 10 .- Listado de Clientes
    ' 11 .- Listado de Proveedores
    ' 12 .- Listado de Calidades
    ' 13 .- Listado de Socios por Secci�n
    ' 15 .- Listado de Horas trababajadas
    
Public NumCod As String 'Para indicar codigo de socio

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(Cliente As String, Observaciones As String)

Private Conexion As Byte
'1.- Conexi�n a BD Ariges  2.- Conexi�n a BD Conta

Private HaDevueltoDatos As Boolean


Private WithEvents frmSec As frmManSeccion 'Secciones
Attribute frmSec.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'Socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmVar As frmComVar 'Variedades
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmCal As frmManCalidades 'Calidades
Attribute frmCal.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'mensajes
Attribute frmMens.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe


Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'n� de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Tabla As String
Dim Codigo As String 'C�digo para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String


Dim PrimeraVez As Boolean
Dim Contabilizada As Byte
Dim indice As Integer

Dim NumCampo As String

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me

End Sub

Private Sub cmdAceptar_Click(Index As Integer)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
    
    
    InicializarVbles
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
'    cadParam = cadParam & "pCodigo=" & NumCod & "|"
'    numParam = numParam + 1
    
    'alta de socios
    If Opcion(0).Value And Opcion1(0).Value Then
        cadTitulo = "Documento Alta de Socios"
        
        If txtcodigo(4).Text <> "" Then
            CadParam = CadParam & "pFecha=""" & txtcodigo(4).Text & """|"
            numParam = numParam + 1
        End If
        
        CadParam = CadParam & "pImporte=""" & txtcodigo(5).Text & """|"
        numParam = numParam + 1
        
        CadParam = CadParam & "pBanco=" & Check1(0).Value & "|"
        numParam = numParam + 1
        
        CadParam = CadParam & "pObserva=""" & txtcodigo(3).Text & """|"
        numParam = numParam + 1
        
        '[Monica]13/03/2014: distinguimos entre escalona y utxera y el resto
        If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
'[Monica]04/06/2014: no tiene pq tener campos
'            If Not AnyadirAFormula(cadFormula, "{rcampos.codpropiet} = " & NumCod & " and isnull({rcampos.fecbajas}) ") Then Exit Sub
'            If Not AnyadirAFormula(cadSelect, "rcampos.codpropiet = " & NumCod & " and rcampos.fecbajas is null ") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{rsocios.codsocio} = " & NumCod) Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "rsocios.codsocio = " & NumCod) Then Exit Sub
        
        
        Else
            If Not AnyadirAFormula(cadFormula, "{rcampos.codsocio} = " & NumCod & " and isnull({rcampos.fecbajas}) ") Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "rcampos.codsocio = " & NumCod & " and rcampos.fecbajas is null ") Then Exit Sub
        End If
    End If
    
    'alta de campos
    If Opcion(1).Value And Opcion1(0).Value Then
        cadTitulo = "Documento Alta de Campos"
    
        CadParam = CadParam & "pFecha=""" & txtcodigo(2).Text & """|"
        numParam = numParam + 1
    
        CadParam = CadParam & "pFechaCons=""" & txtcodigo(9).Text & """|"
        numParam = numParam + 1
    
        If Not AnyadirAFormula(cadFormula, "{rcampos.codsocio} = " & NumCod & " and isnull({rcampos.fecbajas}) ") Then Exit Sub
        If Not AnyadirAFormula(cadSelect, "rcampos.codsocio = " & NumCod & " and rcampos.fecbajas is null ") Then Exit Sub
    End If
    
    'baja de socios
    If Opcion(0).Value And Opcion1(1).Value Then
        '[Monica]19/12/2012: damos aviso si hay entradas esta campa�a
        If HayEntradasSocio(NumCod) Then
            If MsgBox("Este socio tiene entradas esta campa�a. � Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                Exit Sub
            End If
        End If
    
        cadTitulo = "Documento Baja de Socios"
        
        CadParam = CadParam & "pFecha=""" & txtcodigo(7).Text & """|"
        numParam = numParam + 1
    
        CadParam = CadParam & "pCausas=""" & txtcodigo(8).Text & """|"
        numParam = numParam + 1
        
        '[Monica]13/03/2014: para el caso de escalona y utxera enlazamos con el codpropiet del campo
        If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
'[Monica]04/06/2014: no tiene pq tener campos
'            If Not AnyadirAFormula(cadFormula, "{rcampos.codpropiet} = " & NumCod & " and isnull({rcampos.fecbajas}) ") Then Exit Sub
'            If Not AnyadirAFormula(cadSelect, "rcampos.codpropiet = " & NumCod & " and rcampos.fecbajas is null ") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{rsocios.codsocio} = " & NumCod) Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "rsocios.codsocio = " & NumCod) Then Exit Sub
        Else
            If Not AnyadirAFormula(cadFormula, "{rcampos.codsocio} = " & NumCod & " and isnull({rcampos.fecbajas}) ") Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "rcampos.codsocio = " & NumCod & " and rcampos.fecbajas is null ") Then Exit Sub
        End If
    End If
    
    'baja de campos
    If Opcion(1).Value And Opcion1(1).Value Then
        cadTitulo = "Documento Baja de Campos"
         
        CadParam = CadParam & "pFecha=""" & txtcodigo(7).Text & """|"
        numParam = numParam + 1
    
        CadParam = CadParam & "pCausas=""" & txtcodigo(8).Text & """|"
        numParam = numParam + 1
        
        
        Set frmMens = New frmMensajes
        frmMens.cadWHERE = " and rcampos.codsocio = " & NumCod & " and rcampos.fecbajas is null"
        frmMens.OpcionMensaje = 15
        frmMens.Show vbModal
        Set frmMens = Nothing
        
        If Not AnyadirAFormula(cadFormula, "{rcampos.codsocio} = " & NumCod & " and {rcampos.codcampo} in [" & NumCampo & "]") Then Exit Sub
        If Not AnyadirAFormula(cadSelect, "rcampos.codsocio = " & NumCod & " and rcampos.codcampo in (" & NumCampo & ")") Then Exit Sub
         
    End If
    
    'transmision de campos
    If Opcion1(2).Value Then
        If txtcodigo(0).Text = "" Then
            MsgBox "Debe introducir un Socio Receptor. Reintroduzca.", vbExclamation
            Exit Sub
        End If
        cadTitulo = "Documento de Transmisi�n Campos"
'        cadParam = cadParam & "pSocOrigen=" & NumCod & "|"
'        numParam = numParam + 1
        
        '[Monica]13/03/2014: para el caso de escalona y utxera enlazamos con el codpropiet del campo
        If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
            If Not AnyadirAFormula(cadFormula, "{rcampos.codpropiet} = " & NumCod) Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "rcampos.codpropiet = " & NumCod) Then Exit Sub
        Else
            If Not AnyadirAFormula(cadFormula, "{rcampos.codsocio} = " & NumCod) Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "rcampos.codsocio = " & NumCod) Then Exit Sub
        End If
        
        CadParam = CadParam & "pFecha=""" & txtcodigo(1).Text & """|"
        numParam = numParam + 1
        
        If Not AnyadirAFormula(cadFormula, "{rsocios_alias.codsocio} = " & txtcodigo(0).Text) Then Exit Sub

        Set frmMens = New frmMensajes
        
        '[Monica]13/03/2014: para el caso de escalona y utxera enlazamos con el codpropiet del campo
        If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
            frmMens.cadWHERE = " and rcampos.codpropiet = " & NumCod & " and rcampos.fecbajas is null"
        Else
            frmMens.cadWHERE = " and rcampos.codsocio = " & NumCod & " and rcampos.fecbajas is null"
        End If
        frmMens.OpcionMensaje = 15
        frmMens.Show vbModal
        Set frmMens = Nothing
    
        If Not AnyadirAFormula(cadFormula, "{rcampos.codcampo} in [" & NumCampo & "]") Then Exit Sub
        If Not AnyadirAFormula(cadSelect, "{rcampos.codcampo} in (" & NumCampo & ")") Then Exit Sub
    
    End If
    
    'Nombre fichero .rpt a Imprimir
    If Opcion(0) And Opcion1(0) Then indRPT = 16 ' alta socios
    If Opcion(0) And Opcion1(1) Then indRPT = 17 ' baja socios
    If Opcion(1) And Opcion1(0) Then indRPT = 18 ' alta campos
    If Opcion(1) And Opcion1(1) Then indRPT = 19 ' baja campos
    If Opcion1(2) Then indRPT = 28 ' transmision de campos
    
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
    
    frmImprimir.NombreRPT = nomDocu
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If ((Opcion(0) And Opcion1(0)) Or (Opcion(0) And Opcion1(1))) And (vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10) Then
        If HayRegParaInforme("rsocios", cadSelect) Then
            LlamarImprimir
        End If
    Else
        If HayRegParaInforme("rcampos", cadSelect) Then
            LlamarImprimir
        End If
    End If
End Sub

Private Sub cmdBajar_Click()
'Bajar el item seleccionado del listview2
'    BajarItemList Me.ListView1
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdSubir_Click()
'    SubirItemList Me.ListView1
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 12 ' Listado de Calidades
                PonerFoco txtcodigo(18)
        
            Case 13 ' Listado de Socios por seccion
                PonerFoco txtcodigo(8)
                
        End Select
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
    Set List = New Collection
    For H = 24 To 27
        List.Add H
    Next H
    For H = 1 To 10
        List.Add H
    Next H
    List.Add 12
    List.Add 13
    List.Add 14
    List.Add 15
    List.Add 18
    List.Add 19
    
    
    For H = 0 To imgBuscar.Count - 1
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    
' ### [Monica] 09/11/2006    he sustituido el anterior
'    For h = 0 To imgBuscar.Count - 1
'        Me.imgBuscar(h).Picture = frmPpal.imgListImages16.ListImages(1).Picture
'    Next h
'    Me.imgBuscar(8).Picture = frmPpal.imgListImages16.ListImages(1).Picture
'    Me.imgBuscar(9).Picture = frmPpal.imgListImages16.ListImages(1).Picture
'    Me.imgBuscar(10).Picture = frmPpal.imgListImages16.ListImages(1).Picture
'    Me.imgBuscar(11).Picture = frmPpal.imgListImages16.ListImages(1).Picture
'    Me.imgBuscar(16).Picture = frmPpal.imgListImages16.ListImages(1).Picture
'    Me.imgBuscar(17).Picture = frmPpal.imgListImages16.ListImages(1).Picture
'    Me.imgBuscar(18).Picture = frmPpal.imgListImages16.ListImages(1).Picture
'    Me.imgBuscar(19).Picture = frmPpal.imgListImages16.ListImages(1).Picture
'
    
    Set List = Nothing

    '###Descomentar
'    CommitConexion
    FrameCalidadesVisible True, H, W
    
    Me.Opcion(0).Value = True
    Me.Opcion1(0).Value = True
    Opcion_Click (0)
    
    
    CargarListViewOrden (2)
'        Me.lbltitulo2.Caption = "Informe de Calidades"
'        Me.Label2(3).Caption = "Variedades"
    indFrame = 2
    Tabla = "rcalidad"
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
'    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub




Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtcodigo(indice).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmCal_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de calidades
'    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
'    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    NumCampo = CadenaSeleccion
End Sub

Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
'    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
'    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
'    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 8, 9 'SECCION
            AbrirFrmSeccion (Index)
        
        Case 0 'SOCIOS
            AbrirFrmSocios (Index)
        
        Case 18, 19 'VARIEDADES
            AbrirFrmVariedad (Index)
    
        Case 16, 17 'CALIDADES
            AbrirFrmCalidad (Index)
    End Select
    PonerFoco txtcodigo(indCodigo)
End Sub

Private Sub ListView1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
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

    menu = Me.Height - Me.ScaleHeight 'ac� tinc el heigth del men� i de la toolbar

    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40

    ' *** repasar si el camp es txtAux o Text1 ***
    If Index = 1 Then indice = Index + 3
    If Index = 0 Then indice = Index + 2
    If Index = 3 Then indice = Index + 6
    If Index = 2 Then indice = Index + 5
    If Index = 4 Then indice = 1
    
    imgFec(0).Tag = indice '<===
    If txtcodigo(indice).Text <> "" Then frmC.NovaData = txtcodigo(indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtcodigo(indice) '<===
    ' ********************************************
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub Opcion_Click(Index As Integer)

    FrameAltaSocio.visible = (Opcion(0).Value And Opcion1(0).Value) And Not (Opcion(1).Value And Opcion1(0).Value) And Not (Opcion(0).Value And Opcion1(1).Value)
    FrameAltaSocio.Enabled = (Opcion(0).Value And Opcion1(0).Value) And Not (Opcion(1).Value And Opcion1(0).Value) And Not (Opcion(0).Value And Opcion1(1).Value)
    
    FrameAltaCampo.visible = Not (Opcion(0).Value And Opcion1(0).Value) And (Opcion(1).Value And Opcion1(0).Value) And Not (Opcion(0).Value And Opcion1(1).Value)
    FrameAltaCampo.Enabled = Not (Opcion(0).Value And Opcion1(0).Value) And (Opcion(1).Value And Opcion1(0).Value) And Not (Opcion(0).Value And Opcion1(1).Value)
    
    FrameBajaSocio.visible = Not (Opcion(0).Value And Opcion1(0).Value) And Not (Opcion(1).Value And Opcion1(0).Value) And (Opcion1(1).Value)
    FrameBajaSocio.Enabled = Not (Opcion(0).Value And Opcion1(0).Value) And Not (Opcion(1).Value And Opcion1(0).Value) And (Opcion1(1).Value)

    If Opcion1(1).Value And Opcion(0).Value Then FrameBajaSocio.Caption = "Baja Socio"
    If Opcion1(1).Value And Opcion(1).Value Then FrameBajaSocio.Caption = "Baja Campo"

    PonerFocoFrame

End Sub

Private Sub Opcion1_Click(Index As Integer)

    FrameAltaSocio.visible = (Opcion(0).Value And Opcion1(0).Value) And Not (Opcion(1).Value And Opcion1(0).Value) And Not (Opcion1(1).Value)
    FrameAltaSocio.Enabled = (Opcion(0).Value And Opcion1(0).Value) And Not (Opcion(1).Value And Opcion1(0).Value) And Not (Opcion1(1).Value)

    FrameAltaCampo.visible = Not (Opcion(0).Value And Opcion1(0).Value) And (Opcion(1).Value And Opcion1(0).Value) And Not (Opcion1(1).Value)
    FrameAltaCampo.Enabled = Not (Opcion(0).Value And Opcion1(0).Value) And (Opcion(1).Value And Opcion1(0).Value) And Not (Opcion1(1).Value)

    FrameBajaSocio.visible = Not (Opcion(0).Value And Opcion1(0).Value) And Not (Opcion(1).Value And Opcion1(0).Value) And (Opcion1(1).Value)
    FrameBajaSocio.Enabled = Not (Opcion(0).Value And Opcion1(0).Value) And Not (Opcion(1).Value And Opcion1(0).Value) And (Opcion1(1).Value)
    
    FrameTransmision.visible = Opcion1(2).Value
    FrameTransmision.Enabled = Opcion1(2).Value
    
    If Opcion1(1).Value And Opcion(0).Value Then FrameBajaSocio.Caption = "Baja Socio"
    If Opcion1(1).Value And Opcion(1).Value Then FrameBajaSocio.Caption = "Baja Campo"
    
    Frame2.Enabled = (Opcion1(2).Value = 0)
    If Opcion1(2).Value Then Opcion(1).Value = True
    
    PonerFocoFrame
    
    
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
            Case 0: KEYBusqueda KeyAscii, 0 ' socio receptor de transmision
            Case 1: KEYFecha KeyAscii, 4 'fecha de transmision
            Case 2: KEYFecha KeyAscii, 0 'fecha
            Case 4: KEYFecha KeyAscii, 1 'fecha
            Case 7: KEYFecha KeyAscii, 2 'fecha
            Case 9: KEYFecha KeyAscii, 3 'fecha
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0 ' socio receptor
            If txtcodigo(Index).Text <> "" Then
                PonerFormatoEntero txtcodigo(Index)
                txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
            End If
            
        Case 1, 2, 4, 7 'FECHAS
            If txtcodigo(Index).Text <> "" Then PonerFormatoFecha txtcodigo(Index)

        Case 9 ' fechas
            If txtcodigo(Index).Text <> "" Then PonerFormatoFecha txtcodigo(Index), True

        Case 5 ' importe
            If txtcodigo(Index).Text <> "" Then PonerFormatoDecimal txtcodigo(Index), 1
            
    End Select
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim Cad As String
Dim Tabla As String
Dim Titulo As String

    'Llamamos a al form
    Cad = ""
    Conexion = cAgro    'Conexi�n a BD: Ariges
'    Select Case OpcionListado
'        Case 7 'Traspaso de Almacenes
'            cad = cad & "N� Trasp|scatra|codtrasp|N|0000000|40�Almacen Origen|scatra|almaorig|N|000|20�Almacen Destino|scatra|almadest|N|000|20�Fecha|scatra|fechatra|F||20�"
'            Tabla = "scatra"
'            titulo = "Traspaso Almacenes"
'        Case 8 'Movimientos de Almacen
'            cad = cad & "N� Movim.|scamov|codmovim|N|0000000|40�Almacen|scamov|codalmac|N|000|30�Fecha|scamov|fecmovim|F||30�"
'            Tabla = "scamov"
'            titulo = "Movimientos Almacen"
'        Case 9, 12, 13, 14, 15, 16, 17 '9: Movimientos Articulos
'                   '12: Inventario Articulos
'                   '14:Actualizar Diferencias de Stock Inventariado
'                   '16: Listado Valoracion stock inventariado
'            cad = cad & "C�digo|sartic|codartic|T||30�Denominacion|sartic|nomartic|T||70�"
'            Tabla = "sartic"
'            titulo = "Articulos"
'    End Select
          
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vtabla = Tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
        'frmB.vDevuelve = "0|1|"
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = Titulo
        frmB.vSelElem = 0
'        frmB.vConexionGrid = Conexion
'        frmB.vBuscaPrevia = 1
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
'        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub FrameCalidadesVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de clientes
    Me.FrameCalidades.visible = visible
    If visible = True Then
        Me.FrameCalidades.Top = -90
        Me.FrameCalidades.Left = 0
        Me.FrameCalidades.Height = 6255
        Me.FrameCalidades.Width = 6390
        W = Me.FrameCalidades.Width
        H = Me.FrameCalidades.Height
    End If
End Sub

Private Sub FrameSociosSeccionVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
''Frame para el listado de socios por seccion
'    Me.FrameSociosSeccion.visible = visible
'    If visible = True Then
'        Me.FrameSociosSeccion.Top = -90
'        Me.FrameSociosSeccion.Left = 0
'        Me.FrameSociosSeccion.Height = 4820
'        Me.FrameSociosSeccion.Width = 6600
'        w = Me.FrameSociosSeccion.Width
'        h = Me.FrameSociosSeccion.Height
'    End If
End Sub

Private Sub CargarListViewOrden(Index As Integer)
Dim ItmX As ListItem

'    'Los encabezados
''    ListView1(Index).ColumnHeaders.Clear
''    ListView1(Index).ColumnHeaders.Add , , "Campo", 1390
'
'    Select Case Index
'        Case 0
'            Set ItmX = ListView1(Index).ListItems.Add
'            ItmX.Text = "Codigo"
'            Set ItmX = ListView1(Index).ListItems.Add
'            ItmX.Text = "Alfab�tico"
'        Case 1
'            Set ItmX = ListView1(Index).ListItems.Add
'            ItmX.Text = "Clase"
'            Set ItmX = ListView1(Index).ListItems.Add
'            ItmX.Text = "Producto"
'        Case 2
'            Set ItmX = ListView1(Index).ListItems.Add
'            ItmX.Text = "Variedad"
'            Set ItmX = ListView1(Index).ListItems.Add
'            ItmX.Text = "Calidad"
'        Case 3
'            Set ItmX = ListView1(Index).ListItems.Add
'            ItmX.Text = "Seccion"
'            Set ItmX = ListView1(Index).ListItems.Add
'            ItmX.Text = "Socio"
'        Case 4
'            Set ItmX = ListView1(Index).ListItems.Add
'            ItmX.Text = "Trabajador"
'            Set ItmX = ListView1(Index).ListItems.Add
'            ItmX.Text = "Fecha"
'    End Select
'
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
'A�ade a cadFormula y cadSelect la cadena de seleccion:
'       "(codigo>=codD AND codigo<=codH)"
' y a�ade a cadParam la cadena para mostrar en la cabecera informe:
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
        .EnvioEMail = False
        .Titulo = cadTitulo
        .ConSubInforme = True
'        .NombreRPT = cadNombreRPT
        .Opcion = OpcionListado
        .Show vbModal
    End With
End Sub

Private Function PonerGrupo(numGrupo As Byte, cadgrupo As String) As Byte
Dim campo As String
Dim nomCampo As String

    campo = "pGroup" & numGrupo & "="
    nomCampo = "pGroup" & numGrupo & "Name="
    PonerGrupo = 0

    Select Case cadgrupo
'        Case "Codigo"
'            cadParam = cadParam & campo & "{" & Tabla & ".codclien}" & "|"
'            cadParam = cadParam & nomcampo & " {" & "scoope" & ".nomcoope}" & "|"
'            cadParam = cadParam & "pTitulo1" & "=""C�digo""" & "|"
'            numParam = numParam + 3
'
'        Case "Alfabetico"
'            cadParam = cadParam & campo & "{" & Tabla & ".tipsocio}" & "|"
'            cadParam = cadParam & nomcampo & " {" & "tiposoci" & ".nomtipso}" & "|"
'            cadParam = cadParam & "pTitulo1" & "=""Colectivo""" & "|"
'            numParam = numParam + 3
'
        
        'Informe de variedades
        Case "Clase"
            CadParam = CadParam & campo & "{" & Tabla & ".codclase}" & "|"
            CadParam = CadParam & nomCampo & " {" & "clases" & ".nomclase}" & "|"
            CadParam = CadParam & "pTitulo1" & "=""Producto""" & "|"
            numParam = numParam + 3
            
        Case "Producto"
            CadParam = CadParam & campo & "{" & Tabla & ".codprodu}" & "|"
            CadParam = CadParam & nomCampo & " {" & "productos" & ".nomprodu}" & "|"
            CadParam = CadParam & "pTitulo1" & "=""Clase""" & "|"
            numParam = numParam + 3

        'Informe de calibres
        Case "Seccion"
            CadParam = CadParam & campo & "{" & Tabla & ".codsecci}" & "|"
            CadParam = CadParam & nomCampo & "{rseccion.nomsecci}" & "|"
            CadParam = CadParam & "pTitulo1" & "=""Seccion""" & "|"
            numParam = numParam + 3
            
        Case "Socio"
            CadParam = CadParam & campo & "{" & Tabla & ".codsocio}" & "|"
            CadParam = CadParam & nomCampo & " {" & "rsocios" & ".nomsocio}" & "|"
            CadParam = CadParam & "pTitulo1" & "=""Socio""" & "|"
            numParam = numParam + 3
            
        'Informe de calidades
        Case "Variedad"
            CadParam = CadParam & campo & "{" & Tabla & ".codvarie}" & "|"
            CadParam = CadParam & nomCampo & "{variedades.nomvarie}" & "|"
            CadParam = CadParam & "pTitulo1" & "=""Variedad""" & "|"
            numParam = numParam + 3
            
        Case "Calidad"
            CadParam = CadParam & campo & "{" & Tabla & ".codcalid}" & "|"
            CadParam = CadParam & nomCampo & " {" & "rcalidad" & ".nomcalid}" & "|"
            CadParam = CadParam & "pTitulo1" & "=""Calidad""" & "|"
            numParam = numParam + 3
            
            
            
'        'Informe de Horas Trabajadas
'        Case "Trabajador"
'            cadParam = cadParam & campo & "{" & Tabla & ".codtraba}" & "|"
'            cadParam = cadParam & nomcampo & " {" & "straba" & ".nomtraba}" & "|"
'            cadParam = cadParam & "pTitulo1" & "=""Fecha""" & "|"
'            numParam = numParam + 3
'
'        Case "Fecha"
'            cadParam = cadParam & "pGroup1=" & "{" & Tabla & ".fechahora}" & "|"
'            cadParam = cadParam & "pGroup1Name=" & " {" & "horas" & ".fechahora}" & "|"
'            cadParam = cadParam & "pTitulo1" & "=""Trabajadores""" & "|"
'            numParam = numParam + 3
        

End Select

End Function

Private Function PonerOrden(cadgrupo As String) As Byte
Dim campo As String
Dim nomCampo As String

    PonerOrden = 0

    Select Case cadgrupo
        Case "Codigo"
            CadParam = CadParam & "Orden" & "= {" & Tabla
            Select Case OpcionListado
                Case 10
                    CadParam = CadParam & ".codclien}|"
                Case 11
                    CadParam = CadParam & ".codprove}|"
            End Select
            Tipo = "C�digo"
        Case "Alfab�tico"
            CadParam = CadParam & "Orden" & "= {" & Tabla
            Select Case OpcionListado
                Case 10
                    CadParam = CadParam & ".nomclien}|"
                Case 11
                    CadParam = CadParam & ".nomprove}|"
            End Select
            Tipo = "Alfab�tico"
    End Select
    
    numParam = numParam + 1

End Function

Private Sub AbrirFrmCalidad(indice As Integer)
    indCodigo = indice
    Set frmCal = New frmManCalidades
    frmCal.DatosADevolverBusqueda = "2|3|"
'    frmCli.DeConsulta = True
'    frmCli.CodigoActual = txtCodigo(indCodigo)
    frmCal.Show vbModal
    Set frmCal = Nothing
End Sub

Private Sub AbrirFrmSeccion(indice As Integer)
    indCodigo = indice
    Set frmSec = New frmManSeccion
    frmSec.DatosADevolverBusqueda = "0|1|"
'    frmCli.DeConsulta = True
'    frmCli.CodigoActual = txtCodigo(indCodigo)
    frmSec.Show vbModal
    Set frmSec = Nothing
End Sub

Private Sub AbrirFrmSocios(indice As Integer)
    indCodigo = indice
    Set frmSoc = New frmManSocios
    frmSoc.DatosADevolverBusqueda = "0|1|"
'    frmCli.DeConsulta = True
'    frmCli.CodigoActual = txtCodigo(indCodigo)
    frmSoc.Show vbModal
    Set frmSoc = Nothing
End Sub

Private Sub AbrirFrmVariedad(indice As Integer)
    indCodigo = indice
    Set frmVar = New frmComVar
    frmVar.DatosADevolverBusqueda = "0|1|"
'    frmCli.DeConsulta = True
'    frmCli.CodigoActual = txtCodigo(indCodigo)
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub

Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
'        .SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
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
        .Opcion = OpcionListado
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


'Private Function DatosOk() As Boolean
'Dim b As Boolean
'Dim Sql As String
'Dim Sql2 As String
'Dim vClien As CSocio
'' a�adido
'Dim Mens As String
'Dim numfactu As String
'Dim numser As String
'Dim Fecha As Date
'
'    b = True
'    If txtCodigo(9).Text = "" Or txtCodigo(10).Text = "" Or txtCodigo(11).Text = "" Then
'        MsgBox "Debe introducir la letra de serie, el n�mero de factura y la fecha de factura para localizar la factura a rectificar", vbExclamation
'        b = False
'    End If
'    If b And vParamAplic.Cooperativa = 2 Then
'        If txtCodigo(8).Text = "" Then
'            MsgBox "Debe introducir el cliente. Reintroduzca.", vbExclamation
'            b = False
'        Else
'            ' obtenemos la cooperativa del anterior cliente y del nuevo pq tienen que coincidir
'            ' anterior cliente
'            Sql = ""
'            Sql = DevuelveDesdeBDNew(cAgro, "ssocio", "codcoope", "codsocio", txtCodigo(12).Text, "N")
'            ' nuevo cliente
'            Sql2 = ""
'            Sql2 = DevuelveDesdeBDNew(cAgro, "ssocio", "codcoope", "codsocio", txtCodigo(8).Text, "N")
'            If Sql <> Sql2 Then
'                MsgBox "El nuevo cliente debe pertenecer al mismo colectivo que el cliente de la factura a rectificar. Reintroduzca.", vbExclamation
'                b = False
'            End If
'        End If
'    End If
'
''    If b And Contabilizada = 1 And vParamAplic.NumeroConta <> 0 And txtCodigo(8).Text <> "" Then 'comprobamos que la cuenta contable del nuevo cliente existe
''        Set vClien = New CSocio
''        If vClien.LeerDatos(txtCodigo(8).Text) Then
''            sql = ""
''            sql = DevuelveDesdeBDNew(cConta, "cuentas", "codmacta", "codmacta", vClien.CuentaConta, "T")
''            If sql = "" Then
''                MsgBox "La cuenta contable del nuevo cliente no existe. Revise", vbExclamation
''                b = False
''            End If
''        End If
''    End If
'
'' a�adido
''    b = True
'
'    If ConTarjetaProfesional(txtCodigo(9).Text, txtCodigo(10).Text, txtCodigo(11).Text) Then
'        MsgBox "Este Factura tiene alguna tarjeta profesional, no se permite hacer la factura rectificativa", vbExclamation
'        b = False
'    Else
'        If txtCodigo(13).Text = "" Then
'            MsgBox "Debe introducir obligatoriamente una Fecha de Facturaci�n.", vbExclamation
'            b = False
'            PonerFoco txtCodigo(13)
'        Else
'                If Not FechaDentroPeriodoContable(CDate(txtCodigo(13).Text)) Then
'                    Mens = "La Fecha de Facturaci�n no es del ejercicio actual ni siguiente. Reintroduzca."
'                    MsgBox Mens, vbExclamation
'                    b = False
'                    PonerFoco txtCodigo(13)
'                Else
'                    'VRS:2.0.1(0)
'                    If Not FechaSuperiorUltimaLiquidacion(CDate(txtCodigo(13).Text)) Then
'                        Mens = "  La Fecha de Facturaci�n es inferior a la �ltima liquidaci�n de Iva. " & vbCrLf & vbCrLf
'                        ' unicamente si el usuario es root el proceso continuar�
'                        If vSesion.Nivel > 0 Then
'                            Mens = Mens & "  El proceso no continuar�."
'                            MsgBox Mens, vbExclamation
'                            b = False
'                            PonerFoco txtCodigo(13)
'                        Else
'                            Mens = Mens & "                        � Desea continuar ?    " & vbCrLf
'                            If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                                b = False
'                                PonerFoco txtCodigo(13)
'                            End If
'                        End If
'                    End If
'                    ' la fecha de factura no debe ser inferior a la ultima factura de la serie
'                    numser = "letraser"
'                    numfactu = ""
'                    numfactu = DevuelveDesdeBDNew(cAgro, "stipom", "contador", "codtipom", "FAG", "T", numser)
'                    If numfactu <> "" Then
'                        If FechaFacturaInferiorUltimaFacturaSerieHco(CDate(txtCodigo(13).Text), CLng(numfactu), numser, 0) Then
'                            Mens = "La Fecha de Factura es inferior a la �ltima factura de la serie." & vbCrLf & vbCrLf
'                            Mens = Mens & "                        � Desea continuar ?    " & vbCrLf
'                            If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
'                                b = False
'                                PonerFoco txtCodigo(13)
'                            End If
'                        End If
'                    End If
'                End If
'        End If
'    End If
'
'    DatosOk = b
'
'
'' end a�adido
'    If b And txtCodigo(87).Text = "" Then
'        MsgBox "Para rectificar una factura ha de introducir obligatoriamente un motivo. Reintroduzca", vbExclamation
'        b = False
'    End If
'    DatosOk = b
'
'End Function
'

Private Function ConTarjetaProfesional(letraser As String, numfactu As String, fecfactu As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset

    Sql = "select count(*) from slhfac, starje where letraser = " & DBSet(letraser, "T") & " and numfactu = " & DBSet(numfactu, "N")
    Sql = Sql & " and fecfactu = " & DBSet(fecfactu, "F") & " and starje.tiptarje = 2 and slhfac.numtarje = starje.numtarje "
    
    ConTarjetaProfesional = (TotalRegistros(Sql) <> 0)

End Function

Private Sub PonerFocoFrame()

    If Me.FrameAltaCampo.visible Then PonerFoco txtcodigo(2)
    If Me.FrameAltaSocio.visible Then PonerFoco txtcodigo(4)
    If Me.FrameBajaSocio.visible Then PonerFoco txtcodigo(7)
    If Me.FrameTransmision.visible Then PonerFoco txtcodigo(1)

End Sub
