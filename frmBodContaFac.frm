VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBodContaFac 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Integraci�n Contable de Facturas de Retirada"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7095
   Icon            =   "frmBodContaFac.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCobros 
      Height          =   4530
      Left            =   15
      TabIndex        =   6
      Top             =   135
      Width           =   7050
      Begin VB.Frame Frame1 
         Caption         =   "Datos para Selecci�n"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1200
         Left            =   90
         TabIndex        =   7
         Top             =   225
         Width           =   6825
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
            Left            =   3825
            MaxLength       =   10
            TabIndex        =   1
            Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
            Top             =   660
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
            Index           =   5
            Left            =   1470
            MaxLength       =   10
            TabIndex        =   0
            Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
            Top             =   660
            Width           =   1350
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   6
            Left            =   3540
            Picture         =   "frmBodContaFac.frx":000C
            ToolTipText     =   "Buscar fecha"
            Top             =   675
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   5
            Left            =   1185
            Picture         =   "frmBodContaFac.frx":0097
            ToolTipText     =   "Buscar fecha"
            Top             =   645
            Width           =   240
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
            Left            =   2910
            TabIndex        =   17
            Top             =   675
            Width           =   600
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
            Left            =   525
            TabIndex        =   16
            Top             =   645
            Width           =   645
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
            Index           =   2
            Left            =   225
            TabIndex        =   15
            Top             =   315
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos para la contabilizaci�n"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1485
         Left            =   90
         TabIndex        =   8
         Top             =   1500
         Width           =   6840
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
            Index           =   1
            Left            =   2475
            MaxLength       =   10
            TabIndex        =   2
            Tag             =   "C�digo Postal|T|S|||clientes|codposta|||"
            Top             =   405
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
            Index           =   4
            Left            =   3870
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   870
            Width           =   2685
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
            Index           =   4
            Left            =   2475
            MaxLength       =   10
            TabIndex        =   3
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   870
            Width           =   1350
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Vencimiento"
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
            TabIndex        =   14
            Top             =   450
            Width           =   1965
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   1
            Left            =   2205
            Picture         =   "frmBodContaFac.frx":0122
            ToolTipText     =   "Buscar fecha"
            Top             =   405
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Banco Prevista"
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
            Height          =   195
            Index           =   24
            Left            =   180
            TabIndex        =   10
            Top             =   915
            Width           =   1890
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   2205
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   870
            Width           =   240
         End
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
         Height          =   360
         Left            =   5865
         TabIndex        =   5
         Top             =   3810
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
         Height          =   360
         Left            =   4680
         TabIndex        =   4
         Top             =   3810
         Width           =   1035
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   315
         Left            =   90
         TabIndex        =   11
         Top             =   3060
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   13
         Top             =   3450
         Width           =   6750
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   12
         Top             =   3720
         Width           =   4470
      End
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   3465
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   810
      Width           =   2685
   End
   Begin VB.TextBox txtcodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   10
      Left            =   2250
      MaxLength       =   10
      TabIndex        =   18
      Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
      Top             =   810
      Width           =   1125
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Cta.Retenci�n"
      ForeColor       =   &H00972E0B&
      Height          =   195
      Index           =   3
      Left            =   450
      TabIndex        =   20
      Top             =   855
      Width           =   1395
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   10
      Left            =   1980
      ToolTipText     =   "Buscar Cuenta Contable"
      Top             =   810
      Width           =   240
   End
End
Attribute VB_Name = "frmBodContaFac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MONICA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public Tipo As Byte '0 = facturas de retirada de almazara
                    '1 = facturas de retirada de bodega

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto


Private Conexion As Byte
'1.- Conexi�n a BD Ariges  2.- Conexi�n a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCtas As frmCtasConta 'cuentas de contabilidad
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmFPa As frmForpaConta 'formas de pago de la contabilidad
Attribute frmFPa.VB_VarHelpID = -1
Private WithEvents frmSec As frmManSeccion 'secciones
Attribute frmSec.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNomRPT As String 'Nombre del informe
Private conSubRPT As Boolean 'Si el informe tiene subreports



Dim indCodigo As Integer 'indice para txtCodigo

'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Codigo As String 'C�digo para FormulaSelection de Crystal Report
Dim TipCod As String

Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

Dim vSeccion As CSeccion

Dim cContaFra As cContabilizarFacturas



Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me

End Sub

Private Sub cmdAceptar_Click()
Dim SQL As String
Dim I As Byte
Dim cadWHERE As String
Dim cDesde As String
Dim cHasta As String

    If Not DatosOK Then Exit Sub

    cadSelect = "{rbodfacturas.intconta}=0 "

    Select Case Tipo
        Case 0 ' almazara
            If Not AnyadirAFormula(cadSelect, "mid({rbodfacturas.codtipom},2,2)='ZA'") Then Exit Sub
        Case 1 ' bodega
            If Not AnyadirAFormula(cadSelect, "mid({rbodfacturas.codtipom},2,2)='AB'") Then Exit Sub
    End Select

    'D/H Fecha factura
    cDesde = Trim(txtCodigo(5).Text)
    cHasta = Trim(txtCodigo(6).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rbodfacturas.fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If

    If Not HayRegParaInforme("rbodfacturas", cadSelect) Then Exit Sub

    ContabilizarFacturas "rbodfacturas", cadSelect
     'Eliminar la tabla TMP
    BorrarTMPFacturas
    'Desbloqueamos ya no estamos contabilizando facturas
    DesBloqueoManual ("CONFRE") 'CONtabilizar Facturas de REtirada

eError:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "No se ha realizado el proceso de contabilizaci�n de facturas de Retirada. Llame a soporte."
    End If

    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    cmdCancel_Click
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(5)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection
Dim I As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon


    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
    For I = 4 To 4
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    For I = 10 To 10
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    
    Select Case Tipo
        Case 0 ' almazara
            ConexionConta vParamAplic.SeccionAlmaz
        Case 1 ' bodega
            ConexionConta vParamAplic.SeccionBodega
    End Select
    
    
'   cuentas contables
    Select Case Tipo
        Case 0 ' cta de banco prevista de almazara
            txtCodigo(4).Text = vParamAplic.CtaBancoAlmz   ' cuenta contable de banco almz
        Case 1
            txtCodigo(4).Text = vParamAplic.CtaBancoBOD   ' cuenta contable de banco bodega
    End Select
    
    txtNombre(4).Text = PonerNombreCuenta(txtCodigo(4), 0)
    
    txtCodigo(5).Text = Format(Now, "dd/mm/yyyy") ' fecha de factura desde
    txtCodigo(6).Text = Format(Now, "dd/mm/yyyy") ' fecha de factura hasta
    txtCodigo(1).Text = Format(Now, "dd/mm/yyyy") ' fecha de vencimiento
'    txtcodigo(0).Text = Format(Now, "dd/mm/yyyy") ' fecha de recepcion
            
    '###Descomentar
'    CommitConexion

    FrameCobrosVisible True, H, W
    Pb1.visible = False


    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
'   Me.Width = w + 70
'   Me.Height = h + 350


End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not vSeccion Is Nothing Then
        vSeccion.CerrarConta
        Set vSeccion = Nothing
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(1).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Cuentas contables
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFPa_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de formas de pago de contabilidad
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtCodigo(indCodigo).Text = Format(txtCodigo(indCodigo).Text, "000")
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
    imgFec(1).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index).Text <> "" Then frmC.NovaData = txtCodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(1).Tag))
    ' ***************************
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 4 'cuenta contable banco
            AbrirFrmCuentas (Index)
          
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
'14/02/2007
'    KEYpress KeyAscii
' ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 1: KEYFecha KeyAscii, 1 'fecha desde factura
            Case 6: KEYFecha KeyAscii, 6 'fecha hasta factura
            Case 5: KEYFecha KeyAscii, 5 'fecha vencimiento
            Case 4: KEYBusqueda KeyAscii, 4 'cta contable banco
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
        Case 4 ' CUENTAS CONTABLES ( banco )
            If vSeccion Is Nothing Then Exit Sub
        
            If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = PonerNombreCuenta(txtCodigo(Index), 2)
            If txtNombre(Index).Text = "" Then
                MsgBox "N�mero de Cuenta contable no existe en la contabilidad. Reintroduzca.", vbExclamation
            End If

        Case 5, 6 'FECHAS
            If txtCodigo(Index).Text <> "" Then
                If PonerFormatoFecha(txtCodigo(Index)) Then
                    If Index = 5 Then
                        txtCodigo(6).Text = txtCodigo(5).Text
                    End If
                End If
            End If

        Case 1 'FECHAS de vencimiento
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)

    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
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

Private Sub AbrirFrmCuentas(Indice As Integer)
    indCodigo = Indice
    Set frmCtas = New frmCtasConta
    frmCtas.DatosADevolverBusqueda = "0|1|"
    frmCtas.CodigoActual = txtCodigo(indCodigo)
'    frmCtas.Conexion = cContaFacSoc
'    frmCtas.Facturas = False
    frmCtas.Show vbModal
    Set frmCtas = Nothing
End Sub

Private Sub AbrirFrmForpaConta(Indice As Integer)
    indCodigo = Indice
    Set frmFPa = New frmForpaConta
    frmFPa.DatosADevolverBusqueda = "0|1|"
    frmFPa.CodigoActual = txtCodigo(indCodigo)
'    frmFpa.Conexion = cContaFacSoc
    frmFPa.Show vbModal
    Set frmFPa = Nothing
End Sub


Private Function DatosOK() As Boolean
Dim B As Boolean
Dim Orden1 As String
Dim Orden2 As String
Dim FFin As Date
Dim cta As String

   B = True

   If txtCodigo(6).Text = "" Then
        MsgBox "Introduzca la Fecha de Factura a contabilizar.", vbExclamation
        B = False
        PonerFoco txtCodigo(6)
   Else
        ' comprobamos que la contabilizacion se encuentre en los ejercicios contables
         Orden1 = ""
         Orden1 = DevuelveDesdeBDNew(cConta, "parametros", "fechaini", "", "", "", "", "", "", "", "", "", "")

         Orden2 = ""
         Orden2 = DevuelveDesdeBDNew(cConta, "parametros", "fechafin", "", "", "", "", "", "", "", "", "", "")
         FIni = CDate(Orden1)
         FFin = CDate(Orden2)
         If Not (CDate(Orden1) <= CDate(txtCodigo(6).Text) And CDate(txtCodigo(6).Text) < CDate(Day(FIni) & "/" & Month(FIni) & "/" & Year(FIni) + 2)) Then
            MsgBox "La Fecha de la contabilizaci�n no es del ejercicio actual ni del siguiente. Reintroduzca.", vbExclamation
            B = False
            PonerFoco txtCodigo(6)
         End If
   End If


   If txtCodigo(1).Text = "" And B Then
        MsgBox "Introduzca la Fecha de Vencimiento a contabilizar.", vbExclamation
        B = False
        PonerFoco txtCodigo(1)
   End If

   'cta contable de banco
   If B Then
        If txtCodigo(4).Text = "" Then
             MsgBox "Introduzca la Cta.Contable de Banco para contabilizar.", vbExclamation
             B = False
             PonerFoco txtCodigo(4)
        Else
             cta = ""
             cta = DevuelveDesdeBDNew(cConta, "cuentas", "codmacta", "codmacta", txtCodigo(4).Text, "T")
             If cta = "" Then
                 MsgBox "La cuenta contable de Banco no existe. Reintroduzca.", vbExclamation
                 B = False
                 PonerFoco txtCodigo(4)
             End If
        End If
    End If
   
   
   DatosOK = B

End Function



Private Sub ContabilizarFacturas(cadTabla As String, cadWHERE As String)
'Contabiliza Facturas de Clientes o de Proveedores
Dim SQL As String
Dim B As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String

    SQL = "CONFRE" 'contabilizar facturas de REtirada
    
    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (SQL)
    If Not BloqueoManual(SQL, "1") Then
        MsgBox "No se pueden Contabilizar Facturas de Retirada. Hay otro usuario contabilizando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


    'comprobar que se han rellenado los dos campos de fecha
    'sino rellenar con fechaini o fechafin del ejercicio
    'que guardamos en vbles Orden1,Orden2
    If txtCodigo(5).Text = "" Then
       txtCodigo(5).Text = vEmpresa.FechaIni  'fechaini del ejercicio de la conta
    End If

    If txtCodigo(6).Text = "" Then
       txtCodigo(6).Text = vEmpresa.FechaFin  'fecha fin del ejercicio de la conta
    End If


    'Comprobar que el intervalo de fechas D/H esta dentro del ejercicio de la
    'contabilidad par ello mirar en la BD de la Conta los par�metros
    If Not ComprobarFechasConta(6) Then Exit Sub

    'comprobar si existen en Ariagrorec facturas anteriores al periodo solicitado
    'sin contabilizar
    If Me.txtCodigo(5).Text <> "" Then 'anteriores a fechadesde
        SQL = "SELECT COUNT(*) FROM " & cadTabla
        SQL = SQL & " WHERE fecfactu <"
        SQL = SQL & DBSet(txtCodigo(5), "F") & " AND intconta=0 "
        
        Select Case Tipo
            Case 0
                SQL = SQL & " and mid(" & cadTabla & ".codtipom,2,2) = 'ZA'"
            Case 1
                SQL = SQL & " and mid(" & cadTabla & ".codtipom,2,2) = 'AB'"
        End Select
        If RegistrosAListar(SQL) > 0 Then
            MsgBox "Hay Facturas anteriores sin contabilizar.", vbExclamation
            Exit Sub
        End If
    End If


'    'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
'    If Not BloqueaRegistro(cadTabla, cadWhere) Then
'        MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If


    '==========================================================
    'REALIZAR COMPROBACIONES ANTES DE CONTABILIZAR FACTURAS
    '==========================================================

'    Me.lblProgess(0).Caption = "Comprobaciones: "
'    CargarProgres Me.ProgressBar1, 100

    'Cargar tabla TEMP con las Facturas que vamos a Trabajar
    B = CrearTMPFacturas(cadTabla, cadWHERE)
    If Not B Then Exit Sub
    

    'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
'    TerminaBloquear
    SQL = cadTabla & " INNER JOIN tmpFactu ON " & cadTabla
    SQL = SQL & ".codtipom=tmpFactu.codtipom AND "
    SQL = SQL & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "
    
    If Not BloqueaRegistro(SQL, cadWHERE) Then
        MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


    'Visualizar la barra de Progreso
    Me.Pb1.visible = True
    Me.lblProgres(0).Caption = "Comprobaciones: "
    CargarProgres Me.Pb1, 100


    'comprobar que la LETRA SERIE de parametros existen en la contabilidad y en Ariagrorec
    '--------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando letras de serie ..."
    B = ComprobarLetraSerie(cadTabla)
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    If Not B Then Exit Sub


    'comprobar que no haya N� FACTURAS en la contabilidad para esa fecha
    'que ya existan
    '-----------------------------------------------------------------------
    If cadTabla = "rbodfacturas" Then
        Me.lblProgres(1).Caption = "Comprobando N� Facturas en contabilidad ..."
        SQL = "anofaccl>=" & Year(txtCodigo(5).Text) & " AND anofaccl<= " & Year(txtCodigo(6).Text)
        B = ComprobarNumFacturas_new(cadTabla, SQL)
    End If
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not B Then Exit Sub


    'comprobar que todas las CUENTAS de los distintos socios que vamos a
    'contabilizar existen en la Conta: rsocios_seccion.codmaccli IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Contables socios en contabilidad ..."
    Select Case Tipo
        Case 0
            B = ComprobarCtaContable_new("rbodfact1", 1)
        Case 1
            B = ComprobarCtaContable_new("rbodfact2", 1)
    End Select
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not B Then Exit Sub


    'comprobar que todas las CUENTAS de venta de las variedades
    'contabilizar existen en la Conta: sfamia.ctaventa IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Ctbles Ventas en contabilidad ..."
    
    Select Case Tipo
        Case 0 ' ctaventas de almazara
            B = ComprobarCtaContable_new("rbodfact1", 2)
        Case 1 ' ctaventas de bodega
            B = ComprobarCtaContable_new("rbodfact2", 2)
    End Select
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not B Then Exit Sub



    'comprobar que todos las TIPO IVA de las distintas fecturas que vamos a
    'contabilizar existen en la Conta: rbodfacturas.codiiva1 codiiva2 codiiva3 IN (conta.tiposiva.codigiva)
    '--------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Tipos de IVA en contabilidad ..."
    B = ComprobarTiposIVA(cadTabla)
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    If Not B Then Exit Sub
    
    
    'comprobar si hay contabilidad ANALITICA: conta.parametros.autocoste=1
    'y verificar que las cuentas de rparamaplic.ctaventaalmz rparamaplic.ctagastosalmz
    'empiezan por el digito de conta.parametros.grupogto o conta.parametros.grupovta
    'obtener el centro de coste del usuario para insertarlo en linfact
    If vEmpresa.TieneAnalitica Then  'hay contab. analitica
       Me.lblProgres(1).Caption = "Comprobando Contabilidad Anal�tica ..."
           
       Select Case Tipo
            Case 0
               B = ComprobarCtaContable_new("rbodfact1", 7)
            Case 1
               B = ComprobarCtaContable_new("rbodfact2", 7)
       End Select
           
       If B Then
            '(si tiene anal�tica requiere un centro de coste para insertar en conta.linfact)
            CCoste = ""
            B = ComprobarCCoste_new(CCoste, cadTabla)
       End If
       If Not B Then Exit Sub

       CCoste = ""
    End If
    IncrementarProgres Me.Pb1, 10
    Me.Refresh

    If B Then
       Me.lblProgres(1).Caption = "Comprobando Forma de Pago ..."
       B = ComprobarFormadePago(cadTabla)
       If Not B Then Exit Sub
    End If
    IncrementarProgres Me.Pb1, 10
    Me.Refresh





    '===========================================================================
    'CONTABILIZAR FACTURAS
    '===========================================================================
    Me.lblProgres(0).Caption = "Contabilizar Facturas Retirada: "
    CargarProgres Me.Pb1, 10
    Me.lblProgres(1).Caption = "Insertando Facturas en Contabilidad..."


    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    LOG.Insertar 3, vUsu, "Contabilizar facturas Retirada: " & vbCrLf & cadTabla & vbCrLf & cadWHERE
    Set LOG = Nothing
    '-----------------------------------------------------------------------------


    '---- Crear tabla TEMP para los posible errores de facturas
    tmpErrores = CrearTMPErrFact(cadTabla)

    '---- Pasar las Facturas a la Contabilidad
    B = PasarFacturasAContab(cadTabla)

    '---- Mostrar ListView de posibles errores (si hay)
    If Not B Then
        If tmpErrores Then
            'Cargar un listview con la tabla TEMP de Errores y mostrar
            'las facturas que fallaron
            frmMensajes.OpcionMensaje = 10
            frmMensajes.Show vbModal
        Else
            MsgBox "No pueden mostrarse los errores.", vbInformation
        End If
    Else
        MsgBox "El proceso ha finalizado correctamente.", vbInformation
    End If

    'Este bien o mal, si son proveedores abriremos el listado
    'Imprimimiremos un listado de contabilizacion de facturas
    '------------------------------------------------------
    If DevuelveValor("Select count(*) from tmpinformes where codusu = " & vUsu.Codigo) > 0 Then
        InicializarVbles
        CadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
        
        CadParam = "|pDHFecha=""" & vUsu.Nombre & "   Hora: " & Format(Now, "hh:mm") & """|"
        numParam = numParam + 1
        cadFormula = "({tmpinformes.codusu} =" & vUsu.Codigo & ")"
        conSubRPT = False
        cadTitulo = "Listado contabilizacion FRAFRE"
        cadNomRPT = "rContabFRE.rpt"
        LlamarImprimir
    End If


    '---- Eliminar tabla TEMP de Errores
    BorrarTMPErrFact

End Sub

Private Function PasarFacturasAContab(cadTabla As String) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim B As Boolean
Dim I As Integer
Dim numfactu As Integer
Dim Codigo1 As String
Dim AntSocio As Long
Dim TotalTesoreria As Currency
Dim TotalFactura As Currency
Dim Facturas As String
Dim Mens As String
Dim AntFecha As String
Dim CCoste As String

    On Error GoTo EPasarFac

    PasarFacturasAContab = False

    '---- Obtener el total de Facturas a Insertar en la contabilidad
    SQL = "SELECT count(*) "
    SQL = SQL & " FROM " & cadTabla & " INNER JOIN tmpFactu "
    
    Codigo1 = "codtipom"
    SQL = SQL & " ON " & cadTabla & "." & Codigo1 & "=tmpFactu." & Codigo1
    SQL = SQL & " AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        numfactu = Rs.Fields(0)
    Else
        numfactu = 0
    End If
    Rs.Close
    Set Rs = Nothing


    'Modificacion como David
    '-----------------------------------------------------------
    ' Mosrtaremos para cada factura de PROVEEDOR
    ' que numregis le ha asignado
    SQL = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    Set cContaFra = New cContabilizarFacturas
    
    If Not cContaFra.EstablecerValoresInciales(ConnConta) Then
        'NO ha establcedio los valores de la conta.  Le dejaremos seguir, avisando que
        ' obviamente, no va a contabilizar las FRAS
        SQL = "Si continua, las facturas se insertaran en el registro, pero no ser�n contabilizadas" & vbCrLf
        SQL = SQL & "en este momento. Deber�n ser contabilizadas desde el ARICONTA" & vbCrLf & vbCrLf
        SQL = SQL & Space(50) & "�Continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    End If
    
    '---- Pasar cada una de las facturas seleccionadas a la Conta
    If numfactu > 0 Then
        CargarProgres Me.Pb1, numfactu

        'seleccinar todas las facturas que hemos insertado en la temporal (las que vamos a contabilizar)
        SQL = "SELECT * "
        SQL = SQL & " FROM tmpFactu "

        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenStatic, adLockPessimistic, adCmdText
        I = 1

        B = True
        
        ' de momento no tiene analitica
        CCoste = ""
        
        'pasar a contabilidad cada una de las facturas seleccionadas
        While Not Rs.EOF
            SQL = cadTabla & "." & Codigo1 & "=" & DBSet(Rs.Fields(0), "T") & " and numfactu=" & Rs!numfactu
            SQL = SQL & " and fecfactu=" & DBSet(Rs!fecfactu, "F")
            If PasarFacturaBOD(SQL, CCoste, txtCodigo(4).Text, txtCodigo(1).Text, Tipo, cContaFra) = False And B Then B = False

            '---- Laura 26/10/2006
            'Al pasar cada factura al hacer el commit desbloqueamos los registros
            'que teniamos bloqueados y los volvemos a bloquear
            'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
            SQL = cadTabla & " INNER JOIN tmpFactu ON " & cadTabla & "." & Codigo1 & "=tmpFactu." & Codigo1 & " AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "
            If Not BloqueaRegistro(SQL, cadTabla & "." & Codigo1 & "=tmpFactu." & Codigo1 & " AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu") Then
'                MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
'                Screen.MousePointer = vbDefault
'                Exit Sub
            End If
            '----

            IncrementarProgres Me.Pb1, 1
            Me.lblProgres(1).Caption = "Insertando Facturas en Contabilidad...   (" & I & " de " & numfactu & ")"
            Me.Refresh
            I = I + 1
            Rs.MoveNext
        Wend

        Rs.Close
        
        Set Rs = Nothing
    End If
    
    Set cContaFra = Nothing

EPasarFac:
    If Err.Number <> 0 Then B = False

    If B Then
        PasarFacturasAContab = True
    Else
        PasarFacturasAContab = False
    End If
End Function

Private Function ComprobarFechasConta(ind As Integer) As Boolean
'comprobar que el periodo de fechas a contabilizar esta dentro del
'periodo de fechas del ejercicio de la contabilidad
Dim FechaIni As String, FechaFin As String
Dim cad As String
Dim Rs As ADODB.Recordset
    
Dim vSeccion As CSeccion
    
On Error GoTo EComprobar

    ComprobarFechasConta = False
    
    
    If txtCodigo(ind).Text <> "" Then
        FechaIni = "Select fechaini,fechafin From parametros"
        Set Rs = New ADODB.Recordset
        Rs.Open FechaIni, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If Not Rs.EOF Then
            FechaIni = DBLet(Rs!FechaIni, "F")
            FechaFin = DateAdd("yyyy", 1, DBLet(Rs!FechaFin, "F"))
            'nos guardamos los valores
            Orden1 = FechaIni
            Orden2 = FechaFin
        
            If Not EntreFechas(FechaIni, txtCodigo(ind).Text, FechaFin) Then
                 cad = "El per�odo de contabilizaci�n debe estar dentro del ejercicio:" & vbCrLf & vbCrLf
                 cad = cad & "    Desde: " & FechaIni & vbCrLf
                 cad = cad & "    Hasta: " & FechaFin
                 MsgBox cad, vbExclamation
                 txtCodigo(ind).Text = ""
            Else
                ComprobarFechasConta = True
            End If
        End If
        Rs.Close
        Set Rs = Nothing
    Else
        ComprobarFechasConta = True
    End If
            
    
EComprobar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Fechas", Err.Description
End Function


Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = CadParam
        .NumeroParametros = numParam

        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 0
        .Titulo = cadTitulo
        .NombreRPT = cadNomRPT
        .ConSubInforme = conSubRPT
        .Show vbModal
    End With
End Sub


Private Sub ConexionConta(Seccion As Integer)
    
    If vSeccion Is Nothing Then
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(CStr(Seccion)) Then
            vSeccion.AbrirConta
        End If
    Else
        ' si el objeto existia: cerramos la conexion y volvemos crearlo
        vSeccion.CerrarConta
        Set vSeccion = Nothing
        
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(CStr(Seccion)) Then
            vSeccion.AbrirConta
        End If
    End If
End Sub
