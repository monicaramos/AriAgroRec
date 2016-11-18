VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTRAContaFac 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Integración Contable de Facturas de Transporte"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6315
   Icon            =   "frmTRAContaFac.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   6315
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
      Height          =   6780
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   6330
      Begin VB.Frame Frame1 
         Caption         =   "Datos para Selección"
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
         Height          =   2220
         Left            =   90
         TabIndex        =   1
         Top             =   225
         Width           =   6060
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   5
            Left            =   1500
            MaxLength       =   10
            TabIndex        =   18
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   1290
            Width           =   1050
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   6
            Left            =   3630
            MaxLength       =   10
            TabIndex        =   19
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   1320
            Width           =   1050
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   7
            Left            =   1500
            MaxLength       =   7
            TabIndex        =   16
            Tag             =   "Nº de Factura|N|N|0|9999999|schfac|numfactu|0000000|S|"
            Top             =   600
            Width           =   1035
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   8
            Left            =   3630
            MaxLength       =   7
            TabIndex        =   17
            Tag             =   "Nº de Factura|N|N|0|9999999|schfac|numfactu|0000000|S|"
            Top             =   600
            Width           =   1035
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1680
            Width           =   3180
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Factura"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   2
            Left            =   255
            TabIndex        =   33
            Top             =   975
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Desde"
            Height          =   195
            Index           =   15
            Left            =   645
            TabIndex        =   32
            Top             =   1260
            Width           =   465
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   14
            Left            =   2820
            TabIndex        =   31
            Top             =   1350
            Width           =   420
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   5
            Left            =   1230
            Picture         =   "frmTRAContaFac.frx":000C
            ToolTipText     =   "Buscar fecha"
            Top             =   1290
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   6
            Left            =   3360
            Picture         =   "frmTRAContaFac.frx":0097
            ToolTipText     =   "Buscar fecha"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Desde"
            Height          =   195
            Index           =   13
            Left            =   615
            TabIndex        =   30
            Top             =   630
            Width           =   465
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   12
            Left            =   2790
            TabIndex        =   29
            Top             =   630
            Width           =   420
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Factura"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   11
            Left            =   255
            TabIndex        =   28
            Top             =   390
            Width           =   555
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo de Factura"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   27
            Top             =   1740
            Width           =   1245
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos para la contabilización"
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
         Height          =   2715
         Left            =   90
         TabIndex        =   2
         Top             =   2460
         Width           =   6075
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   10
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   26
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   2220
            Width           =   1125
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   10
            Left            =   3195
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   2220
            Width           =   2685
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   3195
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   1530
            Width           =   2685
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   9
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   24
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1530
            Width           =   1125
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   21
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   480
            Width           =   1140
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   23
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1170
            Width           =   1125
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   3195
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   1170
            Width           =   2685
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   22
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   810
            Width           =   1140
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   3195
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   1890
            Width           =   2685
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   25
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1890
            Width           =   1125
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   10
            Left            =   1710
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   2220
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Retención"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   15
            Top             =   2265
            Width           =   1395
         End
         Begin VB.Label Label1 
            Caption         =   "F.Pago Negativas"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   13
            Top             =   1575
            Width           =   1485
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   9
            Left            =   1710
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   1530
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Recepción"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   11
            Top             =   480
            Width           =   1485
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   0
            Left            =   1710
            Picture         =   "frmTRAContaFac.frx":0122
            ToolTipText     =   "Buscar fecha"
            Top             =   480
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   1710
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   1170
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "F.Pago Positivas"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   10
            Top             =   1215
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Vencimiento"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   8
            Top             =   855
            Width           =   1425
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   1
            Left            =   1710
            Picture         =   "frmTRAContaFac.frx":01AD
            ToolTipText     =   "Buscar fecha"
            Top             =   810
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cta.Banco Prevista"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   24
            Left            =   180
            TabIndex        =   4
            Top             =   1935
            Width           =   1395
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   1710
            ToolTipText     =   "Buscar Cuenta Contable"
            Top             =   1890
            Width           =   240
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5100
         TabIndex        =   35
         Top             =   6240
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3915
         TabIndex        =   34
         Top             =   6240
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   5280
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   7
         Top             =   5610
         Width           =   5940
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   90
         TabIndex        =   6
         Top             =   5880
         Width           =   5925
      End
   End
End
Attribute VB_Name = "frmTRAContaFac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MONICA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto


Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

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
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String

Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

Dim vSeccion As CSeccion
Dim Tipo As Byte

Dim cContaFra As cContabilizarFacturas


Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim Sql As String
Dim i As Byte
Dim cadWHERE As String
Dim cDesde As String
Dim cHasta As String

    If Not DatosOk Then Exit Sub

    cadSelect = "rfacttra.contabilizado=0 "
    cadSelect = cadSelect & " and rfacttra.codtipom = " & DBSet(Mid(Combo1(0).Text, 1, 3), "T")

    'D/H Fecha factura
    cDesde = Trim(txtCodigo(5).Text)
    cHasta = Trim(txtCodigo(6).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rfacttra.fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHfechaFac= """) Then Exit Sub
    End If

    'D/H numero de factura
    cDesde = Trim(txtCodigo(7).Text)
    cHasta = Trim(txtCodigo(8).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rfacttra.numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHColec= """) Then Exit Sub
    End If

    If Not HayRegParaInforme("rfacttra", cadSelect) Then Exit Sub

    ContabilizarFacturas "rfacttra", cadSelect
     'Eliminar la tabla TMP
    BorrarTMPFacturas
    'Desbloqueamos ya no estamos contabilizando facturas
    DesBloqueoManual ("CONTRA") 'CONtabilizar facturas TRAnsporte
    
eError:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "No se ha realizado el proceso de contabilización de facturas de transporte. Llame a soporte."
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
        PonerFoco txtCodigo(7)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection
Dim i As Integer

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
    For i = 3 To 4
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 9 To 10
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    ConexionConta
    
    ' formas de pago
    txtCodigo(3).Text = Format(vParamAplic.ForpaPosi, "000")
    If vParamAplic.ContabilidadNueva Then
        txtNombre(3).Text = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtCodigo(3).Text, "N")
        txtCodigo(9).Text = Format(vParamAplic.ForpaNega, "000")
        txtNombre(9).Text = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtCodigo(9).Text, "N")
    Else
        txtNombre(3).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtCodigo(3).Text, "N")
        txtCodigo(9).Text = Format(vParamAplic.ForpaNega, "000")
        txtNombre(9).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtCodigo(9).Text, "N")
    End If
    ' cuentas contables
    txtCodigo(4).Text = vParamAplic.CtaBancoSoc   ' cuenta contable de banco prevista
    txtNombre(4).Text = PonerNombreCuenta(txtCodigo(4), 0)
    txtCodigo(10).Text = vParamAplic.CtaTraReten ' cuenta contable de retencion
    txtNombre(10).Text = PonerNombreCuenta(txtCodigo(10), 0)
    
    txtCodigo(5).Text = Format(Now, "dd/mm/yyyy") ' fecha de factura desde
    txtCodigo(6).Text = Format(Now, "dd/mm/yyyy") ' fecha de factura hasta
    txtCodigo(1).Text = Format(Now, "dd/mm/yyyy") ' fecha de vencimiento
    txtCodigo(0).Text = Format(Now, "dd/mm/yyyy") ' fecha de recepcion
            
    '###Descomentar
'    CommitConexion

    FrameCobrosVisible True, H, W
    Pb1.visible = False

    CargaCombo

    PosicionarCombo Me.Combo1(0), 0
    
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

Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtCodigo(indCodigo).Text = Format(txtCodigo(indCodigo).Text, "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
    ConexionConta
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

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 2 ' seccion
            AbrirFrmSeccion (Index)
        Case 3, 9 ' forma de pago de la tesoreria
            AbrirFrmForpaConta (Index)
        Case 4 'cuenta contable banco
            AbrirFrmCuentas (Index)
        Case 10, 11 ' cuentas contables de retnecion y de aportacion
            AbrirFrmCuentas (Index)
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub Optcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.CmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub OptNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.CmdAceptar
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
            Case 5: KEYFecha KeyAscii, 2 'fecha desde factura
            Case 6: KEYFecha KeyAscii, 3 'fecha hasta factura
            Case 1: KEYFecha KeyAscii, 1 'fecha vencimiento
            Case 4: KEYBusqueda KeyAscii, 4 'cta contable banco
            Case 10: KEYBusqueda KeyAscii, 10 'cta contable retencion
            Case 11: KEYBusqueda KeyAscii, 11 'cta contable aportacion
            Case 3: KEYBusqueda KeyAscii, 3 'forma de pago positivas
            Case 9: KEYBusqueda KeyAscii, 9 'forma de pago negativas
            Case 0: KEYFecha KeyAscii, 0 'fecha de recepcion
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
Dim Cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


    Select Case Index
        Case 2 ' SECCION
            If txtCodigo(Index).Text <> "" Then
                PonerFormatoEntero txtCodigo(Index)
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rseccion", "nomsecci", "codsecci", "N")
                ConexionConta
            Else
                Cad = "Debe introducir obligatoriamente una sección. " & vbCrLf & vbCrLf & "     ¿ Desea continuar ?"
                If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then cmdCancel_Click
            End If

        Case 3, 9 ' FORMAS DE PAGO DE LA CONTABILIDAD(POSITIVAS Y NEGATIVAS)
            If vSeccion Is Nothing Then Exit Sub
            If vParamAplic.ContabilidadNueva Then
                If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtCodigo(Index).Text, "N")
            Else
                If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtCodigo(Index).Text, "N")
            End If
            If txtNombre(Index).Text = "" Then
                MsgBox "Forma de Pago  no existe en la contabilidad. Reintroduzca.", vbExclamation
            End If

        Case 4, 10, 11 ' CUENTAS CONTABLES ( banco, retencion y aportacion )
            If vSeccion Is Nothing Then Exit Sub
        
            If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = PonerNombreCuenta(txtCodigo(Index), 2)
            If txtNombre(Index).Text = "" Then
                MsgBox "Número de Cuenta contable no existe en la contabilidad. Reintroduzca.", vbExclamation
            End If

        Case 5, 6 'FECHAS
            If txtCodigo(Index).Text <> "" Then
                If PonerFormatoFecha(txtCodigo(Index)) Then
                    If Index = 5 Then
                        txtCodigo(6).Text = txtCodigo(5).Text
                    End If
                End If
            End If

        Case 0, 1 'FECHAS de vencimiento
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

Private Sub AbrirFrmCuentas(indice As Integer)
    indCodigo = indice
    Set frmCtas = New frmCtasConta
    frmCtas.DatosADevolverBusqueda = "0|1|"
    frmCtas.CodigoActual = txtCodigo(indCodigo)
'    frmCtas.Conexion = cContaFacSoc
'    frmCtas.Facturas = False
    frmCtas.Show vbModal
    Set frmCtas = Nothing
End Sub

Private Sub AbrirFrmForpaConta(indice As Integer)
    indCodigo = indice
    Set frmFPa = New frmForpaConta
    frmFPa.DatosADevolverBusqueda = "0|1|"
    frmFPa.CodigoActual = txtCodigo(indCodigo)
'    frmFpa.Conexion = cContaFacSoc
    frmFPa.Show vbModal
    Set frmFPa = Nothing
End Sub

Private Sub AbrirFrmSeccion(indice As Integer)
    indCodigo = indice
    Set frmSec = New frmManSeccion
    frmSec.DatosADevolverBusqueda = "0|1|"
    frmSec.CodigoActual = txtCodigo(indCodigo)
'    frmSec.Facturas = False
    frmSec.Show vbModal
    Set frmSec = Nothing
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Orden1 As String
Dim Orden2 As String
Dim FFin As Date
Dim cta As String

   b = True

   If txtCodigo(6).Text = "" Then
        MsgBox "Introduzca la Fecha de Factura a contabilizar.", vbExclamation
        b = False
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
            MsgBox "La Fecha de la contabilización no es del ejercicio actual ni del siguiente. Reintroduzca.", vbExclamation
            b = False
            PonerFoco txtCodigo(6)
         End If
   End If

   If txtCodigo(0).Text = "" And b Then
        MsgBox "Introduzca la Fecha de Recepción de Factura.", vbExclamation
        b = False
        PonerFoco txtCodigo(0)
   End If

   If txtCodigo(1).Text = "" And b Then
        MsgBox "Introduzca la Fecha de Vencimiento a contabilizar.", vbExclamation
        b = False
        PonerFoco txtCodigo(1)
   End If

   If txtCodigo(3).Text = "" And b Then
        MsgBox "Introduzca la Forma de Pago para contabilizar.", vbExclamation
        b = False
        PonerFoco txtCodigo(3)
   End If

   'cta contable de banco
   If b Then
        If txtCodigo(4).Text = "" Then
             MsgBox "Introduzca la Cta.Contable de Banco para contabilizar.", vbExclamation
             b = False
             PonerFoco txtCodigo(4)
        Else
             cta = ""
             cta = DevuelveDesdeBDNew(cConta, "cuentas", "codmacta", "codmacta", txtCodigo(4).Text, "T")
             If cta = "" Then
                 MsgBox "La cuenta contable de Banco no existe. Reintroduzca.", vbExclamation
                 b = False
                 PonerFoco txtCodigo(4)
             End If
        End If
    End If
   
   'cta contable de retencion
   If b Then
        If txtCodigo(10).Text = "" Then
             MsgBox "Introduzca la Cta.Contable de Retención para contabilizar.", vbExclamation
             b = False
             PonerFoco txtCodigo(10)
        Else
             cta = ""
             cta = DevuelveDesdeBDNew(cConta, "cuentas", "codmacta", "codmacta", txtCodigo(10).Text, "T")
             If cta = "" Then
                 MsgBox "La cuenta contable de Retención no existe. Reintroduzca.", vbExclamation
                 b = False
                 PonerFoco txtCodigo(10)
             End If
        End If
    End If
   
   'cta contable de aportacion
'   If b Then
'        If txtcodigo(11).Text = "" Then
'             MsgBox "Introduzca la Cta.Contable de Aportación para contabilizar.", vbExclamation
'             b = False
'             PonerFoco txtcodigo(11)
'        Else
'             cta = ""
'             cta = DevuelveDesdeBDNew(cConta, "cuentas", "codmacta", "codmacta", txtcodigo(11).Text, "T")
'             If cta = "" Then
'                 MsgBox "La cuenta contable de Aportación no existe. Reintroduzca.", vbExclamation
'                 b = False
'                 PonerFoco txtcodigo(11)
'             End If
'        End If
'    End If
   
   'forma de pago positivas
   If b Then
        If txtCodigo(3).Text = "" Then
             MsgBox "Introduzca la Forma de Pago para facturas positivas para contabilizar.", vbExclamation
             b = False
             PonerFoco txtCodigo(3)
        Else
             cta = ""
             If vParamAplic.ContabilidadNueva Then
                cta = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtCodigo(3).Text, "T")
             Else
                cta = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtCodigo(3).Text, "T")
             End If
             If cta = "" Then
                 MsgBox "La Forma de Pago para facturas positivas no existe. Reintroduzca.", vbExclamation
                 b = False
                 PonerFoco txtCodigo(3)
             End If
        End If
    End If
   
   'forma de pago negativas
   If b Then
        If txtCodigo(9).Text = "" Then
             MsgBox "Introduzca la Forma de Pago para facturas negativas para contabilizar.", vbExclamation
             b = False
             PonerFoco txtCodigo(9)
        Else
             cta = ""
             If vParamAplic.ContabilidadNueva Then
                 cta = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtCodigo(9).Text, "T")
             Else
                 cta = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtCodigo(9).Text, "T")
             End If
             If cta = "" Then
                 MsgBox "La Forma de Pago para facturas negativas no existe. Reintroduzca.", vbExclamation
                 b = False
                 PonerFoco txtCodigo(9)
             End If
        End If
   End If
   
   DatosOk = b

End Function



Private Sub ContabilizarFacturas(cadTabla As String, cadWHERE As String)
'Contabiliza Facturas de Clientes o de Proveedores
Dim Sql As String
Dim b As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String



    Sql = "CONTRA" 'contabilizar facturas de transporte
    
    
    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se pueden Contabilizar Facturas de Transporte. Hay otro usuario contabilizando.", vbExclamation
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
     'contabilidad par ello mirar en la BD de la Conta los parámetros
     If Not ComprobarFechasConta(6) Then Exit Sub

    'comprobar si existen  facturas anteriores al periodo solicitado
    'sin contabilizar
    If Me.txtCodigo(5).Text <> "" Then 'anteriores a fechadesde
        Sql = "SELECT COUNT(*) FROM " & cadTabla
        Sql = Sql & " WHERE fecfactu <"
        Sql = Sql & DBSet(txtCodigo(5), "F") & " AND contabilizado=0 and codtipom = " & DBSet(Mid(Combo1(0).Text, 1, 3), "T")
        If RegistrosAListar(Sql) > 0 Then
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
    b = CrearTMPFacturas(cadTabla, cadWHERE)
    If Not b Then Exit Sub
    

    'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
'    TerminaBloquear
    Sql = cadTabla & " INNER JOIN tmpFactu ON " & cadTabla
    Sql = Sql & ".codtipom=tmpFactu.codtipom AND "
    
    Sql = Sql & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu AND " & cadTabla & ".codtrans=tmpfactu.codtrans"
    If Not BloqueaRegistro(Sql, cadWHERE) Then
        MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


    'Visualizar la barra de Progreso
    Me.Pb1.visible = True
    Me.lblProgres(0).Caption = "Comprobaciones: "
    CargarProgres Me.Pb1, 100

' [Monica] 18/05/2010 de momento comentado lo de comprobar letra de serie

'    'comprobar que todas las LETRAS SERIE existen en la contabilidad y en Ariagrorec
'    '--------------------------------------------------------------------------
'    Me.lblProgres(1).Caption = "Comprobando letras de serie ..."
'    b = ComprobarLetraSerie(cadTABLA)
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
'    If Not b Then Exit Sub


    'comprobar que todas las CUENTAS de los distintos socios que vamos a
    'contabilizar existen en la Conta: rsocios_seccion.codmacpro IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Contables transportistas en contabilidad ..."
    
    b = ComprobarCtaContable_new(cadTabla, 1, , CInt(vParamAplic.Seccionhorto))
    IncrementarProgres Me.Pb1, 30
    Me.Refresh
    If Not b Then Exit Sub


    'comprobar que todas las CUENTAS de anticipos/liquidaciones de las variedades que vamos a
    'contabilizar existen en la Conta: variedades.ctaanticipo o variedades.ctaliquidacion IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Contables variedades en contabilidad ..."
    
            
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    If Not b Then Exit Sub
    
    Tipo = DevuelveValor("select tipodocu from usuarios.stipom where codtipom = " & DBSet(Mid(Combo1(0).Text, 1, 3), "T"))
    
    '[Monica] 30/03/2010 en el caso de rectificativas hablar con Manolo
    If Tipo <> 11 Then ' solo si no son rectificativas
        b = ComprobarCtaContable_new(cadTabla, 8, Tipo, CInt(vParamAplic.Seccionhorto))
    End If
    
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    If Not b Then Exit Sub

'    'comprobar que todas las CUENTAS de gastos a pie de factura
'    b = ComprobarCtaContable_new(cadTABLA, 12, tipo, CInt(txtcodigo(2).Text))
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    If Not b Then Exit Sub

    'comprobar que todos las TIPO IVA de las distintas fecturas que vamos a
    'contabilizar existen en la Conta: rfactsoc.tipoiva IN (conta.tiposiva.codigiva)
    '--------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Tipos de IVA en contabilidad ..."
    b = ComprobarIVA(cadTabla)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    If Not b Then Exit Sub


    'comprobar si hay contabilidad ANALITICA: conta.parametros.autocoste=1
    'y verificar que las cuentas de sfamia.ctaventa empiezan por el digito
    'de conta.parametros.grupogto o conta.parametros.grupovta
    'obtener el centro de coste del usuario para insertarlo en linfact
    If vEmpresa.TieneAnalitica Then  'hay contab. analitica
       Me.lblProgres(1).Caption = "Comprobando Contabilidad Analítica ..."
           
       b = ComprobarCtaContable_new(cadTabla, 7, Tipo, CInt(vParamAplic.Seccionhorto))

       '(si tiene analítica requiere un centro de coste para insertar en conta.linfact)
       If b Then
            CCoste = ""
            b = ComprobarCCoste_new(CCoste, cadTabla)
            If Not b Then Exit Sub
       End If
       CCoste = ""
    End If
    IncrementarProgres Me.Pb1, 20
    Me.Refresh


    '===========================================================================
    'CONTABILIZAR FACTURAS
    '===========================================================================
    Me.lblProgres(0).Caption = "Contabilizar Facturas Transportistas: "
    CargarProgres Me.Pb1, 10
    Me.lblProgres(1).Caption = "Insertando Facturas en Contabilidad..."


    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    LOG.Insertar 3, vUsu, "Contabilizar facturas Transportistas: " & vbCrLf & cadTabla & vbCrLf & cadWHERE
    Set LOG = Nothing
    '-----------------------------------------------------------------------------


    '---- Crear tabla TEMP para los posible errores de facturas
    tmpErrores = CrearTMPErrFact(cadTabla)

    '---- Pasar las Facturas a la Contabilidad
    b = PasarFacturasAContab(cadTabla, CCoste)

    '---- Mostrar ListView de posibles errores (si hay)
    If Not b Then
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
        cadTitulo = "Listado contabilizacion FRATRA"
        cadNomRPT = "rContabTRA.rpt"
        
        LlamarImprimir
    End If


    '---- Eliminar tabla TEMP de Errores
    BorrarTMPErrFact

End Sub

Private Function PasarFacturasAContab(cadTabla As String, CCoste As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim b As Boolean
Dim i As Integer
Dim numfactu As Integer
Dim Codigo1 As String

    On Error GoTo EPasarFac

    PasarFacturasAContab = False

    '---- Obtener el total de Facturas a Insertar en la contabilidad
    Sql = "SELECT count(*) "
    Sql = Sql & " FROM " & cadTabla & " INNER JOIN tmpFactu "
    Codigo1 = "codtipom"
    Sql = Sql & " ON " & cadTabla & "." & Codigo1 & "=tmpFactu." & Codigo1
    Sql = Sql & " AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "
    Sql = Sql & " and " & cadTabla & ".codtrans=tmpfactu.codtrans"

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
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
    Sql = "DELETE FROM tmpinformes WHERE codusu = " & vUsu.Codigo
    conn.Execute Sql

    Set cContaFra = New cContabilizarFacturas
    
    If Not cContaFra.EstablecerValoresInciales(ConnConta) Then
        'NO ha establcedio los valores de la conta.  Le dejaremos seguir, avisando que
        ' obviamente, no va a contabilizar las FRAS
        Sql = "Si continua, las facturas se insertaran en el registro, pero no serán contabilizadas" & vbCrLf
        Sql = Sql & "en este momento. Deberán ser contabilizadas desde el ARICONTA" & vbCrLf & vbCrLf
        Sql = Sql & Space(50) & "¿Continuar?"
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    End If

    '---- Pasar cada una de las facturas seleccionadas a la Conta
    If numfactu > 0 Then
        CargarProgres Me.Pb1, numfactu

        'seleccinar todas las facturas que hemos insertado en la temporal (las que vamos a contabilizar)
        Sql = "SELECT * "
        Sql = Sql & " FROM tmpFactu "

        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenStatic, adLockPessimistic, adCmdText
        i = 1

        b = True
        'pasar a contabilidad cada una de las facturas seleccionadas
        While Not Rs.EOF
            Sql = cadTabla & "." & Codigo1 & "=" & DBSet(Rs.Fields(0), "T") & " and numfactu=" & DBSet(Rs!numfactu, "T")
            Sql = Sql & " and fecfactu=" & DBSet(Rs!fecfactu, "F")
            Sql = Sql & " and " & cadTabla & ".codtrans=" & DBSet(Rs!codTrans, "T")
            
            If PasarFacturaTra(Sql, CCoste, Orden2, vParamAplic.Seccionhorto, Tipo, CDate(txtCodigo(0).Text), CDate(txtCodigo(1).Text), txtCodigo(3).Text, txtCodigo(9).Text, txtCodigo(4).Text, txtCodigo(10).Text, "", Mid(Combo1(0).Text, 1, 3), cContaFra) = False And b Then b = False

            '---- Laura 26/10/2006
            'Al pasar cada factura al hacer el commit desbloqueamos los registros
            'que teniamos bloqueados y los volvemos a bloquear
            'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
            Sql = cadTabla & " INNER JOIN tmpFactu ON " & cadTabla & "." & Codigo1 & "=tmpFactu." & Codigo1 & " AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "
            If Not BloqueaRegistro(Sql, cadTabla & "." & Codigo1 & "=tmpFactu." & Codigo1 & " AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu") Then
'                MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
'                Screen.MousePointer = vbDefault
'                Exit Sub
            End If
            '----

            IncrementarProgres Me.Pb1, 1
            Me.lblProgres(1).Caption = "Insertando Facturas en Contabilidad...   (" & i & " de " & numfactu & ")"
            Me.Refresh
            i = i + 1
            Rs.MoveNext
        Wend

        Rs.Close
        Set Rs = Nothing
    End If
    
    Set cContaFra = Nothing

EPasarFac:
    If Err.Number <> 0 Then b = False

    If b Then
        PasarFacturasAContab = True
    Else
        PasarFacturasAContab = False
    End If
End Function

Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim i As Byte

    'tipo de IRPF
    Combo1(0).AddItem "FTR-Transporte"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "FRS-Rectificativa"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1

End Sub

Private Function ComprobarFechasConta(ind As Integer) As Boolean
'comprobar que el periodo de fechas a contabilizar esta dentro del
'periodo de fechas del ejercicio de la contabilidad
Dim FechaIni As String, FechaFin As String
Dim Cad As String
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
                 Cad = "El período de contabilización debe estar dentro del ejercicio:" & vbCrLf & vbCrLf
                 Cad = Cad & "    Desde: " & FechaIni & vbCrLf
                 Cad = Cad & "    Hasta: " & FechaFin
                 MsgBox Cad, vbExclamation
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


Private Sub ConexionConta()
    
    If vSeccion Is Nothing Then
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
            vSeccion.AbrirConta
        End If
    Else
        ' si el objeto existia: cerramos la conexion y volvemos crearlo
        vSeccion.CerrarConta
        Set vSeccion = Nothing
        
        
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
            vSeccion.AbrirConta
        End If
    End If
End Sub
