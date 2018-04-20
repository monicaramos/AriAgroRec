VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPOZIntTesor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Integración en Contabilidad y Tesorería"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6750
   Icon            =   "frmPOZIntTesor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCobros 
      Height          =   5460
      Left            =   135
      TabIndex        =   9
      Top             =   120
      Width           =   6555
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
         Height          =   1950
         Left            =   90
         TabIndex        =   11
         Top             =   1890
         Width           =   6345
         Begin VB.TextBox txtcodigo 
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
            Left            =   2385
            MaxLength       =   40
            TabIndex        =   6
            Text            =   "1234567890123456789012345678901234567890"
            Top             =   1470
            Width           =   3840
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
            Left            =   2385
            MaxLength       =   10
            TabIndex        =   5
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   1110
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
            Left            =   3735
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   1110
            Width           =   2505
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
            Index           =   3
            Left            =   2385
            MaxLength       =   10
            TabIndex        =   4
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   735
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
            Index           =   3
            Left            =   3285
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   735
            Width           =   2955
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
            Index           =   2
            Left            =   2385
            MaxLength       =   10
            TabIndex        =   3
            Tag             =   "Código Postal|T|S|||clientes|codposta|||"
            Top             =   345
            Width           =   1350
         End
         Begin VB.Label Label1 
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
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   24
            Top             =   1515
            Width           =   1530
         End
         Begin VB.Image imgAyuda 
            Height          =   240
            Index           =   0
            Left            =   2115
            MousePointer    =   4  'Icon
            Tag             =   "-1"
            ToolTipText     =   "Ayuda"
            Top             =   1470
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   2115
            ToolTipText     =   "Buscar cuenta"
            Top             =   1110
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
            TabIndex        =   22
            Top             =   1155
            Width           =   1935
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   2115
            ToolTipText     =   "Buscar f.pago"
            Top             =   735
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma de Pago"
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
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   17
            Top             =   780
            Width           =   1620
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
            TabIndex        =   15
            Top             =   390
            Width           =   1920
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   2
            Left            =   2115
            Picture         =   "frmPOZIntTesor.frx":000C
            ToolTipText     =   "Buscar fecha"
            Top             =   345
            Width           =   240
         End
      End
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
         Height          =   1590
         Left            =   90
         TabIndex        =   10
         Top             =   225
         Width           =   6360
         Begin VB.ComboBox Combo1 
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
            Left            =   3810
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   690
            Width           =   2385
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
            Index           =   0
            Left            =   2385
            MaxLength       =   10
            TabIndex        =   0
            Top             =   690
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
            Index           =   1
            Left            =   2385
            MaxLength       =   10
            TabIndex        =   1
            Top             =   1050
            Width           =   1350
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo de Factura"
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
            Left            =   3810
            TabIndex        =   23
            Top             =   450
            Width           =   1815
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   1
            Left            =   2115
            Picture         =   "frmPOZIntTesor.frx":0097
            ToolTipText     =   "Buscar fecha"
            Top             =   1035
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   0
            Left            =   2115
            Picture         =   "frmPOZIntTesor.frx":0122
            ToolTipText     =   "Buscar fecha"
            Top             =   690
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
            Left            =   1395
            TabIndex        =   20
            Top             =   1050
            Width           =   690
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
            Left            =   1395
            TabIndex        =   19
            Top             =   690
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Recibo"
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
            Index           =   1
            Left            =   225
            TabIndex        =   18
            Top             =   450
            Width           =   1815
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
         Height          =   375
         Left            =   5310
         TabIndex        =   8
         Top             =   4815
         Width           =   1065
      End
      Begin VB.CommandButton CmdAceptar 
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
         Left            =   4125
         TabIndex        =   7
         Top             =   4815
         Width           =   1065
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   3855
         Width           =   6270
         _ExtentX        =   11060
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   0
         Left            =   495
         TabIndex        =   14
         Top             =   4140
         Width           =   5265
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   495
         TabIndex        =   13
         Top             =   4410
         Width           =   5295
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPOZIntTesor"
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
Private WithEvents frmFPa As frmForpaConta 'formas de pago de la contabilidad
Attribute frmFPa.VB_VarHelpID = -1
Private WithEvents frmCtas As frmCtasConta 'cuentas contables de contabilidad
Attribute frmCtas.VB_VarHelpID = -1

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

Dim PrimeraVez As Boolean

Dim vSeccion As CSeccion
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report


Dim Salir As Boolean
Dim cContaFra As cContabilizarFacturas


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click()
Dim Sql As String
Dim I As Byte
Dim cadWHERE As String

    If Not DatosOK Then Exit Sub
             
    Sql = "SELECT count(*)" & _
          " FROM rrecibpozos " & _
          "WHERE "
          
    cadWHERE = "contabilizado = 0"
          
    If txtCodigo(0).Text <> "" Then cadWHERE = cadWHERE & " and rrecibpozos.fecfactu >= " & DBSet(txtCodigo(0).Text, "F")
    If txtCodigo(1).Text <> "" Then cadWHERE = cadWHERE & " and rrecibpozos.fecfactu <= " & DBSet(txtCodigo(1).Text, "F")
             
    Sql = Sql & cadWHERE
    
    ' dependiendo del tipo de recibo ponemos el tipo de movimiento
    Select Case Mid(Combo1(0).Text, 1, 3) ' antes Combo1(0).ListIndex
        Case "RCP" '0 contadores
            cadWHERE = cadWHERE & " and rrecibpozos.codtipom = 'RCP'"
        Case "RMP" '1 mantenimiento
            cadWHERE = cadWHERE & " and rrecibpozos.codtipom = 'RMP'"
        Case "RVP" '2
            cadWHERE = cadWHERE & " and rrecibpozos.codtipom = 'RVP'"
        Case "TAL" '3
            cadWHERE = cadWHERE & " and rrecibpozos.codtipom = 'TAL'"
        Case "RMT" '4
            cadWHERE = cadWHERE & " and rrecibpozos.codtipom = 'RMT'"
        '[Monica]14/01/2016: la rectificativas
        Case "RRC" '5
            cadWHERE = cadWHERE & " and rrecibpozos.codtipom = 'RRC'"
        Case "RRM" '6
            cadWHERE = cadWHERE & " and rrecibpozos.codtipom = 'RRM'"
        Case "RRV" '7
            cadWHERE = cadWHERE & " and rrecibpozos.codtipom = 'RRV'"
        Case "RTA" '8
            cadWHERE = cadWHERE & " and rrecibpozos.codtipom = 'RTA'"
'        Case "RRT" '9
'            cadwhere = cadwhere & " and rrecibpozos.codtipom = 'RRT' "
        Case "FIN" '0 internas
            cadWHERE = cadWHERE & " and rrecibpozos.codtipom = 'FIN'"
            
    End Select
             
             
    If RegistrosAListar(Sql) = 0 Then
        MsgBox "No existen datos a contabilizar entre esas fechas.", vbExclamation
        Exit Sub
    End If
    
    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    LOG.Insertar 3, vUsu, "Contabilizar Recibos Pozos: " & vbCrLf & "rrecibpozos" & vbCrLf & cadWHERE
    Set LOG = Nothing
    '-----------------------------------------------------------------------------

    '[Monica]15/10/2012: Ahora Escalona inserta en el registro de iva
    If vParamAplic.Cooperativa = 1 Then 'Or vParamAplic.Cooperativa = 10 Then ' Turis no tiene contabilizacion de facturas Escalona tampoco
        ContabilizarCobros (cadWHERE)
    Else
        ContabilizarFacturas "rrecibpozos", cadWHERE
    End If
    
    BorrarTMPErrComprob
    DesBloqueoManual ("CONTES") 'CONtabilizacion a TESoreria
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    '[Monica]19/01/2016: solo en el caso de escalona no salimos de la integracion contable
    If vParamAplic.Cooperativa <> 10 Then cmdCancel_Click
    
    '[Monica]21/01/2016: cuando vuelvo refrescar el combo
    CargaCombo
    If Combo1(0).ListCount > 0 Then Combo1(0).ListIndex = 0
    
eError:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "No se ha realizado el proceso de contabilización a tesoreria. Llame a soporte."
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    
    If PrimeraVez Then
        PrimeraVez = False
        
        txtCodigo(3).Text = vParamAplic.ForpaRecPOZ
        txtCodigo_LostFocus (3)
        
        PonerFoco txtCodigo(0)
        '[Monica]21/01/2016: si hay recibos para contabilizar situamos el combo en el primero
        If Combo1(0).ListCount > 0 Then
            Combo1(0).ListIndex = 0
        End If
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer, I As Integer
Dim List As Collection
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
    Me.imgBuscar(3).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgBuscar(4).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    txtCodigo(2).Text = Format(Now, "dd/mm/yyyy") ' fecha de vencimiento
    
    For I = 0 To imgAyuda.Count - 1
        imgAyuda(I).Picture = frmPpal.ImageListB.ListImages(10).Picture
    Next I

    ConexionConta
         
    FrameCobrosVisible True, H, W
    Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    
    CargaCombo
    If vParamAplic.Cooperativa = 7 Then
        '[Monica]21/01/2016: si hay recibos para contabilizar situamos el combo en el primero
        If Combo1(0).ListCount > 0 Then Combo1(0).ListIndex = 0
        Combo1(0).Enabled = False
    
        '[Monica]08/05/2012: campo de observaciones del registro de iva de contabilidad
        txtCodigo(5).Text = "Consumo Agua Riegos"
    End If
    Salir = False
    If Combo1(0).ListCount = 0 Then
        MsgBox "No hay recibos pendientes de integrar.", vbExclamation
        Salir = True
    End If

    
    
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
    txtCodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Cuentas contables
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFPa_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de formas de pago de contabilidad
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "Concepto que se graba en el registro de Iva de Cliente " & vbCrLf & _
                      "en la Contabilidad, sólo en caso de contabilizar facturas" & vbCrLf & _
                      "además de los cobros en tesoreria." & vbCrLf & vbCrLf
                                            
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
    
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
    imgFec(0).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index).Text <> "" Then frmC.NovaData = txtCodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(0).Tag) + 1)
    ' ***************************
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 3 ' forma de pago de la tesoreria
            AbrirFrmForpaConta (Index)
        Case 4 'cuenta contable
            AbrirFrmCuentas (Index)
        
    End Select
    PonerFoco txtCodigo(indCodigo)
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
            Case 3: KEYBusqueda KeyAscii, 3 'forma de pago
            Case 0: KEYFecha KeyAscii, 0 'fecha desde
            Case 1: KEYFecha KeyAscii, 1 'fecha hasta
            Case 2: KEYFecha KeyAscii, 2 'fecha de vencimiento
            Case 4: KEYBusqueda KeyAscii, 4 'cuenta banco
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
        Case 3 ' FORMA DE PAGO DE LA CONTABILIDAD
            If vParamAplic.ContabilidadNueva Then
                If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = DevuelveDesdeBDNew(cConta, "formapago", "nomforpa", "codforpa", txtCodigo(3).Text, "N")
            Else
                If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtCodigo(3).Text, "N")
            End If
            If txtNombre(Index).Text = "" Then
                MsgBox "Forma de Pago  no existe en la contabilidad. Reintroduzca.", vbExclamation
            End If
            
        Case 0, 1, 2 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 4 ' CUENTA CONTABLE
            If txtCodigo(Index).Text <> "" Then txtNombre(Index).Text = PonerNombreCuenta(txtCodigo(Index), 2)
            If txtNombre(Index).Text = "" Then
                MsgBox "Número de Cuenta contable no existe en la contabilidad. Reintroduzca.", vbExclamation
            End If

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

 
Private Sub AbrirFrmForpaConta(Indice As Integer)
    indCodigo = Indice
    Set frmFPa = New frmForpaConta
    frmFPa.DatosADevolverBusqueda = "0|1|"
    frmFPa.CodigoActual = txtCodigo(Indice)
'    frmFpa.Conexion = cContaFacSoc
    frmFPa.Show vbModal
    Set frmFPa = Nothing
End Sub



Private Sub AbrirFrmCuentas(Indice As Integer)
    indCodigo = Indice
    Set frmCtas = New frmCtasConta
    frmCtas.DatosADevolverBusqueda = "0|1|"
    frmCtas.CodigoActual = txtCodigo(indCodigo)
    frmCtas.Show vbModal
    Set frmCtas = Nothing
End Sub

Private Sub ContabilizarCobros(cadWHERE As String)
'Contabiliza Facturas de Clientes o de Proveedores
Dim Sql As String
Dim B As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String
Dim cadTabla As String

    Sql = "CONTES" 'contabilizar tesoreria

    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se pueden Contabilizar Cobros. Hay otro usuario contabilizándolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    
    'Visualizar la barra de Progreso
    Me.Pb1.visible = True
'    Me.Pb1.Top = 3350
    
    
    '==========================================================
    'REALIZAR COMPROBACIONES ANTES DE CONTABILIZAR FACTURAS
    '==========================================================
    
    Me.lblProgres(0).Caption = "Comprobaciones: "
    CargarProgres Me.Pb1, 100
        
    BorrarTMPFacturas
    ' nuevo
    B = CrearTMPFacturas("rrecibpozos", cadWHERE)
    If Not B Then Exit Sub
    
    
    'comprobar que todas las CUENTAS de socios existe
    'en la Conta: rsocios.codmacta IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuenta Ctble de Socios ..."
    B = ComprobarCtaContable_new("rrecibpozos", 1)
    IncrementarProgres Me.Pb1, 100
    Me.Refresh
    DoEvents
    
    If Not B Then Exit Sub
   
    '===========================================================================
    'CONTABILIZAR CIERRE
    '===========================================================================
    Me.lblProgres(0).Caption = "Contabilizar a Tesorería: "
    CargarProgres Me.Pb1, 10
    Me.lblProgres(1).Caption = "Insertando Registro en Tesorería..."
    
    
    B = PasarCalculoAContab(cadWHERE)
    
    If B Then
        MsgBox "El proceso ha finalizado correctamente.", vbInformation
    End If
    
End Sub

Private Function DatosOK() As Boolean
Dim B As Boolean
Dim Orden1 As String
Dim Orden2 As String
Dim FFin As Date
Dim Sql As String

   B = True


   If txtCodigo(2).Text = "" And B Then
        MsgBox "Introduzca la Fecha de Vencimiento a contabilizar.", vbExclamation
        B = False
        PonerFoco txtCodigo(2)
   End If
    
   If txtCodigo(3).Text = "" And B Then
        MsgBox "Introduzca la Forma de Pago para contabilizar.", vbExclamation
        B = False
        PonerFoco txtCodigo(3)
   Else
        ' comprobamos que existe la forma de pago en contabilidad
        If vParamAplic.ContabilidadNueva Then
            Sql = DevuelveDesdeBDNew(cConta, "formapago", "codforpa", "codforpa", txtCodigo(3).Text, "N")
        Else
            Sql = DevuelveDesdeBDNew(cConta, "sforpa", "codforpa", "codforpa", txtCodigo(3).Text, "N")
        End If
        If Sql = "" Then
            MsgBox "No existe la forma de pago en Contabilidad. Revise.", vbExclamation
            B = False
            PonerFoco txtCodigo(3)
        End If
   End If
   
   If txtCodigo(4).Text = "" And B Then
        MsgBox "Introduzca la Cta.Contable de Banco para contabilizar.", vbExclamation
        B = False
        PonerFoco txtCodigo(4)
   End If
   
   
   '[Monica]21/01/2016: comprobamos si es escalona que tenga la cuenta de recargo en parametros
   If vParamAplic.Cooperativa = 10 Then
        If vParamAplic.CtaRecargosPOZ = "" Then
            MsgBox "Debe de introducir la Cta.Contable de Recargos en parámetros. Revise.", vbExclamation
            B = False
        End If
   End If
   If Combo1(0).Text = "" And Combo1(0).ListCount > 0 Then
        MsgBox "Debe introducir un Tipo de Factura para integrar. Revise.", vbExclamation
        B = False
        PonerFocoCmb Combo1(0)
   End If
        
        
   
   DatosOK = B
   
End Function

Private Function PasarCalculoAContab(cadWHERE As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim B As Boolean
Dim I As Integer
Dim numlinea As Integer
Dim numdocum As String
Dim Ampliacion As String
Dim ampliaciond As String
Dim ampliacionh As String
Dim ImporteD As Currency
Dim ImporteH As Currency
Dim Diferencia As Currency
Dim Obs As String
Dim cadMen As String
Dim cad As String
Dim CtaDifer As String
Dim Codmacta As String

    On Error GoTo EPasarCal

    PasarCalculoAContab = False
    
    'Total de lineas de asiento a Insertar en la contabilidad
    
    Sql = "SELECT count(distinct codtipom, numfactu,fecfactu)" & _
          " FROM rrecibpozos " & _
          "WHERE " & cadWHERE
             
    numlinea = TotalRegistros(Sql)
    
    If numlinea = 0 Then Exit Function
    
    
    If numlinea > 0 Then
        numlinea = numlinea
        
        CargarProgres Me.Pb1, numlinea
        
        ConnConta.BeginTrans
        conn.BeginTrans
        
        Obs = "Contabilización de Cobro de Recibos de fecha " & Format(txtCodigo(0).Text, "dd/mm/yyyy")

        Sql = "select distinct codtipom, numfactu, fecfactu, codsocio from rrecibpozos where " & cadWHERE
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText


        B = True
        I = 1
        While Not Rs.EOF And B
                IncrementarProgres Me.Pb1, 1
                Me.lblProgres(1).Caption = "Insertando registro en Tesorería...   (" & I & " de " & numlinea & ")"
                Me.Refresh
                DoEvents
                
                I = I + 1
                cadMen = "Insertando en Tesoreria: "
                B = InsertarEnTesoreriaPOZOS(cadMen, Rs, CDate(txtCodigo(2).Text), CInt(txtCodigo(3).Text), txtCodigo(4).Text)
               
                Rs.MoveNext
        Wend
        Rs.Close
            
' de momento comentado para hacer pruebas
        If B Then
            'Poner intconta=1 en ariagroutil.movim
            B = ActualizarCobros(cadWHERE, cadMen)
            cadMen = "Actualizando Movimientos: " & cadMen
        End If
            
   End If
   
EPasarCal:
    If Err.Number <> 0 Or Not B Then
        B = False
        MuestraError Err.Number, "Integrando Recibos de Pozos a Contabilidad", cadMen & " " & Err.Description
    End If
    If B Then
        ConnConta.CommitTrans
        conn.CommitTrans
        PasarCalculoAContab = True
    Else
        ConnConta.RollbackTrans
        conn.RollbackTrans
        PasarCalculoAContab = False
    End If
End Function


Private Function ActualizarCobros(cadWHERE As String, cadErr As String) As Boolean
'Poner el movimiento como contabilizada
Dim Sql As String

    On Error GoTo EActualizar
    
    Sql = "UPDATE rrecibpozos SET contabilizado=1 "
    Sql = Sql & " WHERE " & cadWHERE

    conn.Execute Sql
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarCobros = False
        cadErr = Err.Description
    Else
        ActualizarCobros = True
    End If
End Function


Private Sub ConexionConta()
    
    If vSeccion Is Nothing Then
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.SeccionPOZOS) Then
            vSeccion.AbrirConta
        End If
    Else
        ' si el objeto existia: cerramos la conexion y volvemos crearlo
        vSeccion.CerrarConta
        Set vSeccion = Nothing
        
        
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.SeccionPOZOS) Then
            vSeccion.AbrirConta
        End If
    End If
End Sub

'#######################################################################
'#######    CONTABILIZAR FACTURAS PARA QUATRETONDA Y UTXERA ############
'#######################################################################

Private Sub ContabilizarFacturas(cadTabla As String, cadWHERE As String)
'Contabiliza Facturas de Clientes o de Proveedores
Dim Sql As String
Dim B As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String

    Sql = "CONTES" 'contabilizar recibos de pozos

    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se pueden Contabilizar Recibos de Pozos. Hay otro usuario contabilizando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


    'comprobar que se han rellenado los dos campos de fecha
    'sino rellenar con fechaini o fechafin del ejercicio
    'que guardamos en vbles Orden1,Orden2
    If txtCodigo(0).Text = "" Then
       txtCodigo(0).Text = vEmpresa.FechaIni  'fechaini del ejercicio de la conta
    End If

    If txtCodigo(1).Text = "" Then
       txtCodigo(1).Text = vEmpresa.FechaFin  'fecha fin del ejercicio de la conta
    End If


    'Comprobar que el intervalo de fechas D/H esta dentro del ejercicio de la
    'contabilidad par ello mirar en la BD de la Conta los parámetros
    If Not ComprobarFechasConta(0) Then Exit Sub

    'comprobar si existen en Ariagrorec facturas anteriores al periodo solicitado
    'sin contabilizar
    If Me.txtCodigo(0).Text <> "" Then 'anteriores a fechadesde
        Sql = "SELECT COUNT(*) FROM " & cadTabla
        Sql = Sql & " WHERE fecfactu <"
        Sql = Sql & DBSet(txtCodigo(0), "F") & " AND contabilizado=0 "
        '[Monica]21/01/2016: el combo es de las pendientes de contabilizar
        Select Case Mid(Combo1(0).Text, 1, 3) 'Combo1(0).ListIndex
            Case "RCP" '0
                Sql = Sql & " and codtipom = 'RCP'"
            Case "RMP" '1
                Sql = Sql & " and codtipom = 'RMP'"
            Case "RVP" '2
                Sql = Sql & " and codtipom = 'RVP'"
            Case "TAL" '3
                Sql = Sql & " and codtipom = 'TAL'"
            Case "RMT" '4
                Sql = Sql & " and codtipom = 'RMT'"
        
            '[Monica]14/01/2016: la rectificativas
            Case "RRC" '5
                Sql = Sql & " and codtipom = 'RRC'"
            Case "RRM" '6
                Sql = Sql & " and codtipom = 'RRM'"
            Case "RRV" '7
                Sql = Sql & " and codtipom = 'RRV'"
            Case "RTA" '8
                Sql = Sql & " and codtipom = 'RTA'"
'            Case "RRT" '9
'                SQL = SQL & " and codtipom = 'RRT' "
        
            Case "FIN"
                Sql = Sql & " and codtipom = 'FIN'"
        
        End Select
            
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
    BorrarTMPFacturas
    B = CrearTMPFacturas(cadTabla, cadWHERE)
    If Not B Then Exit Sub


    'Laura: 11/10/2006 bloquear los registros q vamos a contabilizar
'    TerminaBloquear
    Sql = cadTabla & " INNER JOIN tmpFactu ON " & cadTabla
    Sql = Sql & ".codtipom=tmpFactu.codtipom AND "
    Sql = Sql & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "

    If Not BloqueaRegistro(Sql, cadWHERE) Then
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
    DoEvents
    
    If Not B Then Exit Sub

    '[Monica]04/06/2014:
    'comprobar que todos los socios tengan registro en la rsocios_seccion
    Me.lblProgres(1).Caption = "Comprobando Registros en la seccion de Pozos ..."

    B = ComprobarSociosSeccion(cadTabla, vParamAplic.SeccionPOZOS)
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    DoEvents
    
    If Not B Then Exit Sub



    'comprobar que no haya Nº FACTURAS en la contabilidad para esa fecha
    'que ya existan
    '-----------------------------------------------------------------------
    If cadTabla = "rrecibpozos" Then
        Me.lblProgres(1).Caption = "Comprobando Nº Facturas en contabilidad ..."
        Sql = "anofaccl>=" & Year(txtCodigo(0).Text) & " AND anofaccl<= " & Year(txtCodigo(1).Text)
        B = ComprobarNumFacturas_new(cadTabla, Sql)
    End If
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    DoEvents
    
    If Not B Then Exit Sub


    'comprobar que todas las CUENTAS de los distintos socios que vamos a
    'contabilizar existen en la Conta: rsocios_seccion.codmaccli IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuentas Contables socios en contabilidad ..."

    B = ComprobarCtaContable_new(cadTabla, 1)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    DoEvents
    
    If Not B Then Exit Sub


    'comprobar que todas las CUENTAS de venta de parametros
    'contabilizar existen en la Conta: sfamia.ctaventa IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuenta Ctble Venta Consumo en contabilidad ..."
    B = ComprobarCtaContable_new(cadTabla, 2, 1)
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    DoEvents
    
    If Not B Then Exit Sub

    Me.lblProgres(1).Caption = "Comprobando Cuenta Ctble Venta Cuotas en contabilidad ..."
    B = ComprobarCtaContable_new(cadTabla, 2, 2)
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    DoEvents
    
    If Not B Then Exit Sub

    '[Monica]21/01/2016: comprobamos la cuenta de recargo (escalona)
    If vParamAplic.Cooperativa = 10 Then
        Me.lblProgres(1).Caption = "Comprobando Cuenta Ctble Recargos en contabilidad ..."
        B = ComprobarCtaContable_new(cadTabla, 2, 6)
        IncrementarProgres Me.Pb1, 10
        Me.Refresh
        DoEvents
        
        If Not B Then Exit Sub
    End If

    If vParamAplic.Cooperativa = 8 Or vParamAplic.Cooperativa = 10 Then
        '[Monica]27/06/2013: añadimos la contabilizacion de las facturas de contador para Utxera y Escalona
        
        '[Monica]21/01/2016: cambiamos el combo por los que quedan por facturar
        'If Combo1(0).ListIndex = 1 Or Combo1(0).ListIndex = 2 Or Combo1(0).ListIndex = 6 Or Combo1(0).ListIndex = 7 Then ' mantenimiento
        If Mid(Combo1(0).Text, 1, 3) = "RMP" Or Mid(Combo1(0).Text, 1, 3) = "RVP" Or Mid(Combo1(0).Text, 1, 3) = "RRM" Or Mid(Combo1(0).Text, 1, 3) = "RMV" Then
            Me.lblProgres(1).Caption = "Comprobando Cuenta Ctble Venta Mantenimiento en contabilidad ..."
            B = ComprobarCtaContable_new(cadTabla, 2, 4)
            IncrementarProgres Me.Pb1, 10
            Me.Refresh
            DoEvents
            
            If Not B Then Exit Sub
        End If
        
        '[Monica]21/01/2016: cambiamos el combo por los que quedan por facturar
        'If Combo1(0).ListIndex = 3 Or Combo1(0).ListIndex = 8 Then ' talla
        If Mid(Combo1(0).Text, 1, 3) = "TAL" Or Mid(Combo1(0).Text, 1, 3) = "RTA" Then
            Me.lblProgres(1).Caption = "Comprobando Cuenta Ctble Venta Talla en contabilidad ..."
            B = ComprobarCtaContable_new(cadTabla, 2, 3)
            IncrementarProgres Me.Pb1, 10
            Me.Refresh
            DoEvents
            
            If Not B Then Exit Sub
        End If
        
        '[Monica]21/01/2016: cambiamos el combo por los que quedan por facturar
        'If Combo1(0).ListIndex = 4 Or Combo1(0).ListIndex = 9 Then  ' recibo consumo a manta
        If Mid(Combo1(0).Text, 1, 3) = "RMT" Or Mid(Combo1(0).Text, 1, 3) = "RRT" Then
            Me.lblProgres(1).Caption = "Comprobando Cuenta Ctble Venta Manta en contabilidad ..."
            B = ComprobarCtaContable_new(cadTabla, 2, 5)
            IncrementarProgres Me.Pb1, 10
            Me.Refresh
            DoEvents
            
            If Not B Then Exit Sub
        End If
        
    End If


    'comprobar que todos las TIPO IVA de las distintas facturas que vamos a
    'contabilizar existen en la Conta: advfacturas.codiiva1 codiiva2 codiiva3 IN (conta.tiposiva.codigiva)
    '--------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Tipos de IVA en contabilidad ..."
    B = ComprobarTiposIVA(cadTabla)
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    DoEvents
    
    If Not B Then Exit Sub


    'comprobar si hay contabilidad ANALITICA: conta.parametros.autocoste=1
    'y verificar que las cuentas de rparamaplic.ctaventaalmz rparamaplic.ctagastosalmz
    'empiezan por el digito de conta.parametros.grupogto o conta.parametros.grupovta
    'obtener el centro de coste del usuario para insertarlo en linfact
    If vEmpresa.TieneAnalitica Then  'hay contab. analitica
       Me.lblProgres(1).Caption = "Comprobando Contabilidad Analítica ..."

       B = ComprobarCtaContable_new(cadTabla, 7)
       If B Then
            '(si tiene analítica requiere un centro de coste para insertar en conta.linfact)
            CCoste = ""
            B = ComprobarCCoste_new(CCoste, cadTabla)
       End If
       If Not B Then Exit Sub

       CCoste = ""
    End If
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    DoEvents
    

'    If b Then
'       Me.lblProgres(1).Caption = "Comprobando Forma de Pago ..."
'       b = ComprobarFormadePago(cadTABLA)
'       If Not b Then Exit Sub
'    End If
    IncrementarProgres Me.Pb1, 10
    Me.Refresh
    DoEvents
    


    '===========================================================================
    'CONTABILIZAR FACTURAS
    '===========================================================================
    Me.lblProgres(0).Caption = "Contabilizar Recibos Pozos: "
    CargarProgres Me.Pb1, 10
    Me.lblProgres(1).Caption = "Insertando Facturas en Contabilidad..."


'    '------------------------------------------------------------------------------
'    '  LOG de acciones
'    Set LOG = New cLOG
'    LOG.Insertar 3, vUsu, "Contabilizar Recibos Pozos: " & vbCrLf & cadTabla & vbCrLf & cadwhere
'    Set LOG = Nothing
'    '-----------------------------------------------------------------------------


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
        cadTitulo = "Listado contabilizacion FRAFAD"
        cadNomRPT = "rContabFAD.rpt"
        LlamarImprimir
    End If


    '---- Eliminar tabla TEMP de Errores
    BorrarTMPErrFact

End Sub

Private Function PasarFacturasAContab(cadTabla As String) As Boolean
Dim Sql As String
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
    Sql = "SELECT count(distinct tmpfactu.codtipom, tmpfactu.numfactu, tmpfactu.fecfactu) "
    Sql = Sql & " FROM " & cadTabla & " INNER JOIN tmpFactu "
    
    Codigo1 = "codtipom"
    Sql = Sql & " ON " & cadTabla & "." & Codigo1 & "=tmpFactu." & Codigo1
    Sql = Sql & " AND " & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "

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
        Sql = "SELECT distinct codtipom, numfactu, fecfactu  "
        Sql = Sql & " FROM tmpFactu "

        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenStatic, adLockPessimistic, adCmdText
        I = 1

        B = True
        
        ' de momento no tiene analitica
        CCoste = ""
        
        'pasar a contabilidad cada una de las facturas seleccionadas
        While Not Rs.EOF
            Sql = cadTabla & "." & Codigo1 & "=" & DBSet(Rs.Fields(0), "T") & " and numfactu=" & Rs!numfactu
            Sql = Sql & " and fecfactu=" & DBSet(Rs!fecfactu, "F")
            If PasarFacturaPOZOS(Sql, CCoste, txtCodigo(4).Text, txtCodigo(2).Text, Rs.Fields(0), Rs!fecfactu, txtCodigo(5).Text, txtCodigo(3).Text, cContaFra) = False And B Then B = False

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
            Me.lblProgres(1).Caption = "Insertando Facturas en Contabilidad...   (" & I & " de " & numfactu & ")"
            Me.Refresh
            DoEvents
            
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
                 cad = "El período de contabilización debe estar dentro del ejercicio:" & vbCrLf & vbCrLf
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


Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim I As Integer
    
    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For I = 0 To Combo1.Count - 1
        Combo1(I).Clear
    Next I
    
'    'tipo de fichero
'    Combo1(0).AddItem "RCP-Consumo"
'    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
'    Combo1(0).AddItem "RMP-Mantenimiento"
'    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
'    Combo1(0).AddItem "RVP-Contadores"
'    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
'    If vParamAplic.Cooperativa = 10 Then
'        Combo1(0).AddItem "TAL-Talla"
'        Combo1(0).ItemData(Combo1(0).NewIndex) = 3
'        Combo1(0).AddItem "RMT-Consumo Manta"
'        Combo1(0).ItemData(Combo1(0).NewIndex) = 4
'
'        '[Monica]14/01/2016: las rectificativas
'        Combo1(0).AddItem "RRC-Rect.Consumo"
'        Combo1(0).ItemData(Combo1(0).NewIndex) = 5
'        Combo1(0).AddItem "RRM-Rect.Mantenimiento"
'        Combo1(0).ItemData(Combo1(0).NewIndex) = 6
'        Combo1(0).AddItem "RRV-Rect.Contadores"
'        Combo1(0).ItemData(Combo1(0).NewIndex) = 7
'        Combo1(0).AddItem "RTA-Rect.Talla"
'        Combo1(0).ItemData(Combo1(0).NewIndex) = 8
''        Combo1(0).AddItem "RRT-Rect.Consumo Manta"
''        Combo1(0).ItemData(Combo1(0).NewIndex) = 9
'
'    End If
    
    '[Monica]21/01/2016: cargamos los tipos de movimiento de aquellos que hayan facturas pendientes de integrar
    
    Sql = "select codtipom from rrecibpozos where contabilizado = 0 group by 1 order by 1 "
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenKeyset, adLockOptimistic, adCmdText
    
    I = -1
    While Not Rs.EOF
        I = I + 1
        Select Case DBLet(Rs.Fields(0).Value, "N")
            Case "RCP"
                Combo1(0).AddItem "RCP-Consumo"
            Case "RMP"
                Combo1(0).AddItem "RMP-Mantenimiento"
            Case "RVP"
                Combo1(0).AddItem "RVP-Contadores"
            Case "TAL"
                Combo1(0).AddItem "TAL-Talla"
            Case "RMT"
                Combo1(0).AddItem "RMT-Consumo Manta"
            Case "RRC"
                Combo1(0).AddItem "RRC-Rect.Consumo"
            Case "RRM"
                Combo1(0).AddItem "RRM-Rect.Mantenimiento"
            Case "RRV"
                Combo1(0).AddItem "RRV-Rect.Contadores"
            Case "RTA"
                Combo1(0).AddItem "RTA-Rect.Talla"
            Case "RRT"
                Combo1(0).AddItem "RRT-Rect.Consumo Manta"
            Case "FIN"
                Combo1(0).AddItem "FIN-Interna"
        End Select
        Combo1(0).ItemData(Combo1(0).NewIndex) = I
    
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
End Sub

