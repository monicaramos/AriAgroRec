VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIntTesorPago 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contabilización Baja Campos"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7485
   Icon            =   "frmIntTesorPago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCobros 
      Height          =   3840
      Left            =   75
      TabIndex        =   4
      Top             =   120
      Width           =   7275
      Begin VB.Frame Frame2 
         Caption         =   "Datos para la contabilización"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1440
         Left            =   135
         TabIndex        =   5
         Top             =   300
         Width           =   7020
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
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   1
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Text            =   "0000000000"
            Top             =   930
            Width           =   1305
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
            Left            =   3330
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   930
            Width           =   3450
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
            Left            =   1980
            MaxLength       =   10
            TabIndex        =   0
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   495
            Width           =   1305
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
            Left            =   3330
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   495
            Width           =   3450
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   1710
            ToolTipText     =   "Buscar Concepto"
            Top             =   975
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
            Height          =   285
            Index           =   24
            Left            =   180
            TabIndex        =   12
            Top             =   975
            Width           =   1665
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   1710
            ToolTipText     =   "Buscar Concepto"
            Top             =   540
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
            TabIndex        =   10
            Top             =   540
            Width           =   1665
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
         Left            =   5955
         TabIndex        =   3
         Top             =   3075
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
         Left            =   4770
         TabIndex        =   2
         Top             =   3075
         Width           =   1065
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   345
         Left            =   300
         TabIndex        =   6
         Top             =   2010
         Width           =   6810
         _ExtentX        =   12012
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   0
         Left            =   495
         TabIndex        =   8
         Top             =   2415
         Width           =   5265
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   495
         TabIndex        =   7
         Top             =   2730
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
Attribute VB_Name = "frmIntTesorPago"
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

Public Campos As String

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
Private cadNombreRPT As String 'Nombre del informe

Dim indCodigo As Integer 'indice para txtCodigo
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String

Dim PrimeraVez As Boolean

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim SQL As String
Dim I As Byte
Dim cadWHERE As String

    If Not DatosOK Then Exit Sub
             
    SQL = "SELECT count(*)" & _
          " FROM (" & CadTag & ") aaaa "
          
             
    If RegistrosAListar(SQL) = 0 Then
        MsgBox "No existen datos a contabilizar.", vbExclamation
        Exit Sub
    End If
    
    ContabilizarPagos (CadTag)
    DesBloqueoManual ("CONTES") 'CONtabilizacion a TESoreria
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    cmdCancel_Click
    
eError:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "No se ha realizado el proceso de contabilización a tesoreria. Llame a soporte."
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(0)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
     Me.imgBuscar(3).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Me.imgBuscar(4).Picture = frmPpal.imgListImages16.ListImages(1).Picture

    '###Descomentar
'    CommitConexion
         
         
    FrameCobrosVisible True, H, W
    Pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
'   Me.Width = w + 70
'   Me.Height = h + 350
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
    frmFPa.CodigoActual = txtCodigo(indCodigo)
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

Private Sub ContabilizarPagos(cadWHERE As String)
'Contabiliza Facturas de Clientes o de Proveedores
Dim SQL As String
Dim B As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String
Dim cadTabla As String
Dim FFin As Date
Dim cadMen As String

    SQL = "CONTES" 'contabilizar tesoreria

    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (SQL)
    If Not BloqueoManual(SQL, "1") Then
        MsgBox "No se pueden Contabilizar Pagos. Hay otro usuario contabilizándolo.", vbExclamation
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
        
    
    'comprobar que todas las CUENTAS de socio
    'en la Conta: savnic.codmacta IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgres(1).Caption = "Comprobando Cuenta Ctble de Socio ..."
    
    Dim cta As String
    cta = "1501" & Format(RecuperaValor(NumCod, 2), String(vEmpresa.DigitosUltimoNivel - vEmpresa.DigitosNivelAnterior, "0"))
    
    'Obtener los dias de pago de la tabla de parametros: spara1
    B = ComprobarCtaContable_new("", 15, , vParamAplic.Seccionhorto, cta)
    IncrementarProgres Me.Pb1, 20
    Me.Refresh
    DoEvents
    If Not B Then Exit Sub
    
    
    '===========================================================================
    'CONTABILIZAR PAGO
    '===========================================================================
    Me.lblProgres(0).Caption = "Contabilizar Pago: "
    Me.lblProgres(1).Caption = "Insertando Registro en Tesorería..."
    cadMen = "Insertando en Tesoreria"
    B = InsertarEnTesoreriaBajaCampo(CadTag, cadMen, RecuperaValor(NumCod, 3), RecuperaValor(NumCod, 1), txtCodigo(3).Text, txtCodigo(4).Text, RecuperaValor(NumCod, 2), Campos)
    IncrementarProgres Me.Pb1, 80
    Me.Refresh
    DoEvents
    
    If Not B Then
        MuestraError Err.Number, "Contabilizar pagos.", cadMen
    Else
        ContabilizadoOk = True
    End If
    
End Sub

Private Function DatosOK() As Boolean
Dim B As Boolean
Dim Orden1 As String
Dim Orden2 As String
Dim FFin As Date

   B = True


    
   If txtCodigo(3).Text = "" And B Then
        MsgBox "Introduzca la Forma de Pago para contabilizar.", vbExclamation
        B = False
        PonerFoco txtCodigo(3)
   End If
   
   If txtCodigo(4).Text = "" And B Then
        MsgBox "Introduzca la Cta.Contable de Banco para contabilizar.", vbExclamation
        B = False
        PonerFoco txtCodigo(4)
   End If
   
   
   
   DatosOK = B
   
End Function


Private Function ActualizarCobros(cadWHERE As String, cadErr As String) As Boolean
'Poner el movimiento como contabilizada
Dim SQL As String

    On Error GoTo EActualizar
    
    SQL = "UPDATE segpoliza SET inttesor=1 "
    SQL = SQL & " WHERE " & cadWHERE

    conn.Execute SQL
    
EActualizar:
    If Err.Number <> 0 Then
        ActualizarCobros = False
        cadErr = Err.Description
    Else
        ActualizarCobros = True
    End If
End Function



