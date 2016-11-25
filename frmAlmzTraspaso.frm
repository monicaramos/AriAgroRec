VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAlmzTraspaso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6630
   Icon            =   "frmAlmzTraspaso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameTraspaso 
      Height          =   4665
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6555
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Text5"
         Top             =   2145
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   2
         Top             =   2130
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   1
         Top             =   1650
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   1665
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   3
         Top             =   2610
         Width           =   1095
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text5"
         Top             =   1170
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   0
         Top             =   1170
         Width           =   735
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4905
         TabIndex        =   5
         Top             =   4110
         Width           =   975
      End
      Begin VB.CommandButton CmdAcep 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   4110
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar pb2 
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   3090
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   570
         Top             =   3720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "doc"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "(Aceituna)"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   19
         Top             =   2280
         Width           =   720
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1620
         MouseIcon       =   "frmAlmzTraspaso.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar tipo iva"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Iva Prov."
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   2
         Left            =   420
         TabIndex        =   18
         Top             =   2040
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "(Aceite Stock)"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   16
         Top             =   1770
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Iva Cliente"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   0
         Left            =   420
         TabIndex        =   15
         Top             =   1560
         Width           =   1110
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1620
         MouseIcon       =   "frmAlmzTraspaso.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar tipo iva"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   19
         Left            =   420
         TabIndex        =   13
         Top             =   2640
         Width           =   1035
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1620
         Picture         =   "frmAlmzTraspaso.frx":02B0
         ToolTipText     =   "Buscar fecha"
         Top             =   2610
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Traspaso de Almazara"
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
         TabIndex        =   12
         Top             =   450
         Width           =   5025
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1620
         MouseIcon       =   "frmAlmzTraspaso.frx":033B
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1185
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   3
         Left            =   420
         TabIndex        =   11
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label lblProgres 
         Caption         =   "aa"
         Height          =   375
         Index           =   3
         Left            =   180
         TabIndex        =   9
         Top             =   3450
         Width           =   6195
      End
      Begin VB.Label lblProgres 
         Caption         =   "aa"
         Height          =   285
         Index           =   2
         Left            =   180
         TabIndex        =   8
         Top             =   3810
         Width           =   6195
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2640
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmAlmzTraspaso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MONICA +-+-
' +-+-+-+-+-+-+-+-+-+-+-+
' TRASPASO DE ALMAZARA

Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto


Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmVar As frmComVar 'Variedades de comercial
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'Mensajes
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmTIva As frmTipIVAConta ' tipos de iva de contabilidad
Attribute frmTIva.VB_VarHelpID = -1



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
Dim Tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String

Dim indice As Integer

Dim PrimeraVez As Boolean
Dim Contabilizada As Byte
Dim ConSubInforme As Boolean

Dim Fichero1 As String
Dim Fichero2 As String
Dim Fichero3 As String

' *****VARIABLES FACTURAS
' PARA LA INSERCION DE REGISTROS EN LAS TABLAS DE FACTURAS
Dim Socio As String
Dim Fecha As String
Dim Factura As String
Dim Base As String
Dim iva As String
Dim TotFactu As String
Dim ImpReten As String

Dim Calidad As String
Dim CodTipom As String
Dim numlinea As Integer
Dim cantidad As String
Dim Precio As String
Dim Importe As String
Dim Concepto As String
Dim campo As String

Dim CantidadVar As Currency
Dim ImporteVar As Currency

' *****END VARIABLES FACTURAS

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me

End Sub


Private Sub CmdAcep_Click()
Dim Sql As String
Dim i As Byte
Dim cadWHERE As String
Dim b As Boolean
Dim NomFic As String
Dim cadena As String
Dim cadena1 As String
Dim nompath As String
Dim cadTabla As String

On Error GoTo eError

    If Not DatosOk Then Exit Sub
    

    Fichero1 = ""
    Fichero2 = ""
    Fichero3 = ""
 
'    nompath = GetFolder("Selecciona directorio")
    
    Me.CommonDialog2.DefaultExt = "TXT"
    Me.CommonDialog2.FileName = "ACEITEC.TXT"
    CommonDialog2.FilterIndex = 1
    Me.CommonDialog2.ShowOpen
        
    BorrarTMPlineas
    b = CrearTMPlineas()
    If Not b Then
         Exit Sub
    End If
    
    conn.BeginTrans
    
    If Me.CommonDialog2.FileName <> "" Then
        nompath = CurDir(Me.CommonDialog2.FileName)
        If ExistenFicheros(nompath) Then
            Fichero1 = nompath & "\aceitec.txt"
            Fichero2 = nompath & "\aceitunc.txt"
            Fichero3 = nompath & "\stockc.txt"
            
            InicializarVbles
                
                '========= PARAMETROS  =============================
            'Añadir el parametro de Empresa
            CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
            numParam = numParam + 1
    
            If ComprobarErrores() Then
                    cadTabla = "tmpinformes"
                    cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                    
                    Sql = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo
                    
                    If TotalRegistros(Sql) <> 0 Then
                        MsgBox "Hay errores en los ficheros de Traspaso. Debe corregirlos previamente.", vbExclamation
                        cadTitulo = "Errores de Traspaso"
                        cadNombreRPT = "rErroresTraspaso.rpt"
                        LlamarImprimir
                        conn.RollbackTrans
                        Exit Sub
                    Else
                        b = CargarFicheros()
                    End If
            Else
                b = False
            End If
        Else
            cmdCancel_Click
            Exit Sub
        End If
        
    End If
eError:
    If Err.Number <> 0 Or Not b Then
        conn.RollbackTrans
        MsgBox "No se ha podido realizar el proceso. LLame a Ariadna.", vbExclamation
    Else
        conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        pb2.visible = False
        lblProgres(2).Caption = ""
        lblProgres(3).Caption = ""
'        BorrarArchivo Fichero1
'        BorrarArchivo Fichero2
'        BorrarArchivo Fichero3
        cmdCancel_Click
    End If

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me
    
    ConSubInforme = False

    For H = 0 To 1
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    
    Me.imgBuscar(3).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    
    'Ocultar todos los Frames de Formulario
    FrameTraspaso.visible = False
    
    '###Descomentar
'    CommitConexion
    
    FrameTraspasoVisible True, H, W
    pb2.visible = False
    lblProgres(2).Caption = ""
    lblProgres(3).Caption = ""
        
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
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub


Private Sub frmTIva_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgBuscar_Click(Index As Integer)
Dim vSeccion As CSeccion

   Select Case Index
        Case 0 'VARIEDADES
            AbrirFrmVariedad (Index)
            
        Case 1, 3 ' tipo de iva de contabilidad
            indCodigo = Index
            PonerFoco txtcodigo(indCodigo)
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(vParamAplic.SeccionAlmaz) Then
                If vSeccion.AbrirConta Then
                    Set frmTIva = New frmTipIVAConta
                    frmTIva.DeConsulta = True
                    frmTIva.DatosADevolverBusqueda = "0|1|"
                    frmTIva.CodigoActual = txtcodigo(indCodigo).Text
                    frmTIva.Show vbModal
                    Set frmTIva = Nothing
                    PonerFoco txtcodigo(indCodigo)
                End If
            End If
            Set vSeccion = Nothing
        
    
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
            indice = 2
    End Select


    imgFec(0).Tag = indice '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtcodigo(indice).Text <> "" Then frmC.NovaData = txtcodigo(indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtcodigo(indice) '<===
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
            Case 0: KEYBusqueda KeyAscii, 0 'variedad
            Case 1: KEYBusqueda KeyAscii, 1 'tipo de iva
            Case 3: KEYBusqueda KeyAscii, 3 'tipo de iva proveedor
            Case 2: KEYFecha KeyAscii, 0 'fecha factura
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
Dim vSeccion As CSeccion

    'Quitar espacios en blanco por los lados
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 2 'FECHAS
            '[Monica]28/08/2013: comprobamos que la fecha esté en la campaña
            If txtcodigo(Index).Text <> "" Then PonerFormatoFecha txtcodigo(Index), True
            
        Case 1, 3 ' tipo de iva
            If txtcodigo(Index).Text <> "" Then
                Set vSeccion = New CSeccion
                If vSeccion.LeerDatos(vParamAplic.SeccionAlmaz) Then
                    If vSeccion.AbrirConta Then
                        txtNombre(Index).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", txtcodigo(Index).Text, "N")
                    End If
                End If
                Set vSeccion = Nothing
            End If
            
        Case 0 'VARIEDADES
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
        
    End Select
End Sub


Private Sub FrameTraspasoVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el trapaso de almazara
    Me.FrameTraspaso.visible = visible
    If visible = True Then
        Me.FrameTraspaso.Top = -90
        Me.FrameTraspaso.Left = 0
        Me.FrameTraspaso.Height = 4665
        Me.FrameTraspaso.Width = 6555
        W = Me.FrameTraspaso.Width
        H = Me.FrameTraspaso.Height
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
        .EnvioEMail = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .ConSubInforme = ConSubInforme
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmVariedad(indice As Integer)
    indCodigo = indice
    Set frmVar = New frmComVar
    frmVar.DatosADevolverBusqueda = "0|1|"
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
        .Opcion = 0
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



Private Function ActualizarRegistros(cTabla As String, cWhere As String) As Boolean
'Actualizar la marca de impreso
Dim Sql As String

    On Error GoTo eActualizarRegistros

    ActualizarRegistros = False
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "update " & QuitarCaracterACadena(cTabla, "_1") & " set impreso = 1 "
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    conn.Execute Sql
    
    ActualizarRegistros = True
    Exit Function
    
eActualizarRegistros:
    MuestraError Err.Number, "Actualizando registros", Err.Description
End Function



Private Function ComprobarErrores() As Boolean
Dim NF As Long
Dim Cad As String
Dim i As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim NumReg As Long
Dim Sql As String
Dim Sql1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean
Dim Mens As String
Dim Tipo As Integer


    On Error GoTo eComprobarErrores

    ComprobarErrores = False
    
    ' comprobamos que tenga asignada la seccion de almazara
    If vParamAplic.SeccionAlmaz = "" Then
        MsgBox "No tiene asignada la seccion de almazara en parámetros. Revise.", vbExclamation
        Exit Function
    End If
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
'ACEITEC.TXT
    lblProgres(2).Caption = "Comprobando errores fichero ACEITEC.TXT "
    
    NF = FreeFile
    Open Fichero1 For Input As #NF
    
    Line Input #NF, Cad
    i = 0
    
    longitud = FileLen(Fichero1)
    
    pb2.visible = True
    Me.pb2.Max = longitud
    Me.Refresh
    Me.pb2.Value = 0
    
    ' PROCESO DEL FICHERO ACEITEC.TXT
    b = True
    While Not EOF(NF) And b
        i = i + 1
        
        Me.pb2.Value = Me.pb2.Value + Len(Cad)
        lblProgres(3).Caption = "Linea " & i
        Me.Refresh
        
        b = ComprobarRegistro(Cad, "ACEITEC.TXT")
        
        Line Input #NF, Cad
    Wend
    Close #NF
    
    If Cad <> "" And b Then
        i = i + 1
        
        Me.pb2.Value = Me.pb2.Value + Len(Cad)
        lblProgres(3).Caption = "Linea " & i
        Me.Refresh
        
        b = ComprobarRegistro(Cad, "ACEITEC.TXT")
    End If
    
    pb2.visible = False
    lblProgres(2).Caption = ""
    lblProgres(3).Caption = ""

    
'ACEITUNC.TXT
    lblProgres(2).Caption = "Comprobando errores fichero ACEITUNC.TXT "
    
    NF = FreeFile
    Open Fichero2 For Input As #NF
    
    Line Input #NF, Cad
    i = 0
    
    longitud = FileLen(Fichero2)
    
    pb2.visible = True
    Me.pb2.Max = longitud
    Me.Refresh
    Me.pb2.Value = 0
    
    ' PROCESO DEL FICHERO ACEITUNC.TXT
    b = True
    While Not EOF(NF) And b
        i = i + 1
        
        Me.pb2.Value = Me.pb2.Value + Len(Cad)
        lblProgres(3).Caption = "Linea " & i
        Me.Refresh
        
        b = ComprobarRegistro(Cad, "ACEITUNC.TXT")
        
        Line Input #NF, Cad
    Wend
    Close #NF
    
    If Cad <> "" And b Then
        i = i + 1
        
        Me.pb2.Value = Me.pb2.Value + Len(Cad)
        lblProgres(3).Caption = "Linea " & i
        Me.Refresh
        
        b = ComprobarRegistro(Cad, "ACEITUNC.TXT")
    End If
    
    pb2.visible = False
    lblProgres(2).Caption = ""
    lblProgres(3).Caption = ""


'STOCKC.TXT

    lblProgres(2).Caption = "Comprobando errores fichero STOCKC.TXT "
    
    NF = FreeFile
    Open Fichero3 For Input As #NF
    
    Line Input #NF, Cad
    i = 0
    
    longitud = FileLen(Fichero3)
    
    pb2.visible = True
    Me.pb2.Max = longitud
    Me.Refresh
    Me.pb2.Value = 0
    
    ' PROCESO DEL FICHERO STOCKC.TXT
    b = True
    While Not EOF(NF) And b
        i = i + 1
        
        Me.pb2.Value = Me.pb2.Value + Len(Cad)
        lblProgres(3).Caption = "Linea " & i
        Me.Refresh
        
        b = ComprobarRegistro(Cad, "STOCKC.TXT")
        
        Line Input #NF, Cad
    Wend
    Close #NF
    
    If Cad <> "" And b Then
        i = i + 1
        
        Me.pb2.Value = Me.pb2.Value + Len(Cad)
        lblProgres(3).Caption = "Linea " & i
        Me.Refresh
        
        b = ComprobarRegistro(Cad, "STOCKC.TXT")
    End If
    
    pb2.visible = False
    lblProgres(2).Caption = ""
    lblProgres(3).Caption = ""

    ComprobarErrores = b
    Exit Function

eComprobarErrores:
    ComprobarErrores = False
End Function


Private Function CargarFicheros() As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim longitud As Long

Dim cadMen As String
Dim Cad As String

Dim NF As Integer
Dim i As Integer
Dim b As Boolean


    On Error GoTo eCargarFicheros
    
    CargarFicheros = False
    
    ' PROCESO DEL FICHERO ACEITEC.TXT
    lblProgres(2).Caption = "Cargando Fichero ACEITEC.TXT "
    
    NF = FreeFile
    Open Fichero1 For Input As #NF
    
    Line Input #NF, Cad
    i = 0
    
    longitud = FileLen(Fichero1)
    
    pb2.visible = True
    Me.pb2.Max = longitud
    Me.Refresh
    Me.pb2.Value = 0
    
    b = True
    While Not EOF(NF) And b
        i = i + 1
        
        Me.pb2.Value = Me.pb2.Value + Len(Cad)
        lblProgres(3).Caption = "Linea " & i
        Me.Refresh
        
        b = InsertarRegistros(Cad, 0)
            
        Line Input #NF, Cad
    Wend
    Close #NF
    
    If Cad <> "" And b Then
        i = i + 1
        
        Me.pb2.Value = Me.pb2.Value + Len(Cad)
        lblProgres(3).Caption = "Linea " & i
        Me.Refresh
        
        b = InsertarRegistros(Cad, 0)
    End If
    
    pb2.visible = False
    lblProgres(2).Caption = ""
    lblProgres(3).Caption = ""

    
    If b Then
        ' PROCESO DEL FICHERO ACEITUNC.TXT
        lblProgres(2).Caption = "Cargando Fichero ACEITUNC.TXT "
        
        NF = FreeFile
        Open Fichero2 For Input As #NF
        
        Line Input #NF, Cad
        i = 0
        
        longitud = FileLen(Fichero2)
        
        pb2.visible = True
        Me.pb2.Max = longitud
        Me.Refresh
        Me.pb2.Value = 0
        
        While Not EOF(NF) And b
            i = i + 1
            
            Me.pb2.Value = Me.pb2.Value + Len(Cad)
            lblProgres(3).Caption = "Linea " & i
            Me.Refresh
            
            b = InsertarRegistros(Cad, 1)
                
            Line Input #NF, Cad
        Wend
        Close #NF
        
        If Cad <> "" And b Then
            i = i + 1
            
            Me.pb2.Value = Me.pb2.Value + Len(Cad)
            lblProgres(3).Caption = "Linea " & i
            Me.Refresh
            
            b = InsertarRegistros(Cad, 1)
        End If
    End If
        
        
    pb2.visible = False
    lblProgres(2).Caption = ""
    lblProgres(3).Caption = ""


    If b Then
        ' PROCESO DEL FICHERO STOCKC.TXT
        lblProgres(2).Caption = "Cargando Fichero STOCKC.TXT "
        
        NF = FreeFile
        Open Fichero3 For Input As #NF
        
        Line Input #NF, Cad
        i = 0
        
        longitud = FileLen(Fichero3)
        
        pb2.visible = True
        Me.pb2.Max = longitud
        Me.Refresh
        Me.pb2.Value = 0
        
        While Not EOF(NF) And b
            i = i + 1
            
            Me.pb2.Value = Me.pb2.Value + Len(Cad)
            lblProgres(3).Caption = "Linea " & i
            Me.Refresh
            
            b = InsertarRegistros(Cad, 2)
                
            Line Input #NF, Cad
        Wend
        Close #NF
        
        If Cad <> "" And b Then
            i = i + 1
            
            Me.pb2.Value = Me.pb2.Value + Len(Cad)
            lblProgres(3).Caption = "Linea " & i
            Me.Refresh
            
            b = InsertarRegistros(Cad, 2)
        End If
        
        pb2.visible = False
        lblProgres(2).Caption = ""
        lblProgres(3).Caption = ""
    End If
    
    If b Then
        b = InsertarTemporales
    End If
    
    If b Then
        CargarFicheros = True
    End If
    
    Exit Function
    
    
eCargarFicheros:
    MuestraError Err.Number, "Cargar ficheros", Err.Description
End Function



Private Function ExistenFicheros(nompath As String) As Boolean
Dim b1 As Boolean
Dim b2 As Boolean
Dim b3 As Boolean
Dim cadMen As String

    On Error GoTo eExistenFicheros


    ExistenFicheros = False
    b1 = False
    b2 = False
    b3 = False
    
    cadMen = "Los Ficheros : " & vbCrLf
    
    If Dir(nompath & "\aceitec.txt") <> "" Then
        cadMen = cadMen & vbCrLf & "        aceitec.txt"
        b1 = True
    End If
    If Dir(nompath & "\aceitunc.txt") <> "" Then
        cadMen = cadMen & vbCrLf & "        aceitunc.txt"
        b2 = True
    End If
    If Dir(nompath & "\stockc.txt") <> "" Then
        cadMen = cadMen & vbCrLf & "        stockc.txt"
        b3 = True
    End If
    
    
    If Not (b1 And b2 And b3) Then
        cadMen = cadMen & vbCrLf & vbCrLf & "no existen en el directorio seleccionado. Revise." & vbCrLf
        MsgBox cadMen, vbExclamation
    End If
    ExistenFicheros = (b1 And b2 And b3)
    Exit Function
    
eExistenFicheros:
    MuestraError Err.Number, "Error en Existen ficheros"
End Function


Private Function ComprobarRegistro(Cad As String, Fichero As String) As Boolean
Dim Socio As String
Dim Sql As String
Dim Sql1 As String
Dim Mens As String

    On Error GoTo eComprobarRegistro

    ComprobarRegistro = True

    If Mid(Cad, 1, 3) = "CAB" Then
        Socio = Mid(Cad, 6, 5)
        Fecha = Mid(Cad, 50, 10)
        Factura = Mid(Cad, 60, 6)
        Select Case Fichero
            Case "ACEITEC.TXT"
                Tipo = 0
            Case "ACEITUNC.TXT"
                Tipo = 1
            Case "STOCKC.TXT"
                Tipo = 2
        End Select

        'Comprobamos que el socio existe
        If Socio <> "" Then
            Sql = ""
            Sql = DevuelveDesdeBDNew(cAgro, "rsocios", "codsocio", "codsocio", Socio, "N")
            If Sql = "" Then
                
                Mens = "No existe el socio " & Format(Socio, "000000") & "-" & Fichero
                Sql = "insert into tmpinformes (codusu, nombre1) values (" & _
                      vUsu.Codigo & "," & DBSet(Mens, "T") & ")"
                
                conn.Execute Sql
            End If
        
            If Fichero = "ACEITUNC.TXT" Then
                ' comprobamos que el socio tiene un campo para la variedad introducida sin fecha de baja
                Sql1 = "select min(codcampo) from rcampos where codsocio = " & DBSet(Socio, "N")
                Sql1 = Sql1 & " and codvarie = " & DBSet(txtcodigo(0).Text, "N")
                Sql1 = Sql1 & " and fecbajas is null"
                
                If DevuelveValor(Sql1) = 0 Then
                    Mens = "No existe campo del socio " & Format(Socio, "000000") & "-" & Fichero
                    Sql = "insert into tmpinformes (codusu, nombre1) values (" & _
                          vUsu.Codigo & "," & DBSet(Mens, "T") & ")"
                    
                    conn.Execute Sql
                End If
                
                ' comprobamos que el socio es de la seccion de almazara
                Sql1 = "select count(*) from rsocios_seccion where codsocio = " & DBSet(Socio, "N")
                Sql1 = Sql1 & " and codsecci = " & vParamAplic.SeccionAlmaz
                If TotalRegistros(Sql1) = 0 Then
                    Mens = "El socio " & Format(Socio, "000000") & " no es de almazara"
                    Sql = "insert into tmpinformes (codusu, nombre1) values (" & _
                          vUsu.Codigo & "," & DBSet(Mens, "T") & ")"
                    
                    conn.Execute Sql
                
                
                End If
                
            End If
        End If
        
        
        ' COMPROBAMOS QUE LA FACTURA NO EXISTA
        Sql1 = "select count(*) from rcabfactalmz where codsocio = " & DBSet(Socio, "N")
        Sql1 = Sql1 & " and numfactu = " & DBSet(Factura, "N")
        Sql1 = Sql1 & " and fecfactu = " & DBSet(Fecha, "F")
        Sql1 = Sql1 & " and tipofichero = " & DBSet(Tipo, "N")
        
        If TotalRegistros(Sql1) <> 0 Then
            Mens = "Existe la factura " & Format(Factura, "0000000") & "-" & Format(Fecha, "dd/mm/yyyy")
            Sql = "insert into tmpinformes (codusu, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(Mens, "T") & ")"
            
            conn.Execute Sql
        End If
        
        
        
    End If
eComprobarRegistro:
    If Err.Number <> 0 Then
        ComprobarRegistro = False
    End If
End Function


Private Function InsertarRegistros(Cad As String, Tipo As Byte) As Boolean
' Tipo = 0 --> aceitec
'        1 --> aceitunc
'        2 --> stockc
Dim Sql As String
Dim Sql1 As String


Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim TipoIVA As String
Dim TipoIRPF As String
Dim PorcIva As String
Dim PorcReten As Currency
Dim BaseReten As Currency
        
    On Error GoTo eInsertarRegistros
    
    CodTipom = "FAZ"  ' tipo de movimiento de almazara
    
    InsertarRegistros = False
        
    If Mid(Cad, 1, 3) = "CAB" Then
        Socio = Mid(Cad, 6, 5)
        Fecha = Mid(Cad, 50, 10)
        Factura = Mid(Cad, 60, 6)
        
        numlinea = 0
        
        CantidadVar = 0
        ImporteVar = 0
    End If
    
 
    If Mid(Cad, 1, 3) = "DET" Then
        Concepto = ValorNulo
        cantidad = ValorNulo
        Precio = ValorNulo
        Importe = ValorNulo
            
        numlinea = numlinea + 1
        Concepto = Mid(Cad, 12, 30)
        cantidad = Mid(Cad, 43, 9)
        Precio = Mid(Cad, 50, 8)
        Importe = Mid(Cad, 58, 11)
    
        CantidadVar = CantidadVar + cantidad
        ImporteVar = ImporteVar + Importe
    
        
        Sql = "insert into tmprlinfactalmz (tipofichero, numfactu, fecfactu, codsocio, numlinea, "
        Sql = Sql & " concepto, cantidad, precioar, importel) values ( "
        Sql = Sql & DBSet(Tipo, "N") & ","
        Sql = Sql & DBSet(Factura, "N") & ","
        Sql = Sql & DBSet(Fecha, "F") & ","
        Sql = Sql & DBSet(Socio, "N") & ","
        Sql = Sql & DBSet(numlinea, "N") & ","
        Sql = Sql & DBSet(Concepto, "T") & ","
        Sql = Sql & DBSet(cantidad, "N") & ","
        Sql = Sql & DBSet(Precio, "N") & ","
        Sql = Sql & DBSet(Importe, "N") & ")"
        
        conn.Execute Sql
    
    End If
    
    
    If Mid(Cad, 1, 3) = "TOT" Then
        Base = Mid(Cad, 4, 11)
        iva = Mid(Cad, 20, 11)
        TotFactu = Mid(Cad, 31, 11)
        
        ImpReten = ""
        PorcReten = 0
        BaseReten = 0
        If Tipo = 1 Then ' fichero ACEITUNC.TXT
            ImpReten = Mid(Cad, 36, 11)
            TotFactu = Mid(Cad, 47, 11)
            PorcReten = Mid(Cad, 31, 5)
        End If
        
        TipoIVA = ""
        TipoIRPF = ""
        PorcIva = ""
        
        Set vSocio = New cSocio
        If vSocio.LeerDatosSeccion(CStr(Socio), vParamAplic.SeccionAlmaz) Then
            '[Monica]28/10/2015: ya no se coge el iva de la seccion de almazara del socio
            TipoIVA = txtcodigo(3).Text 'vSocio.CodIva
            TipoIRPF = vSocio.TipoIRPF
            If TipoIVA <> "" Then
                Set vSeccion = New CSeccion
                If vSeccion.LeerDatos(vParamAplic.SeccionAlmaz) Then
                    If vSeccion.AbrirConta Then
                        PorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", TipoIVA, "N")
                    End If
                End If
                Set vSeccion = Nothing
            End If
            Select Case TipoIRPF
                Case 0
                    BaseReten = CCur(Base) + CCur(iva)
                Case 1
                    BaseReten = CCur(Base)
                Case 2
                    BaseReten = 0
            End Select
            If PorcReten = 0 Then BaseReten = 0
        End If
        Set vSocio = Nothing
             
        If Tipo = 1 Then ' fichero ACEITUNC.TXT
             
             'CABECERA DE FACTURA
             Sql = "insert into rfactsoc (`codtipom`,`numfactu`,`fecfactu`,`codsocio`,"
             Sql = Sql & "`baseimpo`,`tipoiva`,`porc_iva`,`imporiva`,`tipoirpf`,"
             Sql = Sql & "`basereten`,`porc_ret`,`impreten`,`baseaport`,`porc_apo`,"
             Sql = Sql & "`impapor`,`totalfac`,`impreso`,`contabilizado`,`pasaridoc`,"
             Sql = Sql & "`esanticipogasto`) values ("
             Sql = Sql & DBSet(CodTipom, "T") & ","
             Sql = Sql & DBSet(Factura, "N") & ","
             Sql = Sql & DBSet(txtcodigo(2).Text, "F") & ","
             Sql = Sql & DBSet(Socio, "N") & ","
             Sql = Sql & DBSet(Base, "N") & ","
             Sql = Sql & DBSet(TipoIVA, "N") & "," ' tipo de iva
             Sql = Sql & DBSet(PorcIva, "N") & "," ' porcentaje iva
             Sql = Sql & DBSet(iva, "N") & "," ' importe iva
             Sql = Sql & DBSet(TipoIRPF, "N") & "," ' tipo irfpf
             Sql = Sql & DBSet(BaseReten, "N", "S") & "," ' base de retencion
             Sql = Sql & DBSet(PorcReten, "N", "S") & "," ' porcentaje de retencion
             Sql = Sql & DBSet(ImpReten, "N", "S") & ","
             Sql = Sql & ValorNulo & "," ' base de aportacion
             Sql = Sql & ValorNulo & "," ' porcentaje de aportacion
             Sql = Sql & ValorNulo & "," ' importe de aportacion
             Sql = Sql & DBSet(TotFactu, "N") & "," ' total factura
             Sql = Sql & "0,1,0,0) " ' se introduce como contabilizada
             
             conn.Execute Sql
        
        
            'VARIEDAD DE FACTURA
            Sql1 = "select min(codcampo) from rcampos where codsocio = " & DBSet(Socio, "N") & " and "
            Sql1 = Sql1 & " codvarie = " & DBSet(txtcodigo(0).Text, "N") & " and fecbajas is null "
            
            campo = DevuelveValor(Sql1)
            
            Sql = "insert into temprfactsoc_variedad (codtipom,numfactu,fecfactu,codVarie,codCampo, "
            Sql = Sql & "kilosnet, preciomed, imporvar, descontado) values ("
            Sql = Sql & DBSet(CodTipom, "T") & ","
            Sql = Sql & DBSet(Factura, "N") & ","
            Sql = Sql & DBSet(txtcodigo(2).Text, "F") & ","
            Sql = Sql & DBSet(txtcodigo(0).Text, "N") & ","
            Sql = Sql & DBSet(campo, "N") & "," 'campo
            Sql = Sql & DBSet(CantidadVar, "N") & "," 'kilos
            Sql = Sql & DBSet(Round2(ImporteVar / CantidadVar, 4), "N") & "," ' precio
            Sql = Sql & DBSet(ImporteVar, "N") & ",0)" ' importe y no descontado
            
            conn.Execute Sql
        
            Calidad = CalidadPrimera(txtcodigo(0).Text)
            
            'CALIDAD DE FACTURA
            Sql = "insert into temprfactsoc_calidad (codtipom,numfactu,fecfactu,codVarie,codCampo, "
            Sql = Sql & "codcalid, kilosnet, precio, imporcal) values ("
            Sql = Sql & DBSet(CodTipom, "T") & ","
            Sql = Sql & DBSet(Factura, "N") & ","
            Sql = Sql & DBSet(txtcodigo(2).Text, "F") & ","
            Sql = Sql & DBSet(txtcodigo(0).Text, "N") & ","
            Sql = Sql & DBSet(campo, "N") & "," 'campo
            Sql = Sql & DBSet(Calidad, "N") & "," ' calidad: ponemos la calidad primera
            Sql = Sql & DBSet(CantidadVar, "N") & "," 'kilos
            Sql = Sql & DBSet(Round2(ImporteVar / CantidadVar, 4), "N") & "," ' precio
            Sql = Sql & DBSet(ImporteVar, "N") & ")" ' importe y no descontado
            
            conn.Execute Sql
        
            CantidadVar = 0
            ImporteVar = 0
        
        Else
            TipoIVA = txtcodigo(1).Text
            
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(vParamAplic.SeccionAlmaz) Then
                If vSeccion.AbrirConta Then
                    PorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", TipoIVA, "N")
                End If
            End If
            Set vSeccion = Nothing
            
        End If
        
        ' insertamos cabecera de factura de almazara
        Sql = "insert into rcabfactalmz (tipofichero, numfactu, fecfactu, codsocio, baseimpo, tipoiva, "
        Sql = Sql & "porc_iva, imporiva, tipoirpf, basereten, porc_ret, impreten, totalfac, impreso, "
        Sql = Sql & "contabilizado) values ("
        Sql = Sql & DBSet(Tipo, "N") & ","
        Sql = Sql & DBSet(Factura, "N") & ","
        Sql = Sql & DBSet(Fecha, "F") & ","
        Sql = Sql & DBSet(Socio, "N") & ","
        Sql = Sql & DBSet(Base, "N") & ","
        Sql = Sql & DBSet(TipoIVA, "N") & "," ' tipo iva
        Sql = Sql & DBSet(PorcIva, "N") & "," 'porcentaje iva
        Sql = Sql & DBSet(iva, "N") & "," ' importe iva
        Sql = Sql & DBSet(TipoIRPF, "N") & "," ' tipo irpf
        Sql = Sql & DBSet(BaseReten, "N", "S") & "," ' base de retencion
        Sql = Sql & DBSet(PorcReten, "N", "S") & "," ' porcentaje de retencion
        Sql = Sql & DBSet(ImpReten, "N", "S") & "," 'importe de retencion
        Sql = Sql & DBSet(TotFactu, "N") & "," 'total factura
        Sql = Sql & "0,0)" ' impreso y contabilizado
        
        conn.Execute Sql
        
    End If
    
    InsertarRegistros = True
    Exit Function
    
eInsertarRegistros:
    MuestraError Err.Number, "Insertar Registros", Err.Description
End Function



Private Function CrearTMPlineas() As Boolean
' temporales de lineas para insertar posteriormente en rfactsoc_variedad y rfactsoc_calidad
Dim Sql As String
    
    On Error GoTo ECrear
    
    CrearTMPlineas = False
    
    'rfactsoc_variedad
    Sql = "CREATE TEMPORARY TABLE temprfactsoc_variedad ( "
    Sql = Sql & "`codtipom` char(3) NOT NULL ,"
    Sql = Sql & "`numfactu` int(7) unsigned NOT NULL,"
    Sql = Sql & "`fecfactu` date NOT NULL,"
    Sql = Sql & "`codvarie` int(6) NOT NULL,"
    Sql = Sql & "`codcampo` int(8) unsigned NOT NULL,"
    Sql = Sql & "`kilosnet` int(6) NOT NULL,"
    Sql = Sql & "`preciomed` decimal(6,4) NOT NULL,"
    Sql = Sql & "`imporvar` decimal(8,2) NOT NULL,"
    Sql = Sql & "`descontado` tinyint(1) NOT NULL default '0')"
    
    conn.Execute Sql
    
    'rfactsoc_calidad
    Sql = "CREATE TEMPORARY  TABLE temprfactsoc_calidad ( "
    Sql = Sql & "`codtipom` char(3),"
    Sql = Sql & "`numfactu` int(7) unsigned NOT NULL,"
    Sql = Sql & "`fecfactu` date NOT NULL,"
    Sql = Sql & "`codvarie` int(6) NOT NULL,"
    Sql = Sql & "`codcampo` int(8) unsigned NOT NULL,"
    Sql = Sql & "`codcalid` smallint(2) NOT NULL,"
    Sql = Sql & "`kilosnet` int(6) NOT NULL,"
    Sql = Sql & "`precio` decimal(6,4) NOT NULL,"
    Sql = Sql & "`imporcal` decimal(8,2) NOT NULL)"
    
    conn.Execute Sql
     
    ' si es liquidacion venta campo o no se insertaran en los anticipos
    Sql = "CREATE TEMPORARY  TABLE tmprlinfactalmz ( "
    Sql = Sql & "`tipofichero` tinyint(1) unsigned NOT NULL, "
    Sql = Sql & "`numfactu` int(7) unsigned NOT NULL, "
    Sql = Sql & "`fecfactu` date NOT NULL, "
    Sql = Sql & "`codsocio` int(6) unsigned NOT NULL,"
    Sql = Sql & "`numlinea` smallint(4) unsigned NOT NULL,"
    Sql = Sql & "`concepto` varchar(40) NOT NULL,"
    Sql = Sql & "`cantidad` int(7) NOT NULL,"
    Sql = Sql & "`precioar` decimal(8,4) NOT NULL,"
    Sql = Sql & "`importel` decimal(8,2) NOT NULL) "
    
    conn.Execute Sql
     
    CrearTMPlineas = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPlineas = False
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmprlinfactalmz;"
        conn.Execute Sql
        Sql = " DROP TABLE IF EXISTS temprfactsoc_variedad;"
        conn.Execute Sql
        Sql = " DROP TABLE IF EXISTS temprfactsoc_calidad;"
        conn.Execute Sql
    End If
End Function

Private Sub BorrarTMPlineas()
On Error Resume Next

    conn.Execute " DROP TABLE IF EXISTS tmprlinfactalmz;"
    conn.Execute " DROP TABLE IF EXISTS temprfactsoc_variedad;"
    conn.Execute " DROP TABLE IF EXISTS temprfactsoc_calidad;"
    
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Function InsertarTemporales() As Boolean
Dim Sql As String
Dim Sql1 As String


Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim TipoIVA As String
Dim TipoIRPF As String
Dim PorcIva As String

        
    On Error GoTo eInsertarTemporales
    
    
    InsertarTemporales = False
        
    Sql = "insert into rfactsoc_variedad (codtipom,numfactu,fecfactu,codvarie,codcampo,kilosnet,preciomed,imporvar,descontado) "
    Sql = Sql & " select codtipom,numfactu,fecfactu,codvarie,codcampo,kilosnet,preciomed,imporvar,descontado from temprfactsoc_variedad "
    
    conn.Execute Sql
    
    Sql = "insert into rfactsoc_calidad (codtipom,numfactu,fecfactu,codvarie,codcampo,codcalid,kilosnet,precio,imporcal) "
    Sql = Sql & " select codtipom,numfactu,fecfactu,codvarie,codcampo,codcalid,kilosnet,precio,imporcal from temprfactsoc_calidad "
    
    conn.Execute Sql
    
    Sql = "insert into rlinfactalmz (tipofichero,numfactu,fecfactu,codsocio,numlinea,concepto,cantidad,precioar,importel) "
    Sql = Sql & " select tipofichero,numfactu,fecfactu,codsocio,numlinea,concepto,cantidad,precioar,importel from tmprlinfactalmz "
    conn.Execute Sql
    
    InsertarTemporales = True
    Exit Function
    
eInsertarTemporales:
    MuestraError Err.Number, "Insertar Temporales", Err.Description
End Function

Private Function DatosOk() As Boolean
Dim b As Boolean

    DatosOk = False

    If txtcodigo(0).Text = "" Then
        MsgBox "Debe introducir la variedad.", vbExclamation
        PonerFoco txtcodigo(2)
        Exit Function
    End If
    
    If txtcodigo(1).Text = "" Or txtcodigo(3).Text = "" Then
        MsgBox "Debe introducir los tipos de iva.", vbExclamation
        PonerFoco txtcodigo(1)
        Exit Function
    End If
    
    If txtcodigo(2).Text = "" Then
        MsgBox "Introduzca la fecha de factura.", vbExclamation
        PonerFoco txtcodigo(2)
        Exit Function
    End If
    
    DatosOk = True

End Function
