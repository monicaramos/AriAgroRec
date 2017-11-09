VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAlmzTraspaso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6720
   Icon            =   "frmAlmzTraspaso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameTraspaso 
      Height          =   4665
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6690
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
         Left            =   3060
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Text5"
         Top             =   2145
         Width           =   3375
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
         Left            =   2145
         MaxLength       =   3
         TabIndex        =   2
         Top             =   2130
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
         Index           =   1
         Left            =   2145
         MaxLength       =   3
         TabIndex        =   1
         Top             =   1650
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
         Index           =   1
         Left            =   3060
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   1665
         Width           =   3375
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
         Left            =   2145
         MaxLength       =   10
         TabIndex        =   3
         Top             =   2610
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
         Index           =   0
         Left            =   3060
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text5"
         Top             =   1170
         Width           =   3375
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
         Left            =   2145
         MaxLength       =   6
         TabIndex        =   0
         Top             =   1170
         Width           =   870
      End
      Begin VB.CommandButton CmdCancel 
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
         Left            =   5355
         TabIndex        =   5
         Top             =   4110
         Width           =   1065
      End
      Begin VB.CommandButton CmdAcep 
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
         Left            =   4170
         TabIndex        =   4
         Top             =   4110
         Width           =   1065
      End
      Begin MSComctlLib.ProgressBar pb2 
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   3090
         Width           =   6225
         _ExtentX        =   10980
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
         Index           =   4
         Left            =   345
         TabIndex        =   19
         Top             =   2280
         Width           =   1050
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1890
         MouseIcon       =   "frmAlmzTraspaso.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar tipo iva"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Iva Prov."
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
         Left            =   285
         TabIndex        =   18
         Top             =   2040
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "(Aceite Stock)"
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
         Index           =   1
         Left            =   345
         TabIndex        =   16
         Top             =   1770
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Iva Cliente"
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
         Index           =   0
         Left            =   285
         TabIndex        =   15
         Top             =   1560
         Width           =   1560
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1890
         MouseIcon       =   "frmAlmzTraspaso.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar tipo iva"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   19
         Left            =   285
         TabIndex        =   13
         Top             =   2640
         Width           =   1440
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1890
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
         Left            =   330
         TabIndex        =   12
         Top             =   450
         Width           =   5025
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1890
         MouseIcon       =   "frmAlmzTraspaso.frx":033B
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1185
         Width           =   240
      End
      Begin VB.Label Label2 
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
         Index           =   3
         Left            =   330
         TabIndex        =   11
         Top             =   1155
         Width           =   1080
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
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String

Dim Indice As Integer

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
Dim SQL As String
Dim I As Byte
Dim cadWHERE As String
Dim B As Boolean
Dim NomFic As String
Dim cadena As String
Dim cadena1 As String
Dim nompath As String
Dim cadTabla As String

On Error GoTo eError

    If Not DatosOK Then Exit Sub
    

    Fichero1 = ""
    Fichero2 = ""
    Fichero3 = ""
 
'    nompath = GetFolder("Selecciona directorio")
    
    Me.CommonDialog2.DefaultExt = "TXT"
    Me.CommonDialog2.FileName = "ACEITEC.TXT"
    CommonDialog2.FilterIndex = 1
    Me.CommonDialog2.ShowOpen
        
    BorrarTMPlineas
    B = CrearTMPlineas()
    If Not B Then
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
                    
                    SQL = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo
                    
                    If TotalRegistros(SQL) <> 0 Then
                        MsgBox "Hay errores en los ficheros de Traspaso. Debe corregirlos previamente.", vbExclamation
                        cadTitulo = "Errores de Traspaso"
                        cadNombreRPT = "rErroresTraspaso.rpt"
                        LlamarImprimir
                        conn.RollbackTrans
                        Exit Sub
                    Else
                        B = CargarFicheros()
                    End If
            Else
                B = False
            End If
        Else
            cmdCancel_Click
            Exit Sub
        End If
        
    End If
eError:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        MsgBox "No se ha podido realizar el proceso. LLame a Ariadna.", vbExclamation
    Else
        conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        Pb2.visible = False
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
    'Icono del formulario
    Me.Icon = frmPpal.Icon

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
    Pb2.visible = False
    lblProgres(2).Caption = ""
    lblProgres(3).Caption = ""
        
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
'    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtCodigo(Indice).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
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
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub


Private Sub frmTIva_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgBuscar_Click(Index As Integer)
Dim vSeccion As CSeccion

   Select Case Index
        Case 0 'VARIEDADES
            AbrirFrmVariedad (Index)
            
        Case 1, 3 ' tipo de iva de contabilidad
            indCodigo = Index
            PonerFoco txtCodigo(indCodigo)
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(vParamAplic.SeccionAlmaz) Then
                If vSeccion.AbrirConta Then
                    Set frmTIva = New frmTipIVAConta
                    frmTIva.DeConsulta = True
                    frmTIva.DatosADevolverBusqueda = "0|1|"
                    frmTIva.CodigoActual = txtCodigo(indCodigo).Text
                    frmTIva.Show vbModal
                    Set frmTIva = Nothing
                    PonerFoco txtCodigo(indCodigo)
                End If
            End If
            Set vSeccion = Nothing
        
    
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
            Indice = 2
    End Select


    imgFec(0).Tag = Indice '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtCodigo(Indice).Text <> "" Then frmC.NovaData = txtCodigo(Indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtCodigo(Indice) '<===
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
            Case 0: KEYBusqueda KeyAscii, 0 'variedad
            Case 1: KEYBusqueda KeyAscii, 1 'tipo de iva
            Case 3: KEYBusqueda KeyAscii, 3 'tipo de iva proveedor
            Case 2: KEYFecha KeyAscii, 0 'fecha factura
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
Dim vSeccion As CSeccion

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 2 'FECHAS
            '[Monica]28/08/2013: comprobamos que la fecha esté en la campaña
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index), True
            
        Case 1, 3 ' tipo de iva
            If txtCodigo(Index).Text <> "" Then
                Set vSeccion = New CSeccion
                If vSeccion.LeerDatos(vParamAplic.SeccionAlmaz) Then
                    If vSeccion.AbrirConta Then
                        txtNombre(Index).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", txtCodigo(Index).Text, "N")
                    End If
                End If
                Set vSeccion = Nothing
            End If
            
        Case 0 'VARIEDADES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
    End Select
End Sub


Private Sub FrameTraspasoVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el trapaso de almazara
    Me.FrameTraspaso.visible = visible
    If visible = True Then
        Me.FrameTraspaso.Top = -90
        Me.FrameTraspaso.Left = 0
        Me.FrameTraspaso.Height = 4665
        Me.FrameTraspaso.Width = 6690
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

Private Sub AbrirFrmVariedad(Indice As Integer)
    indCodigo = Indice
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
Dim SQL As String

    On Error GoTo eActualizarRegistros

    ActualizarRegistros = False
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    SQL = "update " & QuitarCaracterACadena(cTabla, "_1") & " set impreso = 1 "
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    conn.Execute SQL
    
    ActualizarRegistros = True
    Exit Function
    
eActualizarRegistros:
    MuestraError Err.Number, "Actualizando registros", Err.Description
End Function



Private Function ComprobarErrores() As Boolean
Dim NF As Long
Dim cad As String
Dim I As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim NumReg As Long
Dim SQL As String
Dim Sql1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim B As Boolean
Dim Mens As String
Dim Tipo As Integer


    On Error GoTo eComprobarErrores

    ComprobarErrores = False
    
    ' comprobamos que tenga asignada la seccion de almazara
    If vParamAplic.SeccionAlmaz = "" Then
        MsgBox "No tiene asignada la seccion de almazara en parámetros. Revise.", vbExclamation
        Exit Function
    End If
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
'ACEITEC.TXT
    lblProgres(2).Caption = "Comprobando errores fichero ACEITEC.TXT "
    
    NF = FreeFile
    Open Fichero1 For Input As #NF
    
    Line Input #NF, cad
    I = 0
    
    longitud = FileLen(Fichero1)
    
    Pb2.visible = True
    Me.Pb2.Max = longitud
    Me.Refresh
    Me.Pb2.Value = 0
    
    ' PROCESO DEL FICHERO ACEITEC.TXT
    B = True
    While Not EOF(NF) And B
        I = I + 1
        
        Me.Pb2.Value = Me.Pb2.Value + Len(cad)
        lblProgres(3).Caption = "Linea " & I
        Me.Refresh
        
        B = ComprobarRegistro(cad, "ACEITEC.TXT")
        
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" And B Then
        I = I + 1
        
        Me.Pb2.Value = Me.Pb2.Value + Len(cad)
        lblProgres(3).Caption = "Linea " & I
        Me.Refresh
        
        B = ComprobarRegistro(cad, "ACEITEC.TXT")
    End If
    
    Pb2.visible = False
    lblProgres(2).Caption = ""
    lblProgres(3).Caption = ""

    
'ACEITUNC.TXT
    lblProgres(2).Caption = "Comprobando errores fichero ACEITUNC.TXT "
    
    NF = FreeFile
    Open Fichero2 For Input As #NF
    
    Line Input #NF, cad
    I = 0
    
    longitud = FileLen(Fichero2)
    
    Pb2.visible = True
    Me.Pb2.Max = longitud
    Me.Refresh
    Me.Pb2.Value = 0
    
    ' PROCESO DEL FICHERO ACEITUNC.TXT
    B = True
    While Not EOF(NF) And B
        I = I + 1
        
        Me.Pb2.Value = Me.Pb2.Value + Len(cad)
        lblProgres(3).Caption = "Linea " & I
        Me.Refresh
        
        B = ComprobarRegistro(cad, "ACEITUNC.TXT")
        
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" And B Then
        I = I + 1
        
        Me.Pb2.Value = Me.Pb2.Value + Len(cad)
        lblProgres(3).Caption = "Linea " & I
        Me.Refresh
        
        B = ComprobarRegistro(cad, "ACEITUNC.TXT")
    End If
    
    Pb2.visible = False
    lblProgres(2).Caption = ""
    lblProgres(3).Caption = ""


'STOCKC.TXT

    lblProgres(2).Caption = "Comprobando errores fichero STOCKC.TXT "
    
    NF = FreeFile
    Open Fichero3 For Input As #NF
    
    Line Input #NF, cad
    I = 0
    
    longitud = FileLen(Fichero3)
    
    Pb2.visible = True
    Me.Pb2.Max = longitud
    Me.Refresh
    Me.Pb2.Value = 0
    
    ' PROCESO DEL FICHERO STOCKC.TXT
    B = True
    While Not EOF(NF) And B
        I = I + 1
        
        Me.Pb2.Value = Me.Pb2.Value + Len(cad)
        lblProgres(3).Caption = "Linea " & I
        Me.Refresh
        
        B = ComprobarRegistro(cad, "STOCKC.TXT")
        
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" And B Then
        I = I + 1
        
        Me.Pb2.Value = Me.Pb2.Value + Len(cad)
        lblProgres(3).Caption = "Linea " & I
        Me.Refresh
        
        B = ComprobarRegistro(cad, "STOCKC.TXT")
    End If
    
    Pb2.visible = False
    lblProgres(2).Caption = ""
    lblProgres(3).Caption = ""

    ComprobarErrores = B
    Exit Function

eComprobarErrores:
    ComprobarErrores = False
End Function


Private Function CargarFicheros() As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim longitud As Long

Dim cadMen As String
Dim cad As String

Dim NF As Integer
Dim I As Integer
Dim B As Boolean


    On Error GoTo eCargarFicheros
    
    CargarFicheros = False
    
    ' PROCESO DEL FICHERO ACEITEC.TXT
    lblProgres(2).Caption = "Cargando Fichero ACEITEC.TXT "
    
    NF = FreeFile
    Open Fichero1 For Input As #NF
    
    Line Input #NF, cad
    I = 0
    
    longitud = FileLen(Fichero1)
    
    Pb2.visible = True
    Me.Pb2.Max = longitud
    Me.Refresh
    Me.Pb2.Value = 0
    
    B = True
    While Not EOF(NF) And B
        I = I + 1
        
        Me.Pb2.Value = Me.Pb2.Value + Len(cad)
        lblProgres(3).Caption = "Linea " & I
        Me.Refresh
        
        B = InsertarRegistros(cad, 0)
            
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" And B Then
        I = I + 1
        
        Me.Pb2.Value = Me.Pb2.Value + Len(cad)
        lblProgres(3).Caption = "Linea " & I
        Me.Refresh
        
        B = InsertarRegistros(cad, 0)
    End If
    
    Pb2.visible = False
    lblProgres(2).Caption = ""
    lblProgres(3).Caption = ""

    
    If B Then
        ' PROCESO DEL FICHERO ACEITUNC.TXT
        lblProgres(2).Caption = "Cargando Fichero ACEITUNC.TXT "
        
        NF = FreeFile
        Open Fichero2 For Input As #NF
        
        Line Input #NF, cad
        I = 0
        
        longitud = FileLen(Fichero2)
        
        Pb2.visible = True
        Me.Pb2.Max = longitud
        Me.Refresh
        Me.Pb2.Value = 0
        
        While Not EOF(NF) And B
            I = I + 1
            
            Me.Pb2.Value = Me.Pb2.Value + Len(cad)
            lblProgres(3).Caption = "Linea " & I
            Me.Refresh
            
            B = InsertarRegistros(cad, 1)
                
            Line Input #NF, cad
        Wend
        Close #NF
        
        If cad <> "" And B Then
            I = I + 1
            
            Me.Pb2.Value = Me.Pb2.Value + Len(cad)
            lblProgres(3).Caption = "Linea " & I
            Me.Refresh
            
            B = InsertarRegistros(cad, 1)
        End If
    End If
        
        
    Pb2.visible = False
    lblProgres(2).Caption = ""
    lblProgres(3).Caption = ""


    If B Then
        ' PROCESO DEL FICHERO STOCKC.TXT
        lblProgres(2).Caption = "Cargando Fichero STOCKC.TXT "
        
        NF = FreeFile
        Open Fichero3 For Input As #NF
        
        Line Input #NF, cad
        I = 0
        
        longitud = FileLen(Fichero3)
        
        Pb2.visible = True
        Me.Pb2.Max = longitud
        Me.Refresh
        Me.Pb2.Value = 0
        
        While Not EOF(NF) And B
            I = I + 1
            
            Me.Pb2.Value = Me.Pb2.Value + Len(cad)
            lblProgres(3).Caption = "Linea " & I
            Me.Refresh
            
            B = InsertarRegistros(cad, 2)
                
            Line Input #NF, cad
        Wend
        Close #NF
        
        If cad <> "" And B Then
            I = I + 1
            
            Me.Pb2.Value = Me.Pb2.Value + Len(cad)
            lblProgres(3).Caption = "Linea " & I
            Me.Refresh
            
            B = InsertarRegistros(cad, 2)
        End If
        
        Pb2.visible = False
        lblProgres(2).Caption = ""
        lblProgres(3).Caption = ""
    End If
    
    If B Then
        B = InsertarTemporales
    End If
    
    If B Then
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


Private Function ComprobarRegistro(cad As String, Fichero As String) As Boolean
Dim Socio As String
Dim SQL As String
Dim Sql1 As String
Dim Mens As String

    On Error GoTo eComprobarRegistro

    ComprobarRegistro = True

    If Mid(cad, 1, 3) = "CAB" Then
        Socio = Mid(cad, 6, 5)
        Fecha = Mid(cad, 50, 10)
        Factura = Mid(cad, 60, 6)
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
            SQL = ""
            SQL = DevuelveDesdeBDNew(cAgro, "rsocios", "codsocio", "codsocio", Socio, "N")
            If SQL = "" Then
                
                Mens = "No existe el socio " & Format(Socio, "000000") & "-" & Fichero
                SQL = "insert into tmpinformes (codusu, nombre1) values (" & _
                      vUsu.Codigo & "," & DBSet(Mens, "T") & ")"
                
                conn.Execute SQL
            End If
        
            If Fichero = "ACEITUNC.TXT" Then
                ' comprobamos que el socio tiene un campo para la variedad introducida sin fecha de baja
                Sql1 = "select min(codcampo) from rcampos where codsocio = " & DBSet(Socio, "N")
                Sql1 = Sql1 & " and codvarie = " & DBSet(txtCodigo(0).Text, "N")
                Sql1 = Sql1 & " and fecbajas is null"
                
                If DevuelveValor(Sql1) = 0 Then
                    Mens = "No existe campo del socio " & Format(Socio, "000000") & "-" & Fichero
                    SQL = "insert into tmpinformes (codusu, nombre1) values (" & _
                          vUsu.Codigo & "," & DBSet(Mens, "T") & ")"
                    
                    conn.Execute SQL
                End If
                
                ' comprobamos que el socio es de la seccion de almazara
                Sql1 = "select count(*) from rsocios_seccion where codsocio = " & DBSet(Socio, "N")
                Sql1 = Sql1 & " and codsecci = " & vParamAplic.SeccionAlmaz
                If TotalRegistros(Sql1) = 0 Then
                    Mens = "El socio " & Format(Socio, "000000") & " no es de almazara"
                    SQL = "insert into tmpinformes (codusu, nombre1) values (" & _
                          vUsu.Codigo & "," & DBSet(Mens, "T") & ")"
                    
                    conn.Execute SQL
                
                
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
            SQL = "insert into tmpinformes (codusu, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(Mens, "T") & ")"
            
            conn.Execute SQL
        End If
        
        
        
    End If
eComprobarRegistro:
    If Err.Number <> 0 Then
        ComprobarRegistro = False
    End If
End Function


Private Function InsertarRegistros(cad As String, Tipo As Byte) As Boolean
' Tipo = 0 --> aceitec
'        1 --> aceitunc
'        2 --> stockc
Dim SQL As String
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
        
    If Mid(cad, 1, 3) = "CAB" Then
        Socio = Mid(cad, 6, 5)
        Fecha = Mid(cad, 50, 10)
        Factura = Mid(cad, 60, 6)
        
        numlinea = 0
        
        CantidadVar = 0
        ImporteVar = 0
    End If
    
 
    If Mid(cad, 1, 3) = "DET" Then
        Concepto = ValorNulo
        cantidad = ValorNulo
        Precio = ValorNulo
        Importe = ValorNulo
            
        numlinea = numlinea + 1
        Concepto = Mid(cad, 12, 30)
        cantidad = Mid(cad, 43, 9)
        Precio = Mid(cad, 50, 8)
        Importe = Mid(cad, 58, 11)
    
        CantidadVar = CantidadVar + cantidad
        ImporteVar = ImporteVar + Importe
    
        
        SQL = "insert into tmprlinfactalmz (tipofichero, numfactu, fecfactu, codsocio, numlinea, "
        SQL = SQL & " concepto, cantidad, precioar, importel) values ( "
        SQL = SQL & DBSet(Tipo, "N") & ","
        SQL = SQL & DBSet(Factura, "N") & ","
        SQL = SQL & DBSet(Fecha, "F") & ","
        SQL = SQL & DBSet(Socio, "N") & ","
        SQL = SQL & DBSet(numlinea, "N") & ","
        SQL = SQL & DBSet(Concepto, "T") & ","
        SQL = SQL & DBSet(cantidad, "N") & ","
        SQL = SQL & DBSet(Precio, "N") & ","
        SQL = SQL & DBSet(Importe, "N") & ")"
        
        conn.Execute SQL
    
    End If
    
    
    If Mid(cad, 1, 3) = "TOT" Then
        Base = Mid(cad, 4, 11)
        iva = Mid(cad, 20, 11)
        TotFactu = Mid(cad, 31, 11)
        
        ImpReten = ""
        PorcReten = 0
        BaseReten = 0
        If Tipo = 1 Then ' fichero ACEITUNC.TXT
            ImpReten = Mid(cad, 36, 11)
            TotFactu = Mid(cad, 47, 11)
            PorcReten = Mid(cad, 31, 5)
        End If
        
        TipoIVA = ""
        TipoIRPF = ""
        PorcIva = ""
        
        Set vSocio = New cSocio
        If vSocio.LeerDatosSeccion(CStr(Socio), vParamAplic.SeccionAlmaz) Then
            '[Monica]28/10/2015: ya no se coge el iva de la seccion de almazara del socio
            TipoIVA = txtCodigo(3).Text 'vSocio.CodIva
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
             SQL = "insert into rfactsoc (`codtipom`,`numfactu`,`fecfactu`,`codsocio`,"
             SQL = SQL & "`baseimpo`,`tipoiva`,`porc_iva`,`imporiva`,`tipoirpf`,"
             SQL = SQL & "`basereten`,`porc_ret`,`impreten`,`baseaport`,`porc_apo`,"
             SQL = SQL & "`impapor`,`totalfac`,`impreso`,`contabilizado`,`pasaridoc`,"
             SQL = SQL & "`esanticipogasto`) values ("
             SQL = SQL & DBSet(CodTipom, "T") & ","
             SQL = SQL & DBSet(Factura, "N") & ","
             SQL = SQL & DBSet(txtCodigo(2).Text, "F") & ","
             SQL = SQL & DBSet(Socio, "N") & ","
             SQL = SQL & DBSet(Base, "N") & ","
             SQL = SQL & DBSet(TipoIVA, "N") & "," ' tipo de iva
             SQL = SQL & DBSet(PorcIva, "N") & "," ' porcentaje iva
             SQL = SQL & DBSet(iva, "N") & "," ' importe iva
             SQL = SQL & DBSet(TipoIRPF, "N") & "," ' tipo irfpf
             SQL = SQL & DBSet(BaseReten, "N", "S") & "," ' base de retencion
             SQL = SQL & DBSet(PorcReten, "N", "S") & "," ' porcentaje de retencion
             SQL = SQL & DBSet(ImpReten, "N", "S") & ","
             SQL = SQL & ValorNulo & "," ' base de aportacion
             SQL = SQL & ValorNulo & "," ' porcentaje de aportacion
             SQL = SQL & ValorNulo & "," ' importe de aportacion
             SQL = SQL & DBSet(TotFactu, "N") & "," ' total factura
             SQL = SQL & "0,1,0,0) " ' se introduce como contabilizada
             
             conn.Execute SQL
        
        
            'VARIEDAD DE FACTURA
            Sql1 = "select min(codcampo) from rcampos where codsocio = " & DBSet(Socio, "N") & " and "
            Sql1 = Sql1 & " codvarie = " & DBSet(txtCodigo(0).Text, "N") & " and fecbajas is null "
            
            campo = DevuelveValor(Sql1)
            
            SQL = "insert into temprfactsoc_variedad (codtipom,numfactu,fecfactu,codVarie,codCampo, "
            SQL = SQL & "kilosnet, preciomed, imporvar, descontado) values ("
            SQL = SQL & DBSet(CodTipom, "T") & ","
            SQL = SQL & DBSet(Factura, "N") & ","
            SQL = SQL & DBSet(txtCodigo(2).Text, "F") & ","
            SQL = SQL & DBSet(txtCodigo(0).Text, "N") & ","
            SQL = SQL & DBSet(campo, "N") & "," 'campo
            SQL = SQL & DBSet(CantidadVar, "N") & "," 'kilos
            SQL = SQL & DBSet(Round2(ImporteVar / CantidadVar, 4), "N") & "," ' precio
            SQL = SQL & DBSet(ImporteVar, "N") & ",0)" ' importe y no descontado
            
            conn.Execute SQL
        
            Calidad = CalidadPrimera(txtCodigo(0).Text)
            
            'CALIDAD DE FACTURA
            SQL = "insert into temprfactsoc_calidad (codtipom,numfactu,fecfactu,codVarie,codCampo, "
            SQL = SQL & "codcalid, kilosnet, precio, imporcal) values ("
            SQL = SQL & DBSet(CodTipom, "T") & ","
            SQL = SQL & DBSet(Factura, "N") & ","
            SQL = SQL & DBSet(txtCodigo(2).Text, "F") & ","
            SQL = SQL & DBSet(txtCodigo(0).Text, "N") & ","
            SQL = SQL & DBSet(campo, "N") & "," 'campo
            SQL = SQL & DBSet(Calidad, "N") & "," ' calidad: ponemos la calidad primera
            SQL = SQL & DBSet(CantidadVar, "N") & "," 'kilos
            SQL = SQL & DBSet(Round2(ImporteVar / CantidadVar, 4), "N") & "," ' precio
            SQL = SQL & DBSet(ImporteVar, "N") & ")" ' importe y no descontado
            
            conn.Execute SQL
        
            CantidadVar = 0
            ImporteVar = 0
        
        Else
            TipoIVA = txtCodigo(1).Text
            
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(vParamAplic.SeccionAlmaz) Then
                If vSeccion.AbrirConta Then
                    PorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", TipoIVA, "N")
                End If
            End If
            Set vSeccion = Nothing
            
        End If
        
        ' insertamos cabecera de factura de almazara
        SQL = "insert into rcabfactalmz (tipofichero, numfactu, fecfactu, codsocio, baseimpo, tipoiva, "
        SQL = SQL & "porc_iva, imporiva, tipoirpf, basereten, porc_ret, impreten, totalfac, impreso, "
        SQL = SQL & "contabilizado) values ("
        SQL = SQL & DBSet(Tipo, "N") & ","
        SQL = SQL & DBSet(Factura, "N") & ","
        SQL = SQL & DBSet(Fecha, "F") & ","
        SQL = SQL & DBSet(Socio, "N") & ","
        SQL = SQL & DBSet(Base, "N") & ","
        SQL = SQL & DBSet(TipoIVA, "N") & "," ' tipo iva
        SQL = SQL & DBSet(PorcIva, "N") & "," 'porcentaje iva
        SQL = SQL & DBSet(iva, "N") & "," ' importe iva
        SQL = SQL & DBSet(TipoIRPF, "N") & "," ' tipo irpf
        SQL = SQL & DBSet(BaseReten, "N", "S") & "," ' base de retencion
        SQL = SQL & DBSet(PorcReten, "N", "S") & "," ' porcentaje de retencion
        SQL = SQL & DBSet(ImpReten, "N", "S") & "," 'importe de retencion
        SQL = SQL & DBSet(TotFactu, "N") & "," 'total factura
        SQL = SQL & "0,0)" ' impreso y contabilizado
        
        conn.Execute SQL
        
    End If
    
    InsertarRegistros = True
    Exit Function
    
eInsertarRegistros:
    MuestraError Err.Number, "Insertar Registros", Err.Description
End Function



Private Function CrearTMPlineas() As Boolean
' temporales de lineas para insertar posteriormente en rfactsoc_variedad y rfactsoc_calidad
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPlineas = False
    
    'rfactsoc_variedad
    SQL = "CREATE TEMPORARY TABLE temprfactsoc_variedad ( "
    SQL = SQL & "`codtipom` char(3) NOT NULL ,"
    SQL = SQL & "`numfactu` int(7) unsigned NOT NULL,"
    SQL = SQL & "`fecfactu` date NOT NULL,"
    SQL = SQL & "`codvarie` int(6) NOT NULL,"
    SQL = SQL & "`codcampo` int(8) unsigned NOT NULL,"
    SQL = SQL & "`kilosnet` int(6) NOT NULL,"
    SQL = SQL & "`preciomed` decimal(6,4) NOT NULL,"
    SQL = SQL & "`imporvar` decimal(8,2) NOT NULL,"
    SQL = SQL & "`descontado` tinyint(1) NOT NULL default '0')"
    
    conn.Execute SQL
    
    'rfactsoc_calidad
    SQL = "CREATE TEMPORARY  TABLE temprfactsoc_calidad ( "
    SQL = SQL & "`codtipom` char(3),"
    SQL = SQL & "`numfactu` int(7) unsigned NOT NULL,"
    SQL = SQL & "`fecfactu` date NOT NULL,"
    SQL = SQL & "`codvarie` int(6) NOT NULL,"
    SQL = SQL & "`codcampo` int(8) unsigned NOT NULL,"
    SQL = SQL & "`codcalid` smallint(2) NOT NULL,"
    SQL = SQL & "`kilosnet` int(6) NOT NULL,"
    SQL = SQL & "`precio` decimal(6,4) NOT NULL,"
    SQL = SQL & "`imporcal` decimal(8,2) NOT NULL)"
    
    conn.Execute SQL
     
    ' si es liquidacion venta campo o no se insertaran en los anticipos
    SQL = "CREATE TEMPORARY  TABLE tmprlinfactalmz ( "
    SQL = SQL & "`tipofichero` tinyint(1) unsigned NOT NULL, "
    SQL = SQL & "`numfactu` int(7) unsigned NOT NULL, "
    SQL = SQL & "`fecfactu` date NOT NULL, "
    SQL = SQL & "`codsocio` int(6) unsigned NOT NULL,"
    SQL = SQL & "`numlinea` smallint(4) unsigned NOT NULL,"
    SQL = SQL & "`concepto` varchar(40) NOT NULL,"
    SQL = SQL & "`cantidad` int(7) NOT NULL,"
    SQL = SQL & "`precioar` decimal(8,4) NOT NULL,"
    SQL = SQL & "`importel` decimal(8,2) NOT NULL) "
    
    conn.Execute SQL
     
    CrearTMPlineas = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPlineas = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmprlinfactalmz;"
        conn.Execute SQL
        SQL = " DROP TABLE IF EXISTS temprfactsoc_variedad;"
        conn.Execute SQL
        SQL = " DROP TABLE IF EXISTS temprfactsoc_calidad;"
        conn.Execute SQL
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
Dim SQL As String
Dim Sql1 As String


Dim vSocio As cSocio
Dim vSeccion As CSeccion
Dim TipoIVA As String
Dim TipoIRPF As String
Dim PorcIva As String

        
    On Error GoTo eInsertarTemporales
    
    
    InsertarTemporales = False
        
    SQL = "insert into rfactsoc_variedad (codtipom,numfactu,fecfactu,codvarie,codcampo,kilosnet,preciomed,imporvar,descontado) "
    SQL = SQL & " select codtipom,numfactu,fecfactu,codvarie,codcampo,kilosnet,preciomed,imporvar,descontado from temprfactsoc_variedad "
    
    conn.Execute SQL
    
    SQL = "insert into rfactsoc_calidad (codtipom,numfactu,fecfactu,codvarie,codcampo,codcalid,kilosnet,precio,imporcal) "
    SQL = SQL & " select codtipom,numfactu,fecfactu,codvarie,codcampo,codcalid,kilosnet,precio,imporcal from temprfactsoc_calidad "
    
    conn.Execute SQL
    
    SQL = "insert into rlinfactalmz (tipofichero,numfactu,fecfactu,codsocio,numlinea,concepto,cantidad,precioar,importel) "
    SQL = SQL & " select tipofichero,numfactu,fecfactu,codsocio,numlinea,concepto,cantidad,precioar,importel from tmprlinfactalmz "
    conn.Execute SQL
    
    InsertarTemporales = True
    Exit Function
    
eInsertarTemporales:
    MuestraError Err.Number, "Insertar Temporales", Err.Description
End Function

Private Function DatosOK() As Boolean
Dim B As Boolean
Dim vSeccion As CSeccion

    DatosOK = False

    If txtCodigo(0).Text = "" Then
        MsgBox "Debe introducir la variedad.", vbExclamation
        PonerFoco txtCodigo(2)
        Exit Function
    End If
    
    If txtCodigo(1).Text = "" Or txtCodigo(3).Text = "" Then
        MsgBox "Debe introducir los tipos de iva.", vbExclamation
        PonerFoco txtCodigo(1)
        Exit Function
    End If
    
    If txtCodigo(2).Text = "" Then
        MsgBox "Introduzca la fecha de factura.", vbExclamation
        PonerFoco txtCodigo(2)
        Exit Function
    End If
    
    '[Monica]20/06/2017: control de fechas que antes no estaba
    Set vSeccion = New CSeccion
    If vSeccion.LeerDatos(vParamAplic.SeccionAlmaz) Then
        If vSeccion.AbrirConta Then
            ResultadoFechaContaOK = EsFechaOKConta(CDate(txtCodigo(2)))
            If ResultadoFechaContaOK > 0 Then
                If ResultadoFechaContaOK <> 4 Then MsgBox MensajeFechaOkConta, vbExclamation
                Exit Function
            End If
        End If
    End If
    Set vSeccion = Nothing
    
    
    
    
    
    DatosOK = True

End Function
