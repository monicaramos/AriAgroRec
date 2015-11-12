VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExpImpExcel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6330
   Icon            =   "frmExpImpExcel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameHorasTrabajadas 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   6375
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   4155
         Left            =   150
         TabIndex        =   4
         Top             =   1860
         Width           =   5820
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   7
            Left            =   2550
            Locked          =   -1  'True
            TabIndex        =   30
            Text            =   "Text5"
            Top             =   3660
            Width           =   3015
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   6
            Left            =   2550
            Locked          =   -1  'True
            TabIndex        =   29
            Text            =   "Text5"
            Top             =   3285
            Width           =   3015
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   6
            Left            =   1710
            MaxLength       =   6
            TabIndex        =   11
            Top             =   3270
            Width           =   750
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   7
            Left            =   1725
            MaxLength       =   6
            TabIndex        =   12
            Top             =   3645
            Width           =   750
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   1740
            MaxLength       =   7
            TabIndex        =   5
            Tag             =   "Nº Parte|N|S|||rpartes|nroparte|0000000|S|"
            Top             =   420
            Width           =   750
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   1740
            MaxLength       =   7
            TabIndex        =   6
            Tag             =   "Nº Parte|N|S|||rpartes|nroparte|0000000|S|"
            Top             =   795
            Width           =   750
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   1710
            MaxLength       =   10
            TabIndex        =   7
            Top             =   1350
            Width           =   1005
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   1710
            MaxLength       =   10
            TabIndex        =   8
            Top             =   1740
            Width           =   1005
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   5
            Left            =   2550
            Locked          =   -1  'True
            TabIndex        =   28
            Text            =   "Text5"
            Top             =   2640
            Width           =   3015
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   4
            Left            =   2550
            Locked          =   -1  'True
            TabIndex        =   27
            Text            =   "Text5"
            Top             =   2265
            Width           =   3015
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   1725
            MaxLength       =   6
            TabIndex        =   9
            Top             =   2265
            Width           =   750
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   5
            Left            =   1725
            MaxLength       =   6
            TabIndex        =   10
            Top             =   2625
            Width           =   750
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   6
            Left            =   1410
            MouseIcon       =   "frmExpImpExcel.frx":000C
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar trabajador"
            Top             =   3300
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   7
            Left            =   1410
            MouseIcon       =   "frmExpImpExcel.frx":015E
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar trabajador"
            Top             =   3690
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Variedad"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   6
            Left            =   420
            TabIndex        =   26
            Top             =   3060
            Width           =   630
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   5
            Left            =   750
            TabIndex        =   25
            Top             =   3660
            Width           =   420
         End
         Begin VB.Label Label2 
            Caption         =   "Desde"
            Height          =   195
            Index           =   4
            Left            =   750
            TabIndex        =   24
            Top             =   3300
            Width           =   465
         End
         Begin VB.Label Label2 
            Caption         =   "Desde"
            Height          =   195
            Index           =   0
            Left            =   810
            TabIndex        =   23
            Top             =   480
            Width           =   465
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   1
            Left            =   810
            TabIndex        =   22
            Top             =   810
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Albarán"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   2
            Left            =   450
            TabIndex        =   21
            Top             =   240
            Width           =   540
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   2
            Left            =   1395
            Picture         =   "frmExpImpExcel.frx":02B0
            Top             =   1350
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   3
            Left            =   1395
            Picture         =   "frmExpImpExcel.frx":033B
            Top             =   1755
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   4
            Left            =   1410
            MouseIcon       =   "frmExpImpExcel.frx":03C6
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar trabajador"
            Top             =   2280
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   1410
            MouseIcon       =   "frmExpImpExcel.frx":0518
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar trabajador"
            Top             =   2670
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   24
            Left            =   420
            TabIndex        =   20
            Top             =   1170
            Width           =   450
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   25
            Left            =   780
            TabIndex        =   19
            Top             =   1725
            Width           =   420
         End
         Begin VB.Label Label2 
            Caption         =   "Desde"
            Height          =   195
            Index           =   26
            Left            =   780
            TabIndex        =   18
            Top             =   1410
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Socio"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   27
            Left            =   420
            TabIndex        =   17
            Top             =   2040
            Width           =   405
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   28
            Left            =   750
            TabIndex        =   16
            Top             =   2640
            Width           =   420
         End
         Begin VB.Label Label2 
            Caption         =   "Desde"
            Height          =   195
            Index           =   29
            Left            =   750
            TabIndex        =   15
            Top             =   2280
            Width           =   465
         End
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   420
         Index           =   0
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "Recolectado|N|N|0|1|rcampos|recolect||N|"
         Top             =   1290
         Width           =   2760
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   4680
         TabIndex        =   14
         Top             =   6075
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3600
         TabIndex        =   13
         Top             =   6060
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Acción"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   3
         Left            =   630
         TabIndex        =   2
         Top             =   1290
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Exportar/Importar fichero Excel "
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
         Left            =   450
         TabIndex        =   1
         Top             =   450
         Width           =   5385
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6690
      Top             =   3780
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmExpImpExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MONICA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(Cliente As String, Observaciones As String)

Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean


Private WithEvents frmSoc As frmManSocios 'mantenimiento de socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmVar As frmComVar 'variedades
Attribute frmVar.VB_VarHelpID = -1

 
Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
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


Dim PrimeraVez As Boolean
Dim Contabilizada As Byte
Dim ConSubInforme As Boolean


Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub


Private Sub cmdAceptar_Click(Index As Integer)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim NomAlmac As String
Dim cTabla As String

    
    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    Select Case Index
        Case 0 ' Proceso de pago de partes de campo
            'D/H Albaran
            cDesde = Trim(txtcodigo(0).Text)
            cHasta = Trim(txtcodigo(1).Text)
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rhisfruta.numalbar}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
            End If
            
            'D/H Fecha
            cDesde = Trim(txtcodigo(2).Text)
            cHasta = Trim(txtcodigo(3).Text)
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rhisfruta.fecalbar}"
                TipCod = "F"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
            End If
            
            'D/H socio
            cDesde = Trim(txtcodigo(4).Text)
            cHasta = Trim(txtcodigo(5).Text)
            nDesde = txtNombre(4).Text
            nHasta = txtNombre(5).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rhisfruta.codsocio}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
            End If
    
            'D/H variedad
            cDesde = Trim(txtcodigo(6).Text)
            cHasta = Trim(txtcodigo(7).Text)
            nDesde = txtNombre(6).Text
            nHasta = txtNombre(7).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rhisfruta.codvarie}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad=""") Then Exit Sub
            End If
   
            If HayRegParaInforme(cTabla, cadSelect) Then
            
                            
            
            
                Shell App.Path & "\gestoria.exe /0"
            
                If ProcesoCargaTemporal(cTabla, cadSelect) Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                    cmdCancel_Click (0)
                    Exit Sub
                Else
                    MsgBox "El Proceso no se ha realizado correctamente. Llame a Ariadna.", vbExclamation
                    Exit Sub
                End If
            End If
    
    
        
    End Select
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(Tabla, cadSelect) Then
        LlamarImprimir
    End If

End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Combo1_Change(Index As Integer)

    Select Case Combo1(0).Text
        Case 0 'exportar
            Frame1.Enabled = True
        
        Case 1 ' importar
            LimpiarCampos
            Frame1.Enabled = False
    End Select
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFocoCmb Combo1(0)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    For H = 4 To 7
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    
    CargaCombo
    
    Me.Combo1(0).ListIndex = 0
    
    Tabla = "rhisfruta"
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(0).Cancel = True
    Me.Width = 6375
    Me.Height = 6935
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtcodigo(CByte(imgFecha(0).Tag) + 14).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 4, 5 'socios
            AbrirFrmManSocios (Index)
    
        Case 6, 7 ' variedad comercial
            AbrirFrmManVariedad (Index)
            
    End Select
    PonerFoco txtcodigo(indCodigo)
End Sub


Private Sub imgFecha_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object

    Set frmC = New frmCal
    
    esq = imgFecha(Index).Left
    dalt = imgFecha(Index).Top
        
    Set obj = imgFecha(Index).Container
      
      While imgFecha(Index).Parent.Name <> obj.Name
            esq = esq + obj.Left
            dalt = dalt + obj.Top
            Set obj = obj.Container
      Wend
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmC.Left = esq + imgFecha(Index).Parent.Left + 30
    frmC.Top = dalt + imgFecha(Index).Parent.Top + imgFecha(Index).Height + menu - 40

    imgFecha(0).Tag = Index '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtcodigo(Index + 14).Text <> "" Then frmC.NovaData = txtcodigo(Index + 14).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtcodigo(CByte(imgFecha(0).Tag) + 14) '<===
    ' ********************************************
End Sub



Private Sub ListView1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
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
            Case 2: KEYBusqueda KeyAscii, 2 'proveedor desde
            Case 3: KEYBusqueda KeyAscii, 3 'proveedor hasta
            Case 4: KEYBusqueda KeyAscii, 4 'cliente desde
            Case 5: KEYBusqueda KeyAscii, 5 'cliente hasta
            Case 6: KEYBusqueda KeyAscii, 6 'producto desde
            Case 7: KEYBusqueda KeyAscii, 7 'producto hasta
            Case 0: KEYBusqueda KeyAscii, 0 'variedad desde
            Case 1: KEYBusqueda KeyAscii, 1 'variedad hasta
            Case 8: KEYBusqueda KeyAscii, 8 'cliente de la factura rectificativa
            Case 18: KEYBusqueda KeyAscii, 14 'trabajador desde
            Case 19: KEYBusqueda KeyAscii, 15 'trabajador hasta
            Case 20: KEYBusqueda KeyAscii, 16 'articulo desde
            Case 21: KEYBusqueda KeyAscii, 17 'articulo hasta
            Case 24: KEYBusqueda KeyAscii, 20 'almacen para el calculo de horas productivas
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
Dim cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1 ' Nro.Albaranes
            PonerFormatoEntero txtcodigo(Index)
    
        Case 2, 3 'FECHAS
            If txtcodigo(Index).Text <> "" Then PonerFormatoFecha txtcodigo(Index)
            
        Case 4, 5 'Socios
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
        
        Case 6, 7 'Variedad
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "variedades", "nomvarie", "codvarie", "N")
    
    End Select
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
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
            cadParam = cadParam & AnyadirParametroDH(param, codD, codH, nomD, nomH)
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function

Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .EnvioEMail = False
        .ConSubInforme = ConSubInforme
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmManVariedad(indice As Integer)
        Set frmVar = New frmComVar
        frmVar.DatosADevolverBusqueda = "0|1|"
        frmVar.Show vbModal
        Set frmVar = Nothing
        PonerFoco txtcodigo(indice)
End Sub


Private Sub AbrirFrmManSocios(indice As Integer)
    indCodigo = indice + 4
    Set frmSoc = New frmManSocios
    frmSoc.DatosADevolverBusqueda = "0|1|"
    frmSoc.Show vbModal
    Set frmSoc = Nothing
End Sub


Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
'        .SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
        .OtrosParametros = cadParam
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

Private Sub PonerValoresFactura()
Dim intconta As String
Dim cad As String
'    txtCodigo(9).Text = RecuperaValor(CadTag, 1)
'    txtCodigo(10).Text = RecuperaValor(CadTag, 2)
'    txtCodigo(11).Text = RecuperaValor(CadTag, 3)
'    txtCodigo(12).Text = RecuperaValor(CadTag, 4)
'    txtNombre(9).Text = RecuperaValor(CadTag, 5)
'    Contabilizada = RecuperaValor(CadTag, 6)
     intconta = "intconta"
     txtcodigo(12).Text = ""
     txtcodigo(12).Text = DevuelveDesdeBDNew(cAgro, "schfac", "codsocio", "letraser", txtcodigo(9).Text, "T", intconta, "numfactu", txtcodigo(10).Text, "N", "fecfactu", txtcodigo(11).Text, "F")
     If txtcodigo(12).Text <> "" Then
        txtNombre(9).Text = PonerNombreDeCod(txtcodigo(12), "ssocio", "nomsocio", "codsocio", "N")
        Contabilizada = CInt(intconta)
     Else
        cad = "No existe la factura. Reintroduzca. " & vbCrLf & vbCrLf
        cad = cad & "   Serie: " & txtcodigo(9).Text & vbCrLf
        cad = cad & "   Factura: " & txtcodigo(10).Text & vbCrLf
        cad = cad & "   Fecha: " & txtcodigo(11).Text & vbCrLf
        cad = cad & vbCrLf
        MsgBox cad, vbExclamation
        txtcodigo(9).Text = ""
        txtcodigo(10).Text = ""
        txtcodigo(11).Text = ""
        PonerFoco txtcodigo(9)
     End If
End Sub


'Private Function DatosOk() As Boolean
'Dim b As Boolean
'Dim Sql As String
'Dim Sql2 As String
'Dim vClien As CSocio
'' añadido
'Dim Mens As String
'Dim numfactu As String
'Dim numser As String
'Dim Fecha As Date
'
'    b = True
'    If txtCodigo(9).Text = "" Or txtCodigo(10).Text = "" Or txtCodigo(11).Text = "" Then
'        MsgBox "Debe introducir la letra de serie, el número de factura y la fecha de factura para localizar la factura a rectificar", vbExclamation
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
'' añadido
''    b = True
'
'    If ConTarjetaProfesional(txtCodigo(9).Text, txtCodigo(10).Text, txtCodigo(11).Text) Then
'        MsgBox "Este Factura tiene alguna tarjeta profesional, no se permite hacer la factura rectificativa", vbExclamation
'        b = False
'    Else
'        If txtCodigo(13).Text = "" Then
'            MsgBox "Debe introducir obligatoriamente una Fecha de Facturación.", vbExclamation
'            b = False
'            PonerFoco txtCodigo(13)
'        Else
'                If Not FechaDentroPeriodoContable(CDate(txtCodigo(13).Text)) Then
'                    Mens = "La Fecha de Facturación no es del ejercicio actual ni siguiente. Reintroduzca."
'                    MsgBox Mens, vbExclamation
'                    b = False
'                    PonerFoco txtCodigo(13)
'                Else
'                    'VRS:2.0.1(0)
'                    If Not FechaSuperiorUltimaLiquidacion(CDate(txtCodigo(13).Text)) Then
'                        Mens = "  La Fecha de Facturación es inferior a la última liquidación de Iva. " & vbCrLf & vbCrLf
'                        ' unicamente si el usuario es root el proceso continuará
'                        If vSesion.Nivel > 0 Then
'                            Mens = Mens & "  El proceso no continuará."
'                            MsgBox Mens, vbExclamation
'                            b = False
'                            PonerFoco txtCodigo(13)
'                        Else
'                            Mens = Mens & "                        ¿ Desea continuar ?    " & vbCrLf
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
'                            Mens = "La Fecha de Factura es inferior a la última factura de la serie." & vbCrLf & vbCrLf
'                            Mens = Mens & "                        ¿ Desea continuar ?    " & vbCrLf
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
'' end añadido
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



Private Function CargarTablaTemporal() As Boolean
Dim Sql As String
Dim SQL1 As String
Dim Rs As ADODB.Recordset

    On Error GoTo eCargarTablaTemporal
    
    CargarTablaTemporal = False

    Sql = "delete from tmpenvasesret where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute Sql

'select albaran_envase.codartic, albaran_envase.fechamov
'from (albaran_envase inner join sartic on albaran_envase.codartic = sartic.codartic) inner join stipar on sartic.codtipar = stipar.codtipar
'Where stipar.esretornable = 1
'Union
'select smoval.codartic, smoval.fechamov
'from (smoval inner join  sartic on smoval.codartic = sartic.codartic) inner join stipar on sartic.codtipar = stipar.codtipar
'Where stipar.esretornable = 1

    Sql = "select " & vUsu.Codigo & ", albaran_envase.codartic, albaran_envase.fechamov, albaran_envase.cantidad, albaran_envase.tipomovi, albaran_envase.numalbar, "
    Sql = Sql & " albaran_envase.codclien, clientes.nomclien, " & DBSet("ALV", "T")
    Sql = Sql & " from ((albaran_envase inner join sartic on albaran_envase.codartic = sartic.codartic) inner join stipar on sartic.codtipar = stipar.codtipar) "
    Sql = Sql & " inner join clientes on albaran_envase.codclien = clientes.codclien "
    Sql = Sql & " where stipar.esretornable = 1 "
    
    If txtcodigo(12).Text <> "" Then Sql = Sql & " and stipar.codtipar >= " & DBSet(txtcodigo(12).Text, "N")
    If txtcodigo(13).Text <> "" Then Sql = Sql & " and stipar.codtipar <= " & DBSet(txtcodigo(13).Text, "N")
    
    If txtcodigo(20).Text <> "" Then Sql = Sql & " and albaran_envase.codartic >= " & DBSet(txtcodigo(20).Text, "T")
    If txtcodigo(21).Text <> "" Then Sql = Sql & " and albaran_envase.codartic <= " & DBSet(txtcodigo(21).Text, "T")
    
    If txtcodigo(22).Text <> "" Then Sql = Sql & " and albaran_envase.codclien >= " & DBSet(txtcodigo(22).Text, "N")
    If txtcodigo(23).Text <> "" Then Sql = Sql & " and albaran_envase.codclien <= " & DBSet(txtcodigo(23).Text, "N")
    
    If txtcodigo(14).Text <> "" Then Sql = Sql & " and albaran_envase.fechamov >= " & DBSet(txtcodigo(14).Text, "F")
    If txtcodigo(15).Text <> "" Then Sql = Sql & " and albaran_envase.fechamov <= " & DBSet(txtcodigo(15).Text, "F")
    
    Sql = Sql & " union "
    
    Sql = Sql & "select " & vUsu.Codigo & ", smoval.codartic, smoval.fechamov, smoval.cantidad, smoval.tipomovi, smoval.document, "
    Sql = Sql & " smoval.codigope, proveedor.nomprove, " & DBSet("ALC", "T")
    Sql = Sql & " from ((smoval inner join sartic on smoval.codartic = sartic.codartic) inner join stipar on sartic.codtipar = stipar.codtipar) "
    Sql = Sql & " inner join proveedor on smoval.codigope = proveedor.codprove "
    Sql = Sql & " where stipar.esretornable = 1 "
    
    If txtcodigo(12).Text <> "" Then Sql = Sql & " and stipar.codtipar >= " & DBSet(txtcodigo(12).Text, "N")
    If txtcodigo(13).Text <> "" Then Sql = Sql & " and stipar.codtipar <= " & DBSet(txtcodigo(13).Text, "N")
    
    If txtcodigo(20).Text <> "" Then Sql = Sql & " and smoval.codartic >= " & DBSet(txtcodigo(20).Text, "T")
    If txtcodigo(21).Text <> "" Then Sql = Sql & " and smoval.codartic <= " & DBSet(txtcodigo(21).Text, "T")
    
    If txtcodigo(14).Text <> "" Then Sql = Sql & " and smoval.fechamov >= " & DBSet(txtcodigo(14).Text, "F")
    If txtcodigo(15).Text <> "" Then Sql = Sql & " and smoval.fechamov <= " & DBSet(txtcodigo(15).Text, "F")

    SQL1 = "insert into tmpenvasesret " & Sql
    conn.Execute SQL1
    
    CargarTablaTemporal = True
    Exit Function
    
eCargarTablaTemporal:
    MuestraError Err.Number, "Carga Tabla Temporal"
End Function

Private Function CalculoHorasProductivas() As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim SQL1 As String

    On Error GoTo eCalculoHorasProductivas

    CalculoHorasProductivas = False

    Sql = "fechahora = " & DBSet(txtcodigo(27).Text, "F") & " and codalmac = " & DBSet(txtcodigo(24), "N")
    Sql = Sql & " and codtraba in (select codtraba from straba where codsecci = 1)"


    If BloqueaRegistro("horas", Sql) Then
        SQL1 = "update horas set horasproduc = round(horasdia * (1 + (" & DBSet(txtcodigo(25), "N") & "/ 100)),2) "
        SQL1 = SQL1 & " where fechahora = " & DBSet(txtcodigo(27).Text, "F")
        SQL1 = SQL1 & " and codalmac = " & DBSet(txtcodigo(24), "N")
        SQL1 = SQL1 & " and codtraba in (select codtraba from straba where codsecci = 1) "
        
        conn.Execute SQL1
    
        CalculoHorasProductivas = True
    End If

    TerminaBloquear
    Exit Function

eCalculoHorasProductivas:
    MuestraError Err.Number, "Calculo Horas Productivas", Err.Description
    TerminaBloquear
End Function


Private Function ProcesoCargaTemporal(cTabla As String, cWhere As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset

    On Error GoTo eProcesoCargaHoras
    
    Screen.MousePointer = vbHourglass
    
    Sql = "CARCLA" 'carga de clasificacion
    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se puede realizar el proceso de Carga de Clasificación. Hay otro usuario realizándolo.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    ProcesoCargaTemporal = False

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "Select rpartes_trabajador.nroparte, rpartes.fechapar, rpartes_trabajador.codtraba, sum(rpartes_trabajador.importe) FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    Sql = Sql & " group by 1, 2, 3"
    Sql = Sql & " order by 1, 2, 3"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
    Sql = "insert into horas (codtraba, fechahora, horasdia, horasproduc, compleme,"
    Sql = Sql & "intconta, pasaridoc, codalmac, nroparte) values "
        
    Sql3 = ""
    While Not Rs.EOF
        Sql2 = "select count(*) from horas where fechahora = " & DBSet(Rs.Fields(1).Value, "F")
        Sql2 = Sql2 & " and codtraba = " & DBSet(Rs.Fields(2).Value, "N")
        Sql2 = Sql2 & " and codalmac = " & DBSet(vParamAplic.AlmacenNOMI, "N")
        
        If TotalRegistros(Sql2) = 0 Then
            Sql3 = Sql3 & "(" & DBSet(Rs.Fields(2).Value, "N") & "," & DBSet(Rs.Fields(1).Value, "F") & ",0,0,"
            Sql3 = Sql3 & DBSet(Rs.Fields(3).Value, "N") & ",0,0," & DBSet(vParamAplic.AlmacenNOMI, "N") & ","
            Sql3 = Sql3 & DBSet(Rs.Fields(0).Value, "N") & "),"
        End If
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    If Sql3 <> "" Then
        ' quitamos la ultima coma
        Sql3 = Mid(Sql3, 1, Len(Sql3) - 1)
        Sql = Sql & Sql3
        
        conn.Execute Sql
    End If
    
    DesBloqueoManual ("CARNOM") 'carga de nominas
    
    Screen.MousePointer = vbDefault
    
    ProcesoCargaTemporal = True
    Exit Function
    
eProcesoCargaHoras:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Proceso de Carga de Horas", Err.Description
End Function



Private Sub CargaCombo()
        
    Combo1(0).Clear
    'tipo de recoleccion
    Combo1(0).AddItem "Exportar"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Importar"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
End Sub
    

Private Sub LimpiarCampos()
Dim I As Byte

    For I = 0 To txtcodigo.Count - 1
        txtcodigo(I).Text = ""
    Next I

End Sub
