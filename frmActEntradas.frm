VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmActEntradas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualizar Entradas de Báscula"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6750
   Icon            =   "frmActEntradas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameEntradasCampo 
      Height          =   4455
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   6615
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Text5"
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Text5"
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1755
         MaxLength       =   3
         TabIndex        =   2
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1755
         MaxLength       =   3
         TabIndex        =   1
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Height          =   440
         Left            =   7860
         Picture         =   "frmActEntradas.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command3 
         Height          =   440
         Left            =   7860
         Picture         =   "frmActEntradas.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text5"
         Top             =   1950
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text5"
         Top             =   1590
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   4
         Top             =   1950
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   3
         Top             =   1590
         Width           =   735
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   4080
         TabIndex        =   7
         Top             =   3765
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   5100
         TabIndex        =   8
         Top             =   3780
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   6
         Top             =   2925
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   5
         Top             =   2520
         Width           =   1095
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   450
         TabIndex        =   24
         Top             =   3420
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProgres 
         Caption         =   "Label1"
         Height          =   240
         Left            =   450
         TabIndex        =   25
         Top             =   3735
         Width           =   3255
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1440
         MouseIcon       =   "frmActEntradas.frx":0620
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   990
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1440
         MouseIcon       =   "frmActEntradas.frx":0772
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   28
         Left            =   825
         TabIndex        =   23
         Top             =   1035
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   18
         Left            =   825
         TabIndex        =   22
         Top             =   645
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clase"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   11
         Left            =   495
         TabIndex        =   21
         Top             =   405
         Width           =   390
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1440
         Picture         =   "frmActEntradas.frx":08C4
         ToolTipText     =   "Buscar fecha"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1440
         Picture         =   "frmActEntradas.frx":094F
         ToolTipText     =   "Buscar fecha"
         Top             =   2925
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1440
         MouseIcon       =   "frmActEntradas.frx":09DA
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1980
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1440
         MouseIcon       =   "frmActEntradas.frx":0B2C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1590
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   26
         Left            =   825
         TabIndex        =   18
         Top             =   2025
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   25
         Left            =   825
         TabIndex        =   17
         Top             =   1635
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   24
         Left            =   495
         TabIndex        =   16
         Top             =   1395
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   21
         Left            =   825
         TabIndex        =   15
         Top             =   2925
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   20
         Left            =   825
         TabIndex        =   14
         Top             =   2580
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   19
         Left            =   495
         TabIndex        =   13
         Top             =   2340
         Width           =   450
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6075
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmActEntradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: LAURA +-+-
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


Private WithEvents frmSec As frmManSeccion 'Secciones
Attribute frmSec.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'Socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmVar As frmComVar 'Variedades de comercial
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmCal As frmManCalidades 'Calidades
Attribute frmCal.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmSit As frmManSituCamp 'Situacion campos
Attribute frmSit.VB_VarHelpID = -1
Private WithEvents frmCla As frmComercial 'Ayuda de Clases de comercial
Attribute frmCla.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'Mensajes
Attribute frmMens.VB_VarHelpID = -1


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

Dim PrimeraVez As Boolean
Dim Contabilizada As Byte
Dim OK As Boolean

Dim CadFormulaImp As String

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub

Private Function DatosOK() As Boolean
Dim B As Boolean

    DatosOK = False
    
    If vParamAplic.Cooperativa = 2 Then
        If txtCodigo(4).Text = "" Or txtCodigo(5).Text = "" Then
            MsgBox "Debe introducir las fechas. Revise.", vbExclamation
            PonerFoco txtCodigo(4)
            Exit Function
        Else
            If txtCodigo(4).Text <> txtCodigo(5).Text Then
                MsgBox "Debe introducir la misma fecha desde y hasta. Revise.", vbExclamation
                PonerFoco txtCodigo(4)
                Exit Function
            End If
        End If
        
        If vParamAplic.PathEntradas = "" Then
            MsgBox "No está configurado el path de impresión de entradas. Revise.", vbExclamation
            Exit Function
        Else
            If Dir(vParamAplic.PathEntradas & "\", vbDirectory) = "" Then
                MsgBox "No existe el directorio seleccionado para impresión de entradas. Revise.", vbExclamation
                Exit Function
            End If
        End If
        
    End If
    
    DatosOK = True
    
End Function


Private Sub cmdAceptar_Click(Index As Integer)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim HayReg As Boolean
Dim cadena As String
    
    '[Monica]17/10/2016: si es Picassent obligamos a que me pongan una fecha de maximo 1 dia
    If Not DatosOK Then Exit Sub
    
    
    InicializarVbles
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    Select Case Index
        Case 0
            '======== FORMULA  ====================================
            'D/H CLASE
            cDesde = Trim(txtCodigo(0).Text)
            cHasta = Trim(txtCodigo(1).Text)
            nDesde = txtNombre(0).Text
            nHasta = txtNombre(1).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{variedades.codclase}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
            End If
            
            'D/H VARIEDAD
            cDesde = Trim(txtCodigo(2).Text)
            cHasta = Trim(txtCodigo(3).Text)
            nDesde = txtNombre(2).Text
            nHasta = txtNombre(3).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rentradas.codvarie}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad=""") Then Exit Sub
            End If

            CadFormulaImp = cadFormula

            'D/H fecha
            cDesde = Trim(txtCodigo(4).Text)
            cHasta = Trim(txtCodigo(5).Text)
            nDesde = ""
            nHasta = ""
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rentradas.fechaent}"
                TipCod = "F"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
            End If
            
            
            tabla = "(rentradas INNER JOIN variedades ON rentradas.codvarie = variedades.codvarie) "

            'Comprobar si hay registros a Mostrar antes de abrir el Informe
            If HayRegParaInforme(tabla, cadSelect) Then
            
                '[Monica]25/03/2014: para el caso de que no haya ausencia de plagas (entradas de quatretonda)
                Dim cadSelect2 As String
                Dim Sql4 As String
                If cadSelect <> "" Then
                    cadSelect2 = cadSelect & " and rentradas.ausenciaplagas = 0"
                Else
                    cadSelect2 = cadSelect & "rentradas.ausenciaplagas = 0"
                End If
                Sql4 = "select count(*) from " & tabla & " where " & cadSelect2
                If TotalRegistros(Sql4) <> 0 And vParamAplic.CodIncidPlaga = 0 Then
                    MsgBox "Debe introducir un código de incidencia de plaga en parámetros.", vbExclamation
                    Exit Sub
                End If
            
                If vParamAplic.HayTraza Then
                    cadena = ""
                    HayReg = HayEntradasSinCRFID(tabla, cadSelect, cadena)
                    
                    If HayReg Then
                        Set frmMens = New frmMensajes
                        frmMens.OpcionMensaje = 21
                        frmMens.cadena = cadena
                        frmMens.Show vbModal
                        Set frmMens = Nothing
                        '[Monica]10/01/2011:añadida la pregunta para que se puedan actualizar entradas que ya han sido volcadas
                        If MsgBox("¿Desea continuar con la actualización?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                            Exit Sub
                        End If
                    End If
                End If
            
            
                If ActualizarTabla(tabla, cadSelect) Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                    cmdCancel_Click (0)
                End If
            End If
    End Select
End Sub


Private Sub cmdCancel_Click(Index As Integer)
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

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me

    For H = 0 To imgBuscar.Count - 1
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    
    Me.Pb1.visible = False
    Me.lblProgres.visible = False
    
    tabla = "rentradas"
 
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtCodigo(CByte(imgFec(0).Tag) + 4).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmCla_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") ' codigo de clase
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) ' descripcion
End Sub




Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 ' Clase
            AbrirFrmClase (Index)
        
        Case 2, 3 'VARIEDADES
            AbrirFrmVariedad (Index)
    
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

    imgFec(0).Tag = Index '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtCodigo(Index + 4).Text <> "" Then frmC.NovaData = txtCodigo(Index + 4).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtCodigo(CByte(imgFec(0).Tag) + 4) '<===
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
            Case 0: KEYBusqueda KeyAscii, 0 'clase desde
            Case 1: KEYBusqueda KeyAscii, 1 'clase hasta
            Case 2: KEYBusqueda KeyAscii, 2 'variedad desde
            Case 3: KEYBusqueda KeyAscii, 3 'variedad hasta
            Case 4: KEYFecha KeyAscii, 4 'fecha desde
            Case 5: KEYFecha KeyAscii, 5 'fecha hasta
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFec_Click (Indice)
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1 'CLASES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "clases", "nomclase", "codclase", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
    
        Case 4, 5 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
            '[Monica]17/10/2016: obligamos a meter la misma fecha si es Picassent
            If vParamAplic.Cooperativa = 2 Then
                If Index = 4 Then txtCodigo(5).Text = txtCodigo(4).Text
            End If
            
        Case 2, 3 'VARIEDADES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
    End Select
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
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmClase(Indice As Integer)
    indCodigo = Indice
    Set frmCla = New frmComercial
    
    AyudaClasesCom frmCla, txtCodigo(Indice).Text
    
    Set frmCla = Nothing
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

Private Function ActualizarTabla(cTabla As String, cWhere As String) As Boolean
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim cadMen As String
Dim I As Long
Dim B As Boolean
Dim CalidadVC As String
Dim CalidadDES As String

Dim Pesadas As String
Dim Retirada As Boolean
Dim Destrio As Boolean
Dim PorcenDestrio As Currency

Dim fr As frmVisReport

    On Error GoTo eActualizarTabla
    
    ActualizarTabla = False

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    SQL = "Select rentradas.* FROM " & QuitarCaracterACadena(cTabla, "_1")
    Sql1 = "select count(*) from " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
        Sql1 = Sql1 & " WHERE " & cWhere
    End If
    
    Pb1.visible = True
    lblProgres.visible = True
    
    Me.Pb1.Max = TotalRegistros(Sql1)
    Me.Refresh
    Me.Pb1.Value = 0
    
    BorrarTMPErr
    CrearTMPErr
    
    conn.BeginTrans
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
    '[Monica]17/10/2016: para el caso de Picassent vamos a imprimir las todas las entradas en un fichero pdf  yyyymmddhhmmss.pdf
'*****
    If vParamAplic.Cooperativa = 2 Then
    
        lblProgres.Caption = "Impresión de entradas"
        DoEvents
    
    
        Set fr = New frmVisReport
        
        Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
        Dim nomDocu As String 'Nombre de Informe rpt de crystal
        Dim ImprimeDirecto As Integer
        
        
        '++monica: seleccionamos que rpt se ha de ejecutar
    
    '            cadParam = "pEmpresa=""AriagroRec""|"
        CadParam = ""
        numParam = 1
        
        indRPT = 25
        If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then
            conn.RollbackTrans
            Exit Function
        End If
        '++
        fr.NumeroParametros = numParam
        fr.OtrosParametros = CadParam
        fr.ConSubInforme = True
        fr.Informe = App.Path & "\Informes\" & nomDocu
        fr.FormulaSeleccion = "{rentradas.fechaent} = Date(" & Mid(txtCodigo(4).Text, 7, 4) & _
                                                    "," & Mid(txtCodigo(4).Text, 4, 2) & _
                                                    "," & Mid(txtCodigo(4).Text, 1, 2) & ")"
                                                    
        If CadFormulaImp <> "" Then fr.FormulaSeleccion = fr.FormulaSeleccion & " and " & CadFormulaImp
        
        fr.FicheroPDF = vParamAplic.PathEntradas & "\" & Mid(txtCodigo(4), 7, 4) & Mid(txtCodigo(4), 4, 2) & Mid(txtCodigo(4), 1, 2) & "_" & Format(Now, "hhmmss") & ".pdf"
        Load fr 'trabaja sin mostrar el formulario
        
    End If
'*****
    
    Pesadas = "("
    
    I = 0
    B = True
    While Not Rs.EOF And B
        I = I + 1
        
        Me.Pb1.Value = Me.Pb1.Value + 1
        lblProgres.Caption = "Linea: " & I & ". Entrada: " & Format(DBLet(Rs!numnotac, "N"), "00000000")
        Me.Refresh

        OK = True
        
        If DBLet(Rs!TipoEntr, "N") = 4 Then
            ' si es una entrada de RETIRADA va todo sobre esta calidad - el porcentaje de destrio
            ' de momento se utiliza en el mantenimiento de entradas de quatretonda
            CalidadVC = CalidadRetirada(CStr(DBLet(Rs!codvarie, "N")))
            If CalidadVC = "" Then
                Retirada = True
            
                Sql1 = "insert into tmpErrEnt (numnotac,codvarie) values ( " & DBSet(Rs!numnotac, "N")
                Sql1 = Sql1 & "," & DBSet(Rs!codvarie, "N") & " )"
                conn.Execute Sql1
            Else
                '[Monica] si hay porcentaje de destrio en la variedad miro a ver si hay calidad de destrio
                PorcenDestrio = DevuelveValor("select eurotria from variedades where codvarie = " & DBSet(Rs!codvarie, "N"))
                If PorcenDestrio <> 0 Then
                    CalidadDES = CalidadDestrio(CStr(DBLet(Rs!codvarie, "N")))
                    If CalidadDES = "" Then
                        Destrio = True
                    
                        Sql1 = "insert into tmpErrEnt (numnotac,codvarie) values ( " & DBSet(Rs!numnotac, "N")
                        Sql1 = Sql1 & "," & DBSet(Rs!codvarie, "N") & " )"
                        conn.Execute Sql1
                    End If
                End If
                
                If PorcenDestrio = 0 Or (PorcenDestrio <> 0 And CalidadDES <> "") Then
                    B = InsertarCabecera(Rs, cadMen)
                    cadMen = "Insertando Cabecera: " & cadMen
                
                    If B Then
                        cadMen = ""
                        
                        If PorcenDestrio = 0 Then
                            B = InsertarClasificacion(Rs, cadMen, CalidadVC)
                        Else
                            B = InsertarClasificacionConDestrio(Rs, cadMen, CalidadVC, CalidadDES, CStr(PorcenDestrio))
                        End If
                        
                        cadMen = "Insertando Clasificacion: " & cadMen
                    End If
                    
                    '[Monica]04/05/2010 Reparto de albaranes
                    If B And vParamAplic.CooproenEntradas Then
                        B = RepartoAlbaranesBascula(Rs!numnotac, cadMen)
                        cadMen = "Reparto Coopropietarios: " & cadMen
                    End If
                    
                    Pesadas = Pesadas & DBSet(Rs!nropesada, "N") & ","
                    
                    'Eliminamos la entrada
                    If B Then
                        cadMen = ""
                        B = EliminarRegistro(Rs, cadMen)
                        cadMen = "Eliminando Registro: " & cadMen
                    End If
                End If
            End If
        Else
            If DBLet(Rs!TipoEntr, "N") <> 1 Then ' si no es VC clasificamos por campo o almacen
                B = InsertarCabecera(Rs, cadMen)
                cadMen = "Insertando Cabecera: " & cadMen
                
                If B Then
                    cadMen = ""
                    B = InsertarClasificacion(Rs, cadMen, "")
                    cadMen = "Insertando Clasificacion: " & cadMen
                End If
            
                If B Then
                    cadMen = ""
                    B = ActualizarTransporte(Rs, cadMen)
                    cadMen = "Actualizando Gastos de Transporte" & cadMen
                End If
                
                '[Monica]04/05/2010 Reparto de albaranes
                If B And vParamAplic.CooproenEntradas Then
                    B = RepartoAlbaranesBascula(Rs!numnotac, cadMen)
                    cadMen = "Reparto Coopropietarios: " & cadMen
                End If
                
                Pesadas = Pesadas & DBSet(Rs!nropesada, "N") & ","
                
                'Eliminamos la entrada
                If B Then
                    cadMen = ""
                    B = EliminarRegistro(Rs, cadMen)
                    cadMen = "Eliminando Registro: " & cadMen
                End If
            Else   ' si es venta campo todos los kilos iran a la calidad de venta campo
                CalidadVC = CalidadVentaCampo(CStr(DBLet(Rs!codvarie, "N")))
                If CalidadVC = "" Then
                    Sql1 = "insert into tmpErrEnt (numnotac,codvarie) values ( " & DBSet(Rs!numnotac, "N")
                    Sql1 = Sql1 & "," & DBSet(Rs!codvarie, "N") & " )"
                    conn.Execute Sql1
                Else
                    B = InsertarCabecera(Rs, cadMen)
                    cadMen = "Insertando Cabecera: " & cadMen
                    
                    If B Then
                        cadMen = ""
                        B = InsertarClasificacion(Rs, cadMen, CalidadVC)
                        cadMen = "Insertando Clasificacion: " & cadMen
                    End If
                    
                    '[Monica]04/05/2010 Reparto de albaranes
                    If B And vParamAplic.CooproenEntradas Then
                        B = RepartoAlbaranesBascula(Rs!numnotac, cadMen)
                        cadMen = "Reparto Coopropietarios: " & cadMen
                    End If
                    
                    Pesadas = Pesadas & DBSet(Rs!nropesada, "N") & ","
                    
                    'Eliminamos la entrada
                    If B Then
                        cadMen = ""
                        B = EliminarRegistro(Rs, cadMen)
                        cadMen = "Eliminando Registro: " & cadMen
                    End If
                End If
            End If
        End If
        Rs.MoveNext
    Wend

    If B And Len(Pesadas) > 1 Then
        'quitamos la ultima coma y añadimos un parentesis de cierre
        Pesadas = Mid(Pesadas, 1, Len(Pesadas) - 1) & ")"
        
        cadMen = ""
        B = EliminarPesada(Pesadas, cadMen)
        cadMen = "Eliminando Pesadas: " & cadMen
    End If


    If B Then
        If TotalRegistros("select count(*) from tmpErrEnt") <> 0 Then
            Set frmMens = New frmMensajes
            
            If Retirada Then
                frmMens.campo = "/retirada"
            Else
                If Destrio Then
                    frmMens.campo = "/destrio"
                End If
            End If
            
            frmMens.OpcionMensaje = 20
            frmMens.Show vbModal
            Set frmMens = Nothing
        End If
    End If
    
eActualizarTabla:
    If Err.Number <> 0 Or Not B Then
        B = False
        MuestraError Err.Number, "Actualizando Entrada: " & vbCrLf & Err.Description & cadMen
    End If
    If B Then
        conn.CommitTrans
        ActualizarTabla = True
    Else
        conn.RollbackTrans
        ActualizarTabla = False
    End If
End Function


Private Function InsertarCabecera(ByRef Rs As ADODB.Recordset, cadErr As String) As Boolean
'Insertando en tabla conta.cabfact
Dim SQL As String
Dim Sql1 As String
Dim RS1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim cad As String
Dim NumCajones As Currency
Dim Transporte As Currency
Dim vPrecio As String
Dim Precio As Currency

    On Error GoTo EInsertar
    
    SQL = "insert into rclasifica (numnotac,fechaent,horaentr,codvarie,codsocio,codcampo,tipoentr,recolect,codtrans,codcapat,"
    SQL = SQL & "codtarif,kilosbru,numcajon,kilosnet,observac,transportadopor,"
    SQL = SQL & "imptrans,impacarr,imprecol,imppenal,tiporecol,horastra,numtraba,numalbar,fecalbar,impreso,kilostra,contrato) values "

    Sql1 = "select imptrans from rportespobla, rpartida, rcampos, variedades "
    Sql1 = Sql1 & " where rpartida.codparti = rcampos.codparti and "
    Sql1 = Sql1 & " variedades.codprodu = rportespobla.codprodu and "
    Sql1 = Sql1 & " rpartida.codpobla = rportespobla.codpobla and "
    Sql1 = Sql1 & " variedades.codvarie = " & DBSet(Rs!codvarie, "N") & " and "
    Sql1 = Sql1 & " rcampos.codcampo = " & DBSet(Rs!codcampo, "N") & " and "
    Sql1 = Sql1 & " rcampos.codvarie = variedades.codvarie "
    
    Set Rs2 = New ADODB.Recordset
    Rs2.Open Sql1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Precio = 0
    If Not Rs2.EOF Then
        Precio = DBLet(Rs2.Fields(0).Value, "N")
    End If
    
    Set Rs2 = Nothing
    
    
    '[Monica]13/06/2014: para el caso de natural el tipo de envase me indica si es caja o no
    If vParamAplic.Cooperativa = 9 Then
        NumCajones = 0
        If EsCaja(DBLet(Rs!tipocajo1, "N")) Then NumCajones = NumCajones + DBLet(Rs!numcajo1, "N")
        If EsCaja(DBLet(Rs!tipocajo2, "N")) Then NumCajones = NumCajones + DBLet(Rs!numcajo2, "N")
        If EsCaja(DBLet(Rs!tipocajo3, "N")) Then NumCajones = NumCajones + DBLet(Rs!numcajo3, "N")
        If EsCaja(DBLet(Rs!tipocajo4, "N")) Then NumCajones = NumCajones + DBLet(Rs!numcajo4, "N")
        If EsCaja(DBLet(Rs!tipocajo5, "N")) Then NumCajones = NumCajones + DBLet(Rs!numcajo5, "N")
        
    Else
        
    '    NumCajones = DBLet(rs!numcajo1, "N") + DBLet(rs!numcajo2, "N") + DBLet(rs!numcajo3, "N") + DBLet(rs!numcajo4, "N") + DBLet(rs!numcajo5, "N")
    ' 05-05-2009: cambiado por esto
        NumCajones = 0
        If vParamAplic.EsCaja1 Then NumCajones = NumCajones + DBLet(Rs!numcajo1, "N")
        If vParamAplic.EsCaja2 Then NumCajones = NumCajones + DBLet(Rs!numcajo2, "N")
        If vParamAplic.EsCaja3 Then NumCajones = NumCajones + DBLet(Rs!numcajo3, "N")
        If vParamAplic.EsCaja4 Then NumCajones = NumCajones + DBLet(Rs!numcajo4, "N")
        If vParamAplic.EsCaja5 Then NumCajones = NumCajones + DBLet(Rs!numcajo5, "N")
    End If
        
    Transporte = Round2(DBLet(Rs!KilosNet, "N") * Precio, 2)
    
    SQL = SQL & "(" & DBSet(Rs!numnotac, "N") & ","
    SQL = SQL & DBSet(Rs!FechaEnt, "F") & ","
    SQL = SQL & DBSet(Rs!horaentr, "FH") & ","
    SQL = SQL & DBSet(Rs!codvarie, "N") & ","
    SQL = SQL & DBSet(Rs!Codsocio, "N") & ","
    SQL = SQL & DBSet(Rs!codcampo, "N") & ","
    SQL = SQL & DBSet(Rs!TipoEntr, "N") & ","
    SQL = SQL & DBSet(Rs!Recolect, "N") & ","
    SQL = SQL & DBSet(Rs!codTrans, "T") & ","  ', "S") & "," [Monica] si es 0 metemos un 0
    SQL = SQL & DBSet(Rs!codcapat, "N") & ","  ', "S") & "," en codtrans, codcapat, codtarif
    SQL = SQL & DBSet(Rs!Codtarif, "N") & ","  ', "S") & ","
    SQL = SQL & DBSet(Rs!KilosBru, "N") & ","
    SQL = SQL & DBSet(NumCajones, "N") & ","
    SQL = SQL & DBSet(Rs!KilosNet, "N") & ","
    SQL = SQL & ValorNulo & ","
    SQL = SQL & DBSet(Rs!transportadopor, "N") & ","
    SQL = SQL & DBSet(Transporte, "N") & ","
    SQL = SQL & ValorNulo & ","
    SQL = SQL & ValorNulo & ","
    SQL = SQL & ValorNulo & ","
    SQL = SQL & "0," 'tiporecol 0=horas 1=destajo no admite valor nulo
    SQL = SQL & ValorNulo & ","
    SQL = SQL & ValorNulo & ","
    '[Monica]09/03/2017: metemos el nro de albarán si lo tenemos en la entrada (SOLO CASO NATURAL) antes metiamos ValorNulo
    SQL = SQL & DBSet(Rs!numalbar, "N", "S") & ","
    SQL = SQL & ValorNulo & ","
    SQL = SQL & "0," & DBSet(Rs!KilosTra, "N") & ","
    '[Monica]04/10/2016: nueva columna de documento COOPIC
    SQL = SQL & DBSet(Rs!contrato, "T") & ")"
    
    conn.Execute SQL
    
    '[Monica]25/03/2014: en el caso de que no haya ausencia de plagas en la entrada (Quatretonda) se inserta la incidencia
    If DBLet(Rs!ausenciaplagas, "N") = 0 Then
        SQL = "select count(*) from rclasifica_incidencia where numnotac = " & DBSet(Rs!numnotac, "N") & " and codincid = " & DBSet(vParamAplic.CodIncidPlaga, "N")
        If TotalRegistros(SQL) = 0 Then
            SQL = "insert into rclasifica_incidencia (numnotac, codincid)  values (" & DBSet(Rs!numnotac, "N") & "," & DBSet(vParamAplic.CodIncidPlaga, "N") & ")"
        End If
        
        conn.Execute SQL
    End If
    
    
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabecera = False
        cadErr = Err.Description
    Else
        InsertarCabecera = True
    End If
End Function

Private Function EsCaja(CodCaja As String) As Boolean
Dim SQL As String

    SQL = "select escaja from confenva where codtipen = " & DBSet(CodCaja, "N")
    EsCaja = (DevuelveValor(SQL) = 1)


End Function

'Private Function InsertarClasificacion(ByRef Rs As ADODB.Recordset, cadErr As String, vCalidad As String) As Boolean
''Insertando en tabla conta.cabfact
'Dim sql As String
'Dim Sql1 As String
'Dim RS1 As ADODB.Recordset
'Dim Cad As String
'Dim KilosMuestra As Currency
'Dim TotalKilos As Currency
'Dim Calidad As Currency
'Dim Diferencia As Currency
'Dim HayReg As Byte
'Dim TipoClasif As Byte
'Dim vTipoClasif As String
'Dim vCalidDest As String
'Dim CalidadClasif As String
'Dim CalidadVC As String
'
'    On Error GoTo EInsertar
'
'    sql = "insert into rclasifica_clasif (numnotac,codvarie, codcalid, muestra, kilosnet) values "
'
'    If vCalidad <> "" Then
'        Sql1 = "insert into rclasifica_clasif (numnotac, codvarie, codcalid, muestra, kilosnet) "
'        Sql1 = Sql1 & "values (" & DBSet(Rs!numnotac, "N") & "," & DBSet(Rs!CodVarie, "N") & ","
'        Sql1 = Sql1 & DBSet(vCalidad, "N") & ",100," & DBSet(Rs!KilosNet, "N") & ")"
'
'        conn.Execute Sql1
'        InsertarClasificacion = True
'        Exit Function
'    End If
'
'
'
'    vTipoClasif = ""
'    vTipoClasif = DevuelveDesdeBDNew(cAgro, "variedades", "tipoclasifica", "codvarie", Rs!CodVarie, "N")
'
'    If CByte(vTipoClasif) = 0 Then ' clasificacion por campo
'
'        Sql1 = "select rcampos_clasif.* from rcampos_clasif where codcampo = " & DBLet(Rs!codcampo, "N")
'
'        Set RS1 = New ADODB.Recordset
'        RS1.Open Sql1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'        If Not RS1.EOF Then
'            Cad = ""
'
'            TotalKilos = 0
'            HayReg = 0
'
'            While Not RS1.EOF
'                HayReg = 1
'
'                KilosMuestra = Round2(DBLet(Rs!KilosNet, "N") * DBLet(RS1!Muestra, "N") / 100, 0)
'                TotalKilos = TotalKilos + KilosMuestra
'
'                Cad = Cad & "(" & DBSet(Rs!numnotac, "N") & "," & DBSet(Rs!CodVarie, "N") & ","
'                Cad = Cad & DBSet(RS1!codcalid, "N") & "," & DBSet(RS1!Muestra, "N") & ","
'                Cad = Cad & DBSet(KilosMuestra, "N") & "),"
'
'                Calidad = DBLet(RS1!codcalid, "N")
'
'                RS1.MoveNext
'            Wend
'
'            Set RS1 = Nothing
'
'            If HayReg = 1 Then
'                ' quitamos la ultima coma de la cadena
'                If Cad <> "" Then
'                    Cad = Mid(Cad, 1, Len(Cad) - 1)
'                End If
'
'                sql = sql & Cad
'
'                conn.Execute sql
'
'                ' si el kilosneto es diferente a la suma de totalkilos actualizamos la ultima linea
'                If TotalKilos <> DBLet(Rs!KilosNet, "N") Then
'                    Diferencia = DBLet(Rs!KilosNet, "N") - TotalKilos
'
'                    vCalidDest = CalidadDestrioenClasificacion(CStr(Rs!CodVarie), CStr(Rs!numnotac))
'                    If vCalidDest <> "" Then Calidad = vCalidDest
'
'                    sql = "update rclasifica_clasif set kilosnet = kilosnet + (" & DBSet(Diferencia, "N") & ")"
'                    sql = sql & " where numnotac = " & DBSet(Rs!numnotac, "N")
'                    sql = sql & " and codvarie = " & DBSet(Rs!CodVarie, "N")
'                    sql = sql & " and codcalid = " & DBSet(Calidad, "N")
'
'                    conn.Execute sql
'                End If
'            End If
'        Else
'            ' el campo no tiene la clasificacion
'            cadErr = "El campo " & DBLet(Rs!codcampo, "N") & " no tiene clasificación. Revise."
'            InsertarClasificacion = False
'            Exit Function
'
'        End If
'    Else
'        ' la clasificacion es en almacen luego insertamos tantos registros como calidades
'        ' tenga la variedad
'        Sql1 = "insert into rclasifica_clasif (numnotac, codvarie, codcalid, muestra, kilosnet) "
'        Sql1 = Sql1 & "select " & DBSet(Rs!numnotac, "N") & ",rcalidad.codvarie, rcalidad.codcalid, " & ValorNulo & "," & ValorNulo & " from rcalidad where codvarie = " & DBLet(Rs!CodVarie, "N")
'
'        conn.Execute Sql1
'
'    End If
'EInsertar:
'    If Err.Number <> 0 Then
'        InsertarClasificacion = False
'        cadErr = Err.Description
'    Else
'        InsertarClasificacion = True
'    End If
'End Function



Private Function EliminarRegistro(ByRef Rs As ADODB.Recordset, cadErr As String) As Boolean
'Insertando en tabla conta.cabfact
Dim SQL As String
Dim RS1 As ADODB.Recordset
Dim cad As String
Dim NumCajones As Currency
Dim Transporte As Currency
Dim vPrecio As String

    On Error GoTo EEliminar
    
'    SQL = "delete from trzpalets where numnotac = " & DBSet(Rs!numnotac, "N")
'    Conn.Execute SQL
'
    If Not Rs.EOF Then
        SQL = "delete from rentradas where numnotac = " & DBSet(Rs!numnotac, "N")
        conn.Execute SQL
    End If

EEliminar:
    If Err.Number <> 0 Then
        EliminarRegistro = False
        cadErr = Err.Description
    Else
        EliminarRegistro = True
    End If
End Function


Private Function EliminarPesada(cadena As String, cadErr As String) As Boolean
Dim SQL As String

    On Error GoTo EEliminar
    
    If cadena <> "" Then
        SQL = "delete from rpesadas where nropesada in " & Trim(cadena)
        conn.Execute SQL
    End If

EEliminar:
    If Err.Number <> 0 Then
        EliminarPesada = False
        cadErr = Err.Description
    Else
        EliminarPesada = True
    End If
End Function


Private Function InsertarClasificacionVC(ByRef Rs As ADODB.Recordset, cadErr As String, OK As Boolean) As Boolean
'Dim Sql As String
'Dim Sql1 As String
'Dim Rs1 As ADODB.Recordset
'Dim HayReg As Byte
'Dim CalidadVC As String
'
'
'    On Error GoTo EInsertar
'
'    Sql = "insert into rclasifica_clasif (numnotac,codvarie, codcalid, muestra, kilosnet) values "
'
'    CalidadVC = CalidadVentaCampo(CStr(DBLet(Rs!CodVarie, "N")))
'    If CalidadVC = "" Then
'        OK = False
'
'        Sql1 = "insert into tmpErrEnt (numnotac,codvarie) values ( " & DBSet(Rs!numnotac, "N")
'        Sql1 = Sql1 & "," & DBSet(Rs!CodVarie, "N") & " )"
'        conn.Execute Sql1
'
'    Else
'        Sql1 = "insert into rclasifica_clasif (numnotac, codvarie, codcalid, muestra, kilosnet) "
'        Sql1 = Sql1 & "values (" & DBSet(Rs!numnotac, "N") & "," & DBSet(Rs!CodVarie, "N") & ","
'        Sql1 = Sql1 & DBSet(CalidadVC, "N") & ",100," & DBSet(Rs!KilosNet, "N") & ")"
'
'        conn.Execute Sql1
'    End If
'EInsertar:
'    If Err.Number <> 0 Then
'        InsertarClasificacionVC = False
'        cadErr = Err.Description
'    Else
'        InsertarClasificacionVC = True
'    End If
End Function

Public Function CrearTMPErr() As Boolean
'Crea una temporal donde insertara la clave primaria de las
'entradas erroneas al clasificar
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPErr = False
    
    SQL = "CREATE TEMPORARY TABLE tmpErrEnt ( "
    SQL = SQL & "numnotac int(7) unsigned NOT NULL default '0',"
    SQL = SQL & "codvarie int(6) unsigned )"
    
    conn.Execute SQL
     
    CrearTMPErr = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPErr = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpErrEnt;"
        conn.Execute SQL
    End If
End Function


Public Sub BorrarTMPErr()
On Error Resume Next
    conn.Execute " DROP TABLE IF EXISTS tmpErrEnt;"
    If Err.Number <> 0 Then Err.Clear
End Sub



'Private Function ActualizarTransporte(ByRef Rs As ADODB.Recordset, cadErr As String) As Boolean
'Dim Sql1 As String
'Dim Rs2 As ADODB.Recordset
'Dim KilosDestrio As Currency
'Dim Precio As Currency
'Dim Transporte As Currency
'Dim Kilos As Currency
'
'
'    On Error GoTo eActualizarTransporte
'
'    If Not Rs.EOF Then
'
'        Sql1 = "select imptrans from rportespobla, rpartida, rcampos, variedades "
'        Sql1 = Sql1 & " where rpartida.codparti = rcampos.codparti and "
'        Sql1 = Sql1 & " variedades.codprodu = rportespobla.codprodu and "
'        Sql1 = Sql1 & " rpartida.codpobla = rportespobla.codpobla and "
'        Sql1 = Sql1 & " variedades.codvarie = " & DBSet(Rs!CodVarie, "N") & " and "
'        Sql1 = Sql1 & " rcampos.codcampo = " & DBSet(Rs!codcampo, "N") & " and "
'        Sql1 = Sql1 & " rcampos.codvarie = variedades.codvarie "
'
'        Set Rs2 = New ADODB.Recordset
'        Rs2.Open Sql1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'        Precio = 0
'        If Not Rs2.EOF Then
'            Precio = DBLet(Rs2.Fields(0).Value, "N")
'        End If
'
'        Set Rs2 = Nothing
'
'        ' cogemos los kilos de la clasificacion que sean de destrio
'        Sql1 = "select kilosnet from rclasifica_clasif, rcalidad where numnotac = " & DBSet(Rs!numnotac, "N")
'        Sql1 = Sql1 & " and rclasifica_clasif.codvarie = rcalidad.codvarie "
'        Sql1 = Sql1 & " and rclasifica_clasif.codcalid = rcalidad.codcalid "
'        Sql1 = Sql1 & " and rcalidad.tipcalid = 1 "
'        KilosDestrio = DevuelveValor(Sql1)
'
'        ' los gastos de transporte se calculan sobre los kilosnetos - los de destrio
'        Kilos = DBLet(Rs!KilosNet, "N") - KilosDestrio
'        Transporte = Round2(Kilos * Precio, 2)
'
'        Sql1 = "update rclasifica set imptrans = " & DBSet(Transporte, "N")
'        Sql1 = Sql1 & " where numnotac = " & DBSet(Rs!numnotac, "N")
'        conn.Execute Sql1
'
'    End If
'
'eActualizarTransporte:
'    If Err.Number <> 0 Then
'        ActualizarTransporte = False
'        cadErr = Err.Description
'    Else
'        ActualizarTransporte = True
'    End If
'
'
'End Function



Private Function HayEntradasSinCRFID(cTabla As String, cWhere As String, cadena As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
Dim tabla As String
    
'[Monica]21/10/2011: antes estaba esto *****  QUITO EL LEFT JOIN : con el left join el sql falla ****
'    Tabla = "(" & cTabla & ") left join trzpalets on rentradas.numnotac = trzpalets.numnotac "
'
'    SQL = "select count(*) from " & Tabla
'    SQL = SQL & " Where trzpalets.crfid Is Null "
'
'    Cadena = "select rentradas.numnotac, rentradas.codvarie, variedades.nomvarie "
'    Cadena = Cadena & " from " & Tabla
'    Cadena = Cadena & " Where trzpalets.crfid Is Null  "
'
'    If cWhere <> "" Then
'        cWhere = QuitarCaracterACadena(cWhere, "{")
'        cWhere = QuitarCaracterACadena(cWhere, "}")
'        cWhere = QuitarCaracterACadena(cWhere, "_1")
'        SQL = SQL & " and " & cWhere
'        Cadena = Cadena & " and " & cWhere
'    End If
'
'    Cadena = Cadena & " and trzpalets.idpalet not in (select idpalet from trzlineas_cargas) "
'    SQL = SQL & " and trzpalets.idpalet not in (select idpalet from trzlineas_cargas) "

'[Monica]21/10/2011: modificado por

    SQL = "select count(*) from " & cTabla
    
    cadena = "select rentradas.numnotac, rentradas.codvarie, variedades.nomvarie "
    cadena = cadena & " from " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " where " & cWhere
        cadena = cadena & " where " & cWhere
    End If
    
    If cWhere <> "" Then
        SQL = SQL & " and "
        cadena = cadena & " and "
    End If
    
    SQL = SQL & "(rentradas.numnotac in (select numnotac from trzpalets "
    SQL = SQL & "  where  trzpalets.crfid Is Null and trzpalets.idpalet not in (select idpalet from trzlineas_cargas)) "
    SQL = SQL & "  or numnotac not in (select numnotac from trzpalets))"
    
    cadena = cadena & "(rentradas.numnotac in (select numnotac from trzpalets "
    cadena = cadena & "  where  trzpalets.crfid Is Null and trzpalets.idpalet not in (select idpalet from trzlineas_cargas)) "
    cadena = cadena & "  or numnotac not in (select numnotac from trzpalets))"
    
    If RegistrosAListar(SQL) = 0 Then
        HayEntradasSinCRFID = False
    Else
        HayEntradasSinCRFID = True
    End If
End Function

