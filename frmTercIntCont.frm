VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTercIntCont 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   5865
   Icon            =   "frmTercIntCont.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameRepxDia 
      Height          =   4455
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   5805
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   2970
         TabIndex        =   11
         Top             =   2700
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   7
         Left            =   4065
         TabIndex        =   12
         Top             =   2700
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Width           =   5100
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   32
            Left            =   3660
            TabIndex        =   7
            Top             =   435
            Width           =   1095
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Index           =   31
            Left            =   1200
            TabIndex        =   6
            Top             =   435
            Width           =   1095
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   5
            Left            =   3360
            Picture         =   "frmTercIntCont.frx":000C
            Top             =   435
            Width           =   240
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   4
            Left            =   915
            Picture         =   "frmTercIntCont.frx":0097
            Top             =   435
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Recepción:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   10
            Top             =   150
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   29
            Left            =   2835
            TabIndex        =   9
            Top             =   435
            Width           =   420
         End
         Begin VB.Label Label2 
            Caption         =   "Desde"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   8
            Top             =   435
            Width           =   465
         End
      End
      Begin VB.Frame FrameProgress 
         Height          =   1200
         Left            =   360
         TabIndex        =   1
         Top             =   3060
         Visible         =   0   'False
         Width           =   4935
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   400
            Left            =   120
            TabIndex        =   2
            Top             =   640
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   714
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label lblProgess 
            Caption         =   "Iniciando el proceso ..."
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   4
            Top             =   375
            Width           =   4575
         End
         Begin VB.Label lblProgess 
            Caption         =   "Comprobaciones:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   135
            Width           =   4455
         End
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Contabilizar Facturas Terceros"
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
         Index           =   0
         Left            =   480
         TabIndex        =   13
         Top             =   580
         Width           =   5055
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   10680
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTercIntCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OpcionListado As Integer

'224 .- Pedir datos para contabilizar facturas socios terceros
    
    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmCta As frmCtasConta
Attribute frmCta.VB_VarHelpID = -1

Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1

'---- Variables para el INFORME ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para el frmImprimir
Private cadNomRPT As String 'Nombre del informe a Imprimir
Private conSubRPT As Boolean 'Si el informe tiene subreports
'-----------------------------------


Dim indCodigo As Integer 'indice para txtCodigo

Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean
Dim indFrame As Single

Dim cContaFra As cContabilizarFacturas


Private Sub KEYpress(KeyAscii As Integer)
    Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click()
'Reparaciones por Dia
Dim devuelve As String
Dim param As String
Dim TotalMante As Integer
Dim Rs As ADODB.Recordset
Dim fecha1 As String, fecha2 As String
Dim NomTabla As String

    InicializarVbles

    '===================================================
    '============ PARAMETROS ===========================
    'Añadir el nombre de la Empresa como parametro
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1

    param = ""

    Codigo = "{rcafter.fecrecep}"
    NomTabla = "rcafter"

    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Desde y Hasta FECHA
    
    'comprobar que se han rellenado los dos campos de fecha
    'sino rellenar con fechaini o fechafin del ejercicio
    'que guardamos en vbles Orden1,Orden2
    If txtcodigo(31).Text = "" Then
       txtcodigo(31).Text = Orden1 'fechaini del ejercicio de la conta
    End If

    If txtcodigo(32).Text = "" Then
       txtcodigo(32).Text = Orden2 'fecha fin del ejercicio de la conta
    End If

     'Comprobar que el intervalo de fechas D/H esta dentro del ejercicio de la
     'contabilidad par ello mirar en la BD de la Conta los parámetros
    If Not ComprobarFechasConta(31) Then Exit Sub
    If Not ComprobarFechasConta(32) Then Exit Sub
    
    '++monica: comprobar si es factura de cliente que se ponen los datos de tesoreria
'        If txtCodigo(33).Text = "" Then
'            MsgBox "Debe introducir los datos de tesoreria.", vbExclamation
'            PonerFoco txtCodigo(33)
'            Exit Sub
'        End If
        

    devuelve = CadenaDesdeHasta(txtcodigo(31).Text, txtcodigo(32).Text, Codigo, "F")
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    'Parametro D/H Fecha
    If devuelve <> "" And param <> "" Then
        CadParam = CadParam & AnyadirParametroDH(param, 31, 32) & """|"
        numParam = numParam + 1
    End If


    '===================================================
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    cadSelect = CadenaDesdeHastaBD(txtcodigo(31).Text, txtcodigo(32).Text, Codigo, "F")
    cadSelect = cadSelect & " AND " & NomTabla & ".intconta=0 "
    
    If Not HayRegParaInforme(NomTabla, cadSelect) Then Exit Sub

    '------------------------------------------------------------------------------
    '  LOG de acciones.                      5: Facturas compras
    Set LOG = New cLOG
    LOG.Insertar 5, vUsu, "Contabilizar facturas terceros: " & vbCrLf & NomTabla & vbCrLf & cadSelect
    Set LOG = Nothing
    '-----------------------------------------------------------------------------

    ContabilizarFacturas NomTabla, cadSelect
    TerminaBloquear
     'Eliminar la tabla TMP
    BorrarTMPFacturas
    'Desbloqueamos ya no estamos contabilizando facturas
    DesBloqueoManual ("TERCON") 'TERceros CONtabilizar
    
    Me.FrameProgress.visible = False
    Me.FrameRepxDia.Height = 3500
    Me.Height = Me.FrameRepxDia.Height + 350
    
    Unload Me
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub


Private Sub Form_Activate()
Dim IndiceFoco As Integer

    If PrimeraVez Then
        PrimeraVez = False
        IndiceFoco = 31
        If IndiceFoco >= 0 Then PonerFoco txtcodigo(IndiceFoco)
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer


    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me

    'Ocultar todos los Frames de Formulario
    Me.FrameRepxDia.visible = False
    
    CommitConexion
    
    
    cadTitulo = ""
    cadNomRPT = ""
    
    PonerFrameRepxDiaVisible True, H, W
    indFrame = 7
        
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub

Private Sub Form_Unload(Cancel As Integer)
    NumCod = ""
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

    Screen.MousePointer = vbHourglass

    Set frmF = New frmCal
    
    esq = imgFecha(Index).Left
    dalt = imgFecha(Index).Top
    
    Set obj = imgFecha(Index).Container

    While imgFecha(Index).Parent.Name <> obj.Name
        esq = esq + obj.Left
        dalt = dalt + obj.Top
        Set obj = obj.Container
    Wend
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmF.Left = esq + imgFecha(Index).Parent.Left + 30
    frmF.Top = dalt + imgFecha(Index).Parent.Top + imgFecha(Index).Height + menu - 40


   imgFecha(0).Tag = Index
'   Set frmF = New frmCal
   frmF.NovaData = Now
   
   Select Case Index
        Case 4
            indCodigo = 31
        Case 5
            indCodigo = 32
   End Select
   
   
   PonerFormatoFecha txtcodigo(indCodigo)
   If txtcodigo(indCodigo).Text <> "" Then frmF.NovaData = CDate(txtcodigo(indCodigo).Text)
   
   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtcodigo(indCodigo)
End Sub


Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtcodigo(Index), 3
End Sub


Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Tabla As String
Dim codcampo As String, nomCampo As String
Dim TipCampo As String, Formato As String
Dim Titulo As String
Dim EsNomCod As Boolean 'Si es campo Cod-Descripcion llama a PonerNombreDeCod


    'Quitar espacios en blanco por los lados
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    EsNomCod = False
    If Index = 1 Or Index = 2 Then
    'el mismo frame ( y por tanto los mismos campos) se utilizan para distintos
    'informes. Según de donde llamemos código de una tabla u otra
        
    ElseIf Index = 3 Or Index = 4 Then
         '7: Informe Traspaso Almacenes
         '8: Informe Movimientos Almacen
         txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "0000000")
    Else
        Select Case Index
        Case 31, 32 'FECHA Desde Hasta
            If txtcodigo(Index).Text <> "" Then
                    PonerFormatoFecha txtcodigo(Index)
                    If OpcionListado = 223 And txtcodigo(Index).Text <> "" Then
                        'Contabilizar facturas
                        If Not ComprobarFechasConta(Index) Then
                            PonerFoco txtcodigo(Index)
                        End If '++
                    End If
            End If
        End Select
    End If
    
'    If EsNomCod Then
'        If TipCampo = "N" Then
'            If PonerFormatoEntero(txtCodigo(Index)) Then
'
'
'                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), tabla, nomcampo, codCampo)
''                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), Tabla, NomCampo, codCampo, Titulo, TipCampo)
'
'                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, Formato)
'            Else
'                txtNombre(Index).Text = ""
'            End If
'        Else
'            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), tabla, nomcampo, codCampo)
''            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), Tabla, NomCampo, codCampo, Titulo, TipCampo)
'        End If
'    End If
    
   
End Sub







Private Sub PonerFrameRepxDiaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para los listados de las Reparaciones x dia, de tabla: scarep

    H = 3500
    W = 6000
    PonerFrameVisible Me.FrameRepxDia, visible, H, W

    If visible = True Then
        Me.Caption = "AriagroRec"
        
        Me.FrameProgress.visible = False
                
        Me.lblTitulo(0).Caption = "Contabilizar Facturas Terceros"
        Me.Label2(2).Caption = "Fecha Recepción:"
        
        Frame2.Top = 1680
    End If
End Sub


Private Function AnyadirParametroDH(Cad As String, indD As Byte, indH As Byte) As String
On Error Resume Next
    
     If txtcodigo(indD).Text <> "" Then
        Cad = Cad & "desde " & txtcodigo(indD).Text
'        If txtNombre(indD).Text <> "" Then Cad = Cad & " - " & txtNombre(indD).Text
    End If
    If txtcodigo(indH).Text <> "" Then
        Cad = Cad & "  hasta " & txtcodigo(indH).Text
'        If txtNombre(indH).Text <> "" Then Cad = Cad & " - " & txtNombre(indH).Text
    End If
    
    AnyadirParametroDH = Cad
    If Err.Number <> 0 Then Err.Clear
End Function


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    CadParam = ""
    numParam = 0
'    cadTitulo = ""
'    cadNomRPT = ""
    conSubRPT = False
End Sub


Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim devuelve As String
Dim Cad As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(txtcodigo(indD).Text, txtcodigo(indH).Text, campo, Tipo)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    
    If Tipo <> "F" Then
        'Fecha para Crystal Report
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        Cad = CadenaDesdeHastaBD(txtcodigo(indD).Text, txtcodigo(indH).Text, campo, Tipo)
        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Function
    End If
    
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            CadParam = CadParam & AnyadirParametroDH(param, indD, indH) & """|"
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
        .Opcion = OpcionListado
        .Titulo = cadTitulo
        .NombreRPT = cadNomRPT
        .ConSubInforme = conSubRPT
        .Show vbModal
    End With
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
    
    ' tenemos que abrir primero la conexion
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If vSeccion.AbrirConta Then
        
            If txtcodigo(ind).Text <> "" Then
                FechaIni = "Select fechaini,fechafin From parametros"
                Set Rs = New ADODB.Recordset
                Rs.Open FechaIni, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
            
                If Not Rs.EOF Then
                    FechaIni = DBLet(Rs!FechaIni, "F")
                    FechaFin = DateAdd("yyyy", 1, DBLet(Rs!FechaFin, "F"))
                    'nos guardamos los valores
                    Orden1 = FechaIni
                    Orden2 = FechaFin
                
                    If Not EntreFechas(FechaIni, txtcodigo(ind).Text, FechaFin) Then
                         Cad = "El período de contabilización debe estar dentro del ejercicio:" & vbCrLf & vbCrLf
                         Cad = Cad & "    Desde: " & FechaIni & vbCrLf
                         Cad = Cad & "    Hasta: " & FechaFin
                         MsgBox Cad, vbExclamation
                         txtcodigo(ind).Text = ""
                    Else
                        ComprobarFechasConta = True
                    End If
                End If
                Rs.Close
                Set Rs = Nothing
            Else
                ComprobarFechasConta = True
            End If
                    
            vSeccion.CerrarConta
        End If
    End If
    
    Set vSeccion = Nothing
    
EComprobar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Fechas", Err.Description
End Function



Private Sub ContabilizarFacturas(cadTabla As String, cadWHERE As String)
'Contabiliza Facturas de Clientes o de Proveedores
Dim Sql As String
Dim b As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String

Dim vSeccion As CSeccion

    Sql = "TERCON" 'contabilizar facturas de terceros
    
    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se pueden Contabilizar Facturas de Terceros. Hay otro usuario contabilizando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


     'comprobar que se han rellenado los dos campos de fecha
     'sino rellenar con fechaini o fechafin del ejercicio
     'que guardamos en vbles Orden1,Orden2
     If txtcodigo(31).Text = "" Then
        txtcodigo(31).Text = Orden1 'fechaini del ejercicio de la conta
     End If

     If txtcodigo(32).Text = "" Then
        txtcodigo(32).Text = Orden2 'fecha fin del ejercicio de la conta
     End If


     'Comprobar que el intervalo de fechas D/H esta dentro del ejercicio de la
     'contabilidad par ello mirar en la BD de la Conta los parámetros
     If Not ComprobarFechasConta(32) Then Exit Sub



    'comprobar si existen en Ariges facturas anteriores al periodo solicitado
    'sin contabilizar
    If Me.txtcodigo(31).Text <> "" Then 'anteriores a fechadesde
        Sql = "SELECT COUNT(*) FROM " & cadTabla
        Sql = Sql & " WHERE fecrecep <"
        Sql = Sql & DBSet(txtcodigo(31), "F") & " AND intconta=0 "
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
    Sql = Sql & ".codsocio=tmpFactu.codsocio AND "
    
    Sql = Sql & cadTabla & ".numfactu=tmpFactu.numfactu AND " & cadTabla & ".fecfactu=tmpFactu.fecfactu "
    If Not BloqueaRegistro(Sql, cadWHERE) Then
        MsgBox "No se pueden Contabilizar Facturas. Hay registros bloqueados.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


    '---- Preparamos la pantalla de Contabilizar
    'Visualizar la barra de Progreso
    Me.FrameRepxDia.Height = 5100
    Me.Height = Me.FrameRepxDia.Height
    Me.FrameProgress.visible = True
    Me.FrameProgress.Top = 3350

    Me.lblProgess(0).Caption = "Comprobaciones: "
    CargarProgres Me.ProgressBar1, 100


    'comprobar que todas las LETRAS SERIE existen en la contabilidad y en Ariges
    '--------------------------------------------------------------------------
    IncrementarProgres Me.ProgressBar1, 10
'    If cadtabla = "facturas" Then
'        Me.lblProgess(1).Caption = "Comprobando letras de serie ..."
'        b = ComprobarLetraSerie(cadtabla)
'    End If
    IncrementarProgres Me.ProgressBar1, 10
    Me.Refresh
    If Not b Then Exit Sub


    'comprobar que no haya Nº FACTURAS en la contabilidad para esa fecha
    'que ya existan
    '-----------------------------------------------------------------------
'    If cadtabla = "facturas" Then
'        Me.lblProgess(1).Caption = "Comprobando Nº Facturas en contabilidad ..."
'        SQL = "anofaccl>=" & Year(txtCodigo(31).Text) & " AND anofaccl<= " & Year(txtCodigo(32).Text)
'        b = ComprobarNumFacturas_new(cadtabla, SQL)
'    End If
    IncrementarProgres Me.ProgressBar1, 20
    Me.Refresh
    If Not b Then Exit Sub


    'comprobar que todas las CUENTAS de los distintos socios terceros que vamos a
    'contabilizar existen en la Conta: rsocios_seccion.codmacta IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    Me.lblProgess(1).Caption = "Comprobando Cuentas Contables en contabilidad ..."
    
    ' tenemos que abrir primero la conexion
    Set vSeccion = New CSeccion
    b = False
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If vSeccion.AbrirConta Then
            b = ComprobarCtaContable_new(cadTabla, 1)
        End If
        vSeccion.CerrarConta
    End If
    
    Set vSeccion = Nothing
    IncrementarProgres Me.ProgressBar1, 20
    Me.Refresh
    If Not b Then Exit Sub


    'comprobar que todas las CUENTAS de compras terceros de las variedades que vamos a
    'contabilizar existen en la Conta: variedades.ctacomtercero IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    ' tenemos que abrir primero la conexion
    Set vSeccion = New CSeccion
    b = False
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If vSeccion.AbrirConta Then
            'comprobar que todas las CUENTAS de compras terceros de las variedades que vamos a
            'contabilizar existen en la Conta: (variedades.ctacomtercero) IN (conta.cuentas.codmacta)
            b = ComprobarCtaContable_new(cadTabla, 8)
        End If
        vSeccion.CerrarConta
    End If
    
    Set vSeccion = Nothing
    IncrementarProgres Me.ProgressBar1, 10
    Me.Refresh
    If Not b Then Exit Sub
    
    
    '[Monica]23/09/2013
    
    'comprobar que todas las CUENTAS de concepto de cargo (MONTIFRUT) que vamos a
    'contabilizar existen en la Conta: fvarconcep.codmacpr IN (conta.cuentas.codmacta)
    '-----------------------------------------------------------------------------
    ' tenemos que abrir primero la conexion
    Set vSeccion = New CSeccion
    b = False
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If vSeccion.AbrirConta Then
            'comprobar que todas las CUENTAS de compras terceros de las variedades que vamos a
            'contabilizar existen en la Conta: (fvarconcep.codmacpr) IN (conta.cuentas.codmacta)
            b = ComprobarCtaContable_new(cadTabla, 13)
        End If
        vSeccion.CerrarConta
    End If
    
    Set vSeccion = Nothing
    IncrementarProgres Me.ProgressBar1, 10
    Me.Refresh
    If Not b Then Exit Sub
    


    ' comprobamos que existe las cuentas vparamaplic.ctatrareten
    ' tenemos que abrir primero la conexion
    Set vSeccion = New CSeccion
    b = False
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If vSeccion.AbrirConta Then
            b = ComprobarCtaContable_new(cadTabla, 4)
        End If
        vSeccion.CerrarConta
    End If
    
    Set vSeccion = Nothing
    IncrementarProgres Me.ProgressBar1, 20
    Me.Refresh
    If Not b Then Exit Sub





    'comprobar que todos las TIPO IVA de las distintas fecturas que vamos a
    'contabilizar existen en la Conta: scafac.codigiv1,codigiv2,codigiv3 IN (conta.tiposiva.codigiva)
    '--------------------------------------------------------------------------
    Me.lblProgess(1).Caption = "Comprobando Tipos de IVA en contabilidad ..."
    ' tenemos que abrir primero la conexion
    Set vSeccion = New CSeccion
    b = False
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If vSeccion.AbrirConta Then
            b = ComprobarTiposIVA(cadTabla)
        End If
        vSeccion.CerrarConta
    End If
    
    Set vSeccion = Nothing
    IncrementarProgres Me.ProgressBar1, 10
    Me.Refresh
    If Not b Then Exit Sub


    'comprobar si hay contabilidad ANALITICA: conta.parametros.autocoste=1
    'y verificar que las cuentas de sfamia.ctaventa empiezan por el digito
    'de conta.parametros.grupogto o conta.parametros.grupovta
    'obtener el centro de coste del usuario para insertarlo en linfact
    
    ' tenemos que abrir primero la conexion
    Set vSeccion = New CSeccion
    b = False
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If vSeccion.AbrirConta Then
            If vEmpresa.TieneAnalitica Then  'hay contab. analitica
               Me.lblProgess(1).Caption = "Comprobando Contabilidad Analítica ..."
                   
               b = ComprobarCtaContable_new(cadTabla, 7)
               
               If Not b Then
                    vSeccion.CerrarConta
                    Set vSeccion = Nothing
                    Exit Sub
               End If
        
               '(si tiene analítica requiere un centro de coste para insertar en conta.linfact)
               CCoste = ""
               b = ComprobarCCoste_new(CCoste, cadTabla)
               If Not b Then Exit Sub
    
               CCoste = ""
            End If
       End If
       
        vSeccion.CerrarConta
    End If
    
    Set vSeccion = Nothing
       
    IncrementarProgres Me.ProgressBar1, 10
    Me.Refresh


    '===========================================================================
    'CONTABILIZAR FACTURAS
    '===========================================================================
    Me.lblProgess(0).Caption = "Contabilizar Facturas: "
    CargarProgres Me.ProgressBar1, 10
    Me.lblProgess(1).Caption = "Insertando Facturas en Contabilidad..."


    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    LOG.Insertar 3, vUsu, "Contabilizar facturas: " & vbCrLf & cadTabla & vbCrLf & cadWHERE
    Set LOG = Nothing
    '-----------------------------------------------------------------------------


    '---- Crear tabla TEMP para los posible errores de facturas
    tmpErrores = CrearTMPErrFact(cadTabla)

    '---- Pasar las Facturas a la Contabilidad
    Set vSeccion = New CSeccion
    b = False
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If vSeccion.AbrirConta Then
    
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
            If cadTabla = "scafpc" Or cadTabla = "rcafter" Then
                If DevuelveValor("Select count(*) from tmpinformes where codusu = " & vUsu.Codigo) > 0 Then
                    InicializarVbles
                    CadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
                    numParam = numParam + 1
                    
                    CadParam = "|pDHFecha=""" & vUsu.Nombre & "   Hora: " & Format(Now, "hh:mm") & """|"
                    numParam = numParam + 1
                    cadFormula = "({tmpinformes.codusu} =" & vUsu.Codigo & ")"
                    conSubRPT = False
                    If cadTabla = "scafpc" Then
                        cadTitulo = "Listado contabilizacion FRAPRO"
                        cadNomRPT = "rContabPRO.rpt"
                    Else
                        cadTitulo = "Listado contabilizacion FRATER"
                        cadNomRPT = "rContabTER.rpt"
                    End If
                    
                    LlamarImprimir
                End If
            End If


            '---- Eliminar tabla TEMP de Errores
            BorrarTMPErrFact

        End If
        vSeccion.CerrarConta
    End If
    
    Set vSeccion = Nothing

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
    Codigo1 = "codsocio"
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
        CargarProgres Me.ProgressBar1, numfactu

        'seleccinar todas las facturas que hemos insertado en la temporal (las que vamos a contabilizar)
        Sql = "SELECT * "
        Sql = Sql & " FROM tmpFactu "

        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenStatic, adLockPessimistic, adCmdText
        i = 1

        b = True
        'pasar a contabilidad cada una de las facturas seleccionadas
        While Not Rs.EOF
            Sql = cadTabla & "." & Codigo1 & "=" & DBSet(Rs.Fields(0), "N") & " and numfactu=" & DBSet(Rs!numfactu, "T")
            Sql = Sql & " and fecfactu=" & DBSet(Rs!fecfactu, "F")
            If PasarFacturaTerc(Sql, CCoste, Orden2, cContaFra) = False And b Then b = False

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

            IncrementarProgres Me.ProgressBar1, 1
            Me.lblProgess(1).Caption = "Insertando Facturas en Contabilidad...   (" & i & " de " & numfactu & ")"
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


