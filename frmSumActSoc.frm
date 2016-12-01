VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSumActSoc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alta / Modificación Masiva de Socios en Suministros"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6360
   Icon            =   "frmSumActSoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   6360
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
      Height          =   5670
      Left            =   0
      TabIndex        =   5
      Top             =   30
      Width           =   6330
      Begin VB.Frame Frame3 
         Height          =   675
         Left            =   120
         TabIndex        =   18
         Top             =   3000
         Width           =   6045
         Begin VB.OptionButton Option1 
            Caption         =   "Sólo Insertar"
            Height          =   195
            Index           =   0
            Left            =   1290
            TabIndex        =   20
            Top             =   270
            Width           =   1755
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Insertar y Modificar"
            Height          =   285
            Index           =   1
            Left            =   3210
            TabIndex        =   19
            Top             =   240
            Width           =   1815
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
         Height          =   1530
         Left            =   90
         TabIndex        =   6
         Top             =   225
         Width           =   6060
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   2715
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   1020
            Width           =   2685
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   1500
            MaxLength       =   6
            TabIndex        =   1
            Tag             =   "Socio|N|S|||rsocios|codsocio|000000||"
            Top             =   1020
            Width           =   1125
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   1500
            MaxLength       =   6
            TabIndex        =   0
            Tag             =   "Socio|N|S|||rsocios|codsocio|000000||"
            Top             =   660
            Width           =   1125
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   2715
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   660
            Width           =   2685
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   2
            Left            =   1230
            ToolTipText     =   "Buscar Socio"
            Top             =   1020
            Width           =   240
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   1
            Left            =   1230
            ToolTipText     =   "Buscar Socio"
            Top             =   660
            Width           =   240
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   14
            Left            =   600
            TabIndex        =   15
            Top             =   1050
            Width           =   420
         End
         Begin VB.Label Label4 
            Caption         =   "Desde"
            Height          =   195
            Index           =   15
            Left            =   615
            TabIndex        =   14
            Top             =   645
            Width           =   465
         End
         Begin VB.Label Label4 
            Caption         =   "Socio"
            ForeColor       =   &H00972E0B&
            Height          =   255
            Index           =   2
            Left            =   225
            TabIndex        =   13
            Top             =   405
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos por defecto"
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
         Height          =   1065
         Left            =   90
         TabIndex        =   7
         Top             =   1830
         Width           =   6075
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   1620
            MaxLength       =   10
            TabIndex        =   2
            Tag             =   "Cta.Contable|T|S|||sparam|ctaconta|||"
            Top             =   450
            Width           =   1125
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   2835
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   450
            Width           =   2685
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   1350
            ToolTipText     =   "Buscar Forma Pago"
            Top             =   450
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Forma de Pago"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   12
            Top             =   495
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5100
         TabIndex        =   4
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3915
         TabIndex        =   3
         Top             =   4800
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   3930
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProgres 
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   4260
         Width           =   5940
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   4500
         Width           =   5925
      End
   End
End
Attribute VB_Name = "frmSumActSoc"
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
Private WithEvents frmFPa As frmForpaSumi 'formas de pago de suministros
Attribute frmFPa.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'ssocios
Attribute frmSoc.VB_VarHelpID = -1

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

Dim Nregs As Long

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me

End Sub

Private Sub cmdAceptar_Click()
Dim Sql As String
Dim i As Byte
Dim cadWHERE As String
Dim cDesde As String
Dim cHasta As String
Dim cTabla As String
Dim vSQL As String


    If Not DatosOk Then Exit Sub

    'D/H Socio
    cDesde = Trim(txtcodigo(1).Text)
    cHasta = Trim(txtcodigo(2).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rsocios.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHSocio= """) Then Exit Sub
    End If

    cTabla = "rsocios INNER JOIN rsocios_seccion ON rsocios.codsocio = rsocios_seccion.codsocio "
    ' que sean de la seccion de suministros
    cTabla = cTabla & " and rsocios_seccion.codsecci = " & vParamAplic.SeccionSumi
    ' que esten dados de alta en dicha seccion
    cTabla = cTabla & " and (rsocios_seccion.fecbaja is null or rsocios_seccion.fecbaja='0000-00-00')"
    
    
    cadSelect = Replace(Replace(cadSelect, "}", ""), "{", "")
    
    vSQL = "select count(*) from " & cTabla & " where " & cadSelect
    
    Nregs = TotalRegistros(vSQL)
    If Nregs = 0 Then
        MsgBox "No hay datos entre esos límites.", vbExclamation
        Exit Sub
    Else
        ActualizarSocios cTabla, cadSelect
    End If


eError:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "No se ha realizado el proceso de actualización de datos de socio en Suministros. Llame a soporte."
    End If

    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    
    'Desbloqueamos ya no estamos actualizando datos socios
    DesBloqueoManual ("ACTSOC") 'ACTualizar datos SOCios en ariges
    
    cmdCancel_Click
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtcodigo(5)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection
Dim i As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon


    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
    For i = 1 To 3
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    ConexionAriges
    
    '###Descomentar
'    CommitConexion

    FrameCobrosVisible True, H, W
    Pb1.visible = False

    Me.Option1(0).Value = True
    Me.Option1(1).Value = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
'   Me.Width = w + 70
'   Me.Height = h + 350


End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not vSeccion Is Nothing Then
        vSeccion.CerrarAriges
        Set vSeccion = Nothing
    End If
End Sub


Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Cuentas contables
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFPa_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de formas de pago de contabilidad
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtcodigo(indCodigo).Text = Format(txtcodigo(indCodigo).Text, "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 1, 2 ' socio
            AbrirFrmSocios (Index)
        Case 3 'forma de pago de ariges
            AbrirFrmForpaSumi (Index)
    End Select
    PonerFoco txtcodigo(indCodigo)
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
    ConseguirFoco txtcodigo(Index), 3
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
            Case 1: KEYBusqueda KeyAscii, 1 'socio desde
            Case 2: KEYBusqueda KeyAscii, 2 'socio hasta
            Case 3: KEYBusqueda KeyAscii, 3 'forma de pago suministros
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

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 1, 2 ' socios
            If txtcodigo(Index).Text <> "" Then txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rsocios", "nomsocio", "codsocio", "N")

        Case 3 ' FORMAS DE PAGO DE suministros
            If txtcodigo(Index).Text <> "" Then txtNombre(Index).Text = DevuelveDesdeBDNew(cAriges, "sforpa", "nomforpa", "codforpa", txtcodigo(Index).Text, "N")
            If txtNombre(Index).Text = "" Then
                MsgBox "Forma de Pago  no existe en suministros. Reintroduzca.", vbExclamation
            End If
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
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

Private Sub AbrirFrmSocios(indice As Integer)
    indCodigo = indice
    Set frmSoc = New frmManSocios
    frmSoc.DatosADevolverBusqueda = "0|1|"
    frmSoc.Show vbModal
    Set frmSoc = Nothing
End Sub

Private Sub AbrirFrmForpaSumi(indice As Integer)
    indCodigo = indice
    Set frmFPa = New frmForpaSumi
    frmFPa.DatosADevolverBusqueda = "0|1|"
    frmFPa.CodigoActual = txtcodigo(indCodigo)
'    frmFpa.Conexion = cContaFacSoc
    frmFPa.Show vbModal
    Set frmFPa = Nothing
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Orden1 As String
Dim Orden2 As String
Dim FFin As Date
Dim cta As String

   b = True

   If txtcodigo(3).Text = "" Then
        MsgBox "Introduzca la Forma de pago del Socio.", vbExclamation
        b = False
        PonerFoco txtcodigo(3)
   End If
   
   If b Then
        b = ComprobarReferenciales
   End If
   
   DatosOk = b

End Function

Private Sub ActualizarSocios(cadTabla As String, cadWHERE As String)
'Contabiliza Facturas de Clientes o de Proveedores
Dim Sql As String
Dim Sql2 As String
Dim vSQL As String
Dim Rs As ADODB.Recordset
Dim rsAriges As ADODB.Recordset
Dim Rs4 As ADODB.Recordset
Dim b As Boolean
Dim tmpErrores As Boolean 'Indica si se creo correctamente la tabla de errores
Dim CCoste As String
Dim CtaClien As String
Dim CodIva As String
Dim Sql4 As String

    On Error GoTo eActualizarSocios


    Sql = "ACTSOC" 'ACTualizar datos de SOCios de suminitros
    
    'Bloquear para que nadie mas pueda contabilizar
    DesBloqueoManual (Sql)
    If Not BloqueoManual(Sql, "1") Then
        MsgBox "No se pueden Actualizar Datos de Socios de Suministros. Hay otro usuario actualizando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If


    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    LOG.Insertar 3, vUsu, "Actualizar Datos Socios Ariges: " & vbCrLf & cadTabla & vbCrLf & cadWHERE
    Set LOG = Nothing
    '-----------------------------------------------------------------------------

    Sql = "select * from " & cadTabla & " where " & cadWHERE

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Visualizar la barra de Progreso
    Me.Pb1.visible = True
    Me.lblProgres(0).Caption = "Procesando Socios: "
    ' si insertamos o modificamos el numero de registros nos coincide
    Pb1.Max = Nregs
    Pb1.Value = 0

    While Not Rs.EOF
        Me.Pb1.Value = Me.Pb1.Value + 1
        Me.lblProgres(1).Caption = "Socio: " & DBLet(Rs!Codsocio, "N") & "-" & UCase(DBLet(Rs!nomsocio, "T"))
        Me.Refresh
    
        If ExisteCliente(CStr(DBLet(Rs!Codsocio, "N"))) Then
            ' si solo insertar
            If Option1(0).Value Then
                Me.lblProgres(0).Caption = "Anteriormente insertado"
                Me.Refresh
                ' no hacemos nada
            End If
            
            If Option1(1).Value Then
                Me.lblProgres(0).Caption = "Modificando"
                Me.Refresh
                
                CtaClien = ""
                Sql4 = "select codmaccli from rsocios_seccion where codsocio = " & DBSet(Rs!Codsocio, "N")
                Sql4 = Sql4 & " and codsecci = " & vParamAplic.SeccionSumi
                
                Set Rs4 = New ADODB.Recordset
                Rs4.Open Sql4, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not Rs4.EOF Then
                    CtaClien = DBLet(Rs4!codmaccli, "T")
                End If
            
                ' si insertar y modificar
                Sql2 = "update " & vParamAplic.BDAriges & ".sclien "
                Sql2 = Sql2 & " set nomclien = " & DBSet(Rs!nomsocio, "T")
                Sql2 = Sql2 & ", domclien = " & DBSet(Rs!dirsocio, "T")
                Sql2 = Sql2 & ", codpobla = " & DBSet(Rs!codpostal, "T")
                Sql2 = Sql2 & ", pobclien = " & DBSet(Rs!pobsocio, "T")
                Sql2 = Sql2 & ", proclien = " & DBSet(Rs!prosocio, "T")
                Sql2 = Sql2 & ", nifclien = " & DBSet(Rs!nifSocio, "T")
                Sql2 = Sql2 & ", codforpa = " & DBSet(txtcodigo(3).Text, "N")
                Sql2 = Sql2 & ", codbanco = " & DBSet(Rs!CodBanco, "N")
                Sql2 = Sql2 & ", codsucur = " & DBSet(Rs!CodSucur, "N")
                Sql2 = Sql2 & ", digcontr = " & DBSet(Rs!digcontr, "T")
                Sql2 = Sql2 & ", cuentaba = " & DBSet(Rs!CuentaBa, "T")
                Sql2 = Sql2 & ", codmacta = " & DBSet(CtaClien, "T")
                Sql2 = Sql2 & ", telclie1 = " & DBSet(Rs!telsoci1, "T")
                Sql2 = Sql2 & ", maiclie1 = " & DBSet(Rs!maisocio, "T")
                Sql2 = Sql2 & " where codclien = " & DBSet(Rs!Codsocio, "N")
                
                ConnAriges.Execute Sql2
                
            End If
        Else
            ' obtenemos la cuenta contable de la seccion de Suministros
            CtaClien = ""
            Sql4 = "select codmaccli from rsocios_seccion where codsocio = " & DBSet(Rs!Codsocio, "N")
            Sql4 = Sql4 & " and codsecci = " & vParamAplic.SeccionSumi
            
            Set Rs4 = New ADODB.Recordset
            Rs4.Open Sql4, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not Rs4.EOF Then
                CtaClien = DBLet(Rs4!codmaccli, "T")
            End If
            Set Rs4 = Nothing
            
            ' obtenemos los datos por defecto de la tabla spara1 de ariges1
            vSQL = "select defactividad, defenvio, defzona, defruta, defstituacion, deftarifa, defagente "
            vSQL = vSQL & " from spara1 "
            
            Set rsAriges = New ADODB.Recordset
            rsAriges.Open vSQL, ConnAriges, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If Not rsAriges.EOF Then
                Me.lblProgres(0).Caption = "Insertando"
                Me.Refresh
            
                ' tanto si solo insertar o insertar y modificar
                Sql2 = "insert into " & vParamAplic.BDAriges & ".sclien ("
                Sql2 = Sql2 & "codclien, nomclien, nomcomer, domclien, codpobla, pobclien, proclien, nifclien, "
                Sql2 = Sql2 & "fechaalt, codactiv, codenvio, codzonas, codrutas, codagent, codforpa, codbanco, codsucur, digcontr, cuentaba, codmacta,"
                Sql2 = Sql2 & "telclie1, maiclie1, clivario, tipoiva,  tipofact, albarcon, periodof, codtarif, "
                Sql2 = Sql2 & "codsitua) values "
                Sql2 = Sql2 & "( "
                Sql2 = Sql2 & DBSet(Rs!Codsocio, "N") & "," 'codigo
                Sql2 = Sql2 & DBSet(Rs!nomsocio, "T") & "," 'nomclien
                Sql2 = Sql2 & DBSet(Rs!nomsocio, "T") & "," 'nombre comercial
                Sql2 = Sql2 & DBSet(Rs!dirsocio, "T") & "," 'direccion
                Sql2 = Sql2 & DBSet(Rs!codpostal, "T") & "," 'codigo postal
                Sql2 = Sql2 & DBSet(Rs!pobsocio, "T") & "," 'poblacion
                Sql2 = Sql2 & DBSet(DBLet(Rs!prosocio, "T"), "T", "N") & "," 'provincia
                Sql2 = Sql2 & DBSet(Rs!nifSocio, "T") & "," 'nifsocio
                Sql2 = Sql2 & DBSet(Now, "F") & ","         'fecha de alta
                Sql2 = Sql2 & DBSet(rsAriges!defactividad, "N") & "," 'codigo de actividad
                Sql2 = Sql2 & DBSet(rsAriges!defenvio, "N") & "," 'codenvio
                Sql2 = Sql2 & DBSet(rsAriges!defzona, "N") & "," 'codzonas
                Sql2 = Sql2 & DBSet(rsAriges!defruta, "N") & "," 'codrutas
                Sql2 = Sql2 & DBSet(rsAriges!defagente, "N") & "," 'codigo de agente
                Sql2 = Sql2 & DBSet(txtcodigo(3).Text, "N") & "," 'codigo de forma de pago
                Sql2 = Sql2 & DBSet(Rs!CodBanco, "N") & "," 'codigo de banco
                Sql2 = Sql2 & DBSet(Rs!CodSucur, "N") & "," 'codigo de sucursal
                Sql2 = Sql2 & DBSet(Rs!digcontr, "T") & "," 'digcontr
                Sql2 = Sql2 & DBSet(Rs!CuentaBa, "T") & "," 'cuenta banco
                Sql2 = Sql2 & DBSet(CtaClien, "T") & "," 'codmacta de la seccion de suministros
                Sql2 = Sql2 & DBSet(Rs!telsoci1, "T") & "," ' telefono
                Sql2 = Sql2 & DBSet(Rs!maisocio, "T") & "," 'mail
                Sql2 = Sql2 & "0," 'clivario
                Sql2 = Sql2 & "0," 'tipoiva
                Sql2 = Sql2 & "0," 'tipo de factura
                Sql2 = Sql2 & "0," 'albarcon
                Sql2 = Sql2 & "0," 'periodof
                Sql2 = Sql2 & DBSet(rsAriges!deftarifa, "N") & "," 'tarifa
                Sql2 = Sql2 & DBSet(rsAriges!defstituacion, "N") & ")" 'situacion
                
                ConnAriges.Execute Sql2
            End If
        End If
        
        Rs.MoveNext
    Wend

    Set Rs = Nothing


    MsgBox "El proceso ha finalizado correctamente.", vbInformation
    
    Exit Sub
    
eActualizarSocios:
    MuestraError Err.Number, "Actualizar socios", Err.Description
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


Private Function ConexionAriges() As Boolean
    
    ConexionAriges = False
    
    If vSeccion Is Nothing Then
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.SeccionSumi) Then
            ConexionAriges = vSeccion.AbrirAriges
        End If
    Else
        ' si el objeto existia: cerramos la conexion y volvemos crearlo
        vSeccion.CerrarAriges
        Set vSeccion = Nothing
        
        
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.SeccionSumi) Then
            ConexionAriges = vSeccion.AbrirAriges
        End If
    End If
    
End Function


Private Function ExisteCliente(Socio As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error GoTo eExisteCliente
    
    ExisteCliente = False
    
    Sql = "select * from sclien where codclien = " & DBSet(Socio, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, ConnAriges, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        ExisteCliente = True
    End If
    Exit Function
    
eExisteCliente:
    MuestraError Err.Number, "Existe Cliente", Err.Description
End Function


Private Function ComprobarReferenciales() As Boolean
Dim vSQL As String
Dim b As Boolean
Dim rsAriges As ADODB.Recordset

    On Error GoTo EComprobarReferenciales
    
    ComprobarReferenciales = False
            
    ' obtenemos los datos por defecto de la tabla spara1 de ariges1
    vSQL = "select defactividad, defenvio, defzona, defruta, defstituacion, deftarifa, defagente "
    vSQL = vSQL & " from spara1 "
    
    Set rsAriges = New ADODB.Recordset
    rsAriges.Open vSQL, ConnAriges, adOpenForwardOnly, adLockPessimistic, adCmdText

    ' comprobamos que se cumplan las claves referenciales
    b = True
    'actividad
    vSQL = ""
    vSQL = DevuelveDesdeBDNew(cAriges, "sactiv", "nomactiv", "codactiv", rsAriges!defactividad, "N")
    If vSQL = "" Then
        MsgBox "No existe la actividad por defecto en suministros: " & DBLet(rsAriges!defactividad, "N"), vbExclamation
        b = False
    End If
    'envio
    If b Then
        vSQL = ""
        vSQL = DevuelveDesdeBDNew(cAriges, "senvio", "nomenvio", "codenvio", rsAriges!defenvio, "N")
        If vSQL = "" Then
            MsgBox "No existe el codigo de envio por defecto en suministros: " & DBLet(rsAriges!defenvio, "N"), vbExclamation
            b = False
        End If
    End If
    'zona
    If b Then
        vSQL = ""
        vSQL = DevuelveDesdeBDNew(cAriges, "szonas", "nomzonas", "codzonas", rsAriges!defzona, "N")
        If vSQL = "" Then
            MsgBox "No existe el codigo de zona por defecto en suministros: " & DBLet(rsAriges!defzona, "N"), vbExclamation
            b = False
        End If
    End If
    'rutas
    If b Then
        vSQL = ""
        vSQL = DevuelveDesdeBDNew(cAriges, "srutas", "nomrutas", "codrutas", rsAriges!defruta, "N")
        If vSQL = "" Then
            MsgBox "No existe el codigo de ruta por defecto en suministros: " & DBLet(rsAriges!defrutas, "N"), vbExclamation
            b = False
        End If
    End If
    'situacion
    If b Then
        vSQL = ""
        vSQL = DevuelveDesdeBDNew(cAriges, "ssitua", "nomsitua", "codsitua", rsAriges!defstituacion, "N")
        If vSQL = "" Then
            MsgBox "No existe el codigo de situación por defecto en suministros: " & DBLet(rsAriges!defstituacion, "N"), vbExclamation
            b = False
        End If
    End If
    'agente
    If b Then
        vSQL = ""
        vSQL = DevuelveDesdeBDNew(cAriges, "sagent", "nomagent", "codagent", rsAriges!defagente, "N")
        If vSQL = "" Then
            MsgBox "No existe el codigo de agente por defecto en suministros: " & DBLet(rsAriges!defagente, "N"), vbExclamation
            b = False
        End If
    End If
    ' tarifa
    If b Then
        vSQL = ""
        vSQL = DevuelveDesdeBDNew(cAriges, "starif", "nomlista", "codlista", rsAriges!deftarifa, "N")
        If vSQL = "" Then
            MsgBox "No existe el codigo de tarifa por defecto en suministros: " & DBLet(rsAriges!deftarifa, "N"), vbExclamation
            b = False
        End If
    End If
    
    
    
    Set rsAriges = Nothing
    
    ComprobarReferenciales = b
    Exit Function
    
EComprobarReferenciales:
    Set rsAriges = Nothing
    
    MuestraError Err.Number, "Comprobar Referenciales", Err.Description
End Function
