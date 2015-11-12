VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTrasTraza 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6600
   Icon            =   "frmTrasTraza.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNombre 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   68
      Left            =   3030
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Text5"
      Top             =   4965
      Width           =   3375
   End
   Begin VB.TextBox txtcodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   73
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   7
      Top             =   4950
      Width           =   1095
   End
   Begin VB.Frame FrameTrasTraza 
      Height          =   4545
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   6555
      Begin MSComctlLib.ProgressBar pb1 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdAcepTras 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3930
         TabIndex        =   4
         Top             =   3690
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelTras 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5145
         TabIndex        =   5
         Top             =   3690
         Width           =   975
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   570
         Top             =   3390
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "doc"
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   2
         Left            =   270
         TabIndex        =   9
         Top             =   3270
         Width           =   6195
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Proceso que realiza la lectura de entradas de traza para incorporarlas a la aplicación."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   525
         Index           =   37
         Left            =   300
         TabIndex        =   3
         Top             =   630
         Width           =   5820
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   2
         Top             =   2865
         Width           =   6195
      End
      Begin VB.Label lblProgres 
         Height          =   405
         Index           =   0
         Left            =   270
         TabIndex        =   1
         Top             =   2340
         Width           =   6195
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5280
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   68
      Left            =   1620
      MouseIcon       =   "frmTrasTraza.frx":000C
      MousePointer    =   4  'Icon
      ToolTipText     =   "Buscar clase"
      Top             =   4950
      Width           =   240
   End
End
Attribute VB_Name = "frmTrasTraza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: LAURA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public OpcionListado As Byte
    '==== Listados BASICOS ====
    '=============================
    ' 11 .- Listado de Entradas de Pesadas
    ' 12 .- Listado de Calidades
    ' 13 .- Listado de Socios por Sección
    ' 14 .- Listado de Entradas en Bascula
    ' 15 .- Listado de Campos
    ' 16 .- Listado de Entradas clasificacion
    ' 17 .- Reimpresion de albaranes de Clasificacion
    ' 18 .- Informe de Kilos/Gastos (rhisfruta)
    ' 19 .- Grabación de Fichero Agriweb
    ' 20 .- Informe de Kilos Por Producto
    ' 21 .- Traspaso desde el calibrador
    ' 22 .- Traspaso TRAZABILIDAD
    
    
    ' 23 .- Baja de Socios (dentro del mantenimiento socios)
    
    ' 24 .- Traspaso de Facturas Cooperativa ( traspaso liquidacion )
    ' 25 .- Listado de Kilos recolectados socio / cooperativa
    ' 26 .- Traspaso de ROPAS solo para Catadau
    ' 27 .- Traspaso de datos a Almazara solo para Mogente
    
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
Private WithEvents frmProd As frmComercial 'Ayuda Productos de comercial
Attribute frmProd.VB_VarHelpID = -1
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
Private WithEvents frmMens1 As frmMensajes 'Mensajes
Attribute frmMens1.VB_VarHelpID = -1
Private WithEvents frmSitu As frmManSituacion 'Situacion de socio
Attribute frmSitu.VB_VarHelpID = -1
Private WithEvents frmCoop As frmManCoope 'Cooperativa
Attribute frmCoop.VB_VarHelpID = -1


'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe
Private cadSelect1 As String 'Cadena para comprobar si hay datos antes de abrir Informe


Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Tabla As String
Dim Tabla1 As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String

Dim indice As Integer

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


Private Sub cmdAcepTras_Click()
Dim SQL As String
Dim i As Byte
Dim cadWHERE As String
Dim b As Boolean
Dim NomFic As String
Dim Cadena As String
Dim cadena1 As String
Dim Directorio As String
Dim Fec As String
Dim nomDir As String

Dim nRegs As Long
Dim cadTabla As String
Dim NomFic1 As String

Dim File1 As FileSystemObject

On Error GoTo eError

    If Not DatosOk Then Exit Sub

    
    Me.CommonDialog1.Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist


    Me.CommonDialog1.DefaultExt = "txt"
    CommonDialog1.Filter = "Archivos TXT|*.txt|"
    CommonDialog1.FilterIndex = 1
    Me.CommonDialog1.FileName = "*.txt"
    
    
    Me.CommonDialog1.CancelError = True
    Me.CommonDialog1.ShowOpen
    Set File1 = New FileSystemObject
    
    Directorio = File1.GetParentFolderName(Me.CommonDialog1.FileName)

    If Directorio <> "" Then

        SQL = "DROP TABLE IF EXISTS tmpentradas; "
        conn.Execute SQL
        
        SQL = "CREATE TEMPORARY TABLE `tmpentradas` ("
        SQL = SQL & "`numnotac` int(7), "
        SQL = SQL & "`fechaent` varchar(8), "
        SQL = SQL & "`codsocio` int(6), "
        SQL = SQL & "`codcampo` int(8), "
        SQL = SQL & "`codpobla` varchar(6), "
        SQL = SQL & "`codprodu` int(7), "
        SQL = SQL & "`codvarie` int(7), "
        SQL = SQL & "`kilosbru` int(8), "
        SQL = SQL & "`numcajo1` int(7), "
        SQL = SQL & "`numcajo2` int(7), "
        SQL = SQL & "`numcajo3` int(7), "
        SQL = SQL & "`tipoentr` varchar(1),"
        SQL = SQL & "`recolect` varchar(1)"
        SQL = SQL & " ) ENGINE=InnoDB DEFAULT CHARSET=latin1"
    
    
        conn.Execute SQL
        
        conn.BeginTrans

        nomDir = Directorio & "\"

        NomFic = Dir(nomDir & "*.txt")  ' Recupera la primera entrada.
        NomFic1 = nomDir & "E*.txt"
        
        ' Cargamos en la tabla temporal todas las entradas de los ficheros del directorio seleccionado
        
        Do While NomFic <> ""   ' Inicia el bucle.
           ' Ignora el directorio actual y el que lo abarca.
           If NomFic <> "." And NomFic <> ".." And UCase(Mid(NomFic, 1, 1)) = "E" Then
              ' Realiza una comparación a nivel de bit para asegurarse de que MiNombre es un directorio.
'              If (GetAttr(nomDir & NomFic) And vbArchive) = vbArchive Then
              
                lblProgres(0).Caption = "Procesando Fichero: " & NomFic
              
                SQL = "load data local infile '" & Replace(nomDir & NomFic, "\", "/") & "' into table `tmpentradas` fields terminated by '|' lines terminated by '\n' "
                SQL = SQL & "(`numnotac`,`fechaent`,`codsocio`,`codcampo`,`codpobla`,`codprodu`,`codvarie`,`kilosbru`,`numcajo1`,`numcajo2`,`numcajo3`,`tipoentr`,`recolect`)  "
                conn.Execute SQL
                
'              End If
           End If
           NomFic = Dir   ' Obtiene siguiente entrada.
        Loop

        SQL = "select count(*) from tmpentradas"
        nRegs = TotalRegistros(SQL)
        If nRegs <> 0 Then
            pb1.visible = True
            pb1.Max = nRegs
            pb1.Value = 0
            Me.Refresh
            DoEvents
                
            InicializarVbles
                
                '========= PARAMETROS  =============================
            'Añadir el parametro de Empresa
            cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
            numParam = numParam + 1
    
            If ComprobarErrores(pb1) Then
                cadTabla = "tmpinformes"
                cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                
                SQL = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo
                
                If TotalRegistros(SQL) <> 0 Then
                    MsgBox "Hay errores en el Traspaso de Trazabilidad." & vbCrLf & "   Debe corregirlos previamente.", vbExclamation
                    cadTitulo = "Errores de Traspaso de TRAZABILIDAD"
                    cadNombreRPT = "rErroresTrasTraza.rpt"
                    LlamarImprimir
                    conn.RollbackTrans
                    lblProgres(0).Caption = ""
                    lblProgres(1).Caption = ""
                    lblProgres(2).Caption = ""
                    Exit Sub
                Else
                    b = CargarEntradas()
                End If
            Else
                b = False
            End If
                
        End If

    End If
    
eError:
    If Err.Number = 32755 Then Exit Sub ' le han dado a cancelar

    If Err.Number <> 0 Or Not b Then
        conn.RollbackTrans
        MsgBox "No se ha podido realizar el proceso.", vbExclamation
    Else
        conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        pb1.visible = False
        lblProgres(0).Caption = ""
        lblProgres(1).Caption = ""
        lblProgres(2).Caption = ""

        BorrarArchivo NomFic1
        cmdCancelTras_Click
    End If
    
End Sub

Private Sub cmdCancelTras_Click()
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

    
    'Ocultar todos los Frames de Formulario
    FrameTrasTraza.visible = False
    '###Descomentar
'    CommitConexion
        
    FrameTrasTrazaVisible True, H, W
    pb1.visible = False
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
'    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub


Private Sub FrameTrasTrazaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameTrasTraza.visible = visible
    If visible = True Then
        Me.FrameTrasTraza.Top = -90
        Me.FrameTrasTraza.Left = 0
        Me.FrameTrasTraza.Height = 4665
        Me.FrameTrasTraza.Width = 6555
        W = Me.FrameTrasTraza.Width
        H = Me.FrameTrasTraza.Height
    End If
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadSelect1 = ""
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
        .EnvioEMail = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .ConSubInforme = ConSubInforme
        .Opcion = OpcionListado
        .Show vbModal
    End With
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
        .Opcion = OpcionListado
'        .ExportarPDF = (chkEMAIL.Value = 1)
        .Show vbModal
    End With
    
'    If Me.chkEMAIL.Value = 1 Then
'    '####Descomentar
'        If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
'    End If
    Unload Me
End Sub


Private Function ComprobarErrores(ByRef pb1 As ProgressBar) As Boolean
Dim NF As Long
Dim Cad As String
Dim i As Integer
Dim longitud As Long
Dim RS As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim Numreg As Long
Dim SQL As String
Dim SQL1 As String
Dim total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean
Dim Mens As String
Dim Tipo As Integer
Dim FechaEnt As String
Dim Variedad As String


    On Error GoTo eComprobarErrores

    ComprobarErrores = False
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    i = 0
    lblProgres(1).Caption = "Comprobando errores Tabla temporal entradas "
    
    SQL = "select * from tmpentradaS"
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    b = True
    i = 0
    While Not RS.EOF And b
        i = i + 1

        Me.pb1.Value = Me.pb1.Value + 1
        lblProgres(2).Caption = "Linea " & i
        Me.Refresh

        Variedad = Format(RS!codprodu, "00") & Format(RS!codvarie, "00")

        ' comprobamos la fecha
        FechaEnt = DBLet(RS!FechaEnt, "T")
        If Not EsFechaOK(FechaEnt) Then
            Mens = "Fecha incorrecta"
            SQL = "insert into tmpinformes (codusu, campo1, codigo1, importe1, importe2, fecha1, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(Variedad, "N") & "," & DBSet(RS!Codsocio, "N") & "," & _
                  DBSet(RS!codCampo, "N") & "," & DBSet(RS!Numnotac, "N") & "," & _
                  DBSet(FechaEnt, "F") & "," & DBSet(Mens, "T") & ")"
            conn.Execute SQL
        End If

        ' comprobamos que exista la variedad
        SQL = "select count(*) from variedades where codvarie = " & DBSet(Variedad, "N")
        If TotalRegistros(SQL) = 0 Then
            Mens = "Variedad no existe"
            SQL = "insert into tmpinformes (codusu, campo1, codigo1, importe1, importe2, fecha1, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(Variedad, "N") & "," & DBSet(RS!Codsocio, "N") & "," & _
                  DBSet(RS!codCampo, "N") & "," & DBSet(RS!Numnotac, "N") & "," & _
                  DBSet(FechaEnt, "F") & "," & DBSet(Mens, "T") & ")"
            conn.Execute SQL
        Else
            If EsVariedadGrupo5(Variedad) Or EsVariedadGrupo6(Variedad) Then
                Mens = "Variedad no es del grupo correcto."
                SQL = "insert into tmpinformes (codusu, campo1, codigo1, importe1, importe2, fecha1, nombre1) values (" & _
                      vUsu.Codigo & "," & DBSet(Variedad, "N") & "," & DBSet(RS!Codsocio, "N") & "," & _
                      DBSet(RS!codCampo, "N") & "," & DBSet(RS!Numnotac, "N") & "," & _
                      DBSet(FechaEnt, "F") & "," & DBSet(Mens, "T") & ")"
                conn.Execute SQL
            End If
        End If

        ' comprobamos que exista el socio
        SQL = "select count(*) from rsocios where codsocio = " & DBSet(RS!Codsocio, "N")
        If TotalRegistros(SQL) = 0 Then
            Mens = "Socio no existe"
            SQL = "insert into tmpinformes (codusu, campo1, codigo1, importe1, importe2, fecha1, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(Variedad, "N") & "," & DBSet(RS!Codsocio, "N") & "," & _
                  DBSet(RS!codCampo, "N") & "," & DBSet(RS!Numnotac, "N") & "," & _
                  DBSet(FechaEnt, "F") & "," & DBSet(Mens, "T") & ")"
            conn.Execute SQL
        End If

        ' comprobamos que exista el campo
        SQL = "select count(*) from rcampos where codsocio = " & DBSet(RS!Codsocio, "N")
        SQL = SQL & " and nrocampo = " & DBSet(RS!codCampo, "N")
        SQL = SQL & " and codvarie = " & DBSet(Variedad, "N")
        SQL = SQL & " and fecbajas is null "
        If TotalRegistros(SQL) = 0 Then
            Mens = "Campo no existe o con fecha de baja"
            SQL = "insert into tmpinformes (codusu, campo1, codigo1, importe1, importe2, fecha1, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(Variedad, "N") & "," & DBSet(RS!Codsocio, "N") & "," & _
                  DBSet(RS!codCampo, "N") & "," & DBSet(RS!Numnotac, "N") & "," & _
                  DBSet(FechaEnt, "F") & "," & DBSet(Mens, "T") & ")"
            conn.Execute SQL
        End If

        ' comprobamos que no exista mas de un campo con ese numero de orden campo (scampo.codcampo MB)
        SQL = "select count(*) from rcampos where codsocio = " & DBSet(RS!Codsocio, "N")
        SQL = SQL & " and nrocampo = " & DBSet(RS!codCampo, "N")
        SQL = SQL & " and codvarie = " & DBSet(Variedad, "N")
        If TotalRegistros(SQL) > 1 Then
            Mens = "Campo con más de un registro"
            SQL = "insert into tmpinformes (codusu, campo1, codigo1, importe1, importe2, fecha1, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(Variedad, "N") & "," & DBSet(RS!Codsocio, "N") & "," & _
                  DBSet(RS!codCampo, "N") & "," & DBSet(RS!Numnotac, "N") & "," & _
                  DBSet(FechaEnt, "F") & "," & DBSet(Mens, "T") & ")"
            conn.Execute SQL
        End If

        If Not EsVariedadGrupo5(Variedad) And Not EsVariedadGrupo6(Variedad) Then
            ' comprobamos que no exista el albaran en rentradas
            SQL = "select count(*) from rentradas where numnotac = " & DBSet(RS!Numnotac, "N")
            If TotalRegistros(SQL) > 0 Then
                Mens = "Nro.Nota ya existe en entradas báscula"
                SQL = "insert into tmpinformes (codusu, campo1, codigo1, importe1, importe2, fecha1, nombre1) values (" & _
                      vUsu.Codigo & "," & DBSet(Variedad, "N") & "," & DBSet(RS!Codsocio, "N") & "," & _
                      DBSet(RS!codCampo, "N") & "," & DBSet(RS!Numnotac, "N") & "," & _
                      DBSet(FechaEnt, "F") & "," & DBSet(Mens, "T") & ")"
                conn.Execute SQL
            End If
    
            ' comprobamos que no exista el albaran en rclasifica
            SQL = "select count(*) from rclasifica where numnotac = " & DBSet(RS!Numnotac, "N")
            If TotalRegistros(SQL) > 0 Then
                Mens = "Nro.Nota ya existe en entradas clasificadas"
                SQL = "insert into tmpinformes (codusu, campo1, codigo1, importe1, importe2, fecha1, nombre1) values (" & _
                      vUsu.Codigo & "," & DBSet(Variedad, "N") & "," & DBSet(RS!Codsocio, "N") & "," & _
                      DBSet(RS!codCampo, "N") & "," & DBSet(RS!Numnotac, "N") & "," & _
                      DBSet(FechaEnt, "F") & "," & DBSet(Mens, "T") & ")"
                conn.Execute SQL
            End If
    
            ' comprobamos que no exista el albaran en el historico
            SQL = "select numalbar from rhisfruta_entradas where numnotac = " & DBSet(RS!Numnotac, "N")
            If DevuelveValor(SQL) <> 0 Then
                Mens = "Nro.Nota existe en hco.albarán:" & DevuelveValor(SQL)
                SQL = "insert into tmpinformes (codusu, campo1, codigo1, importe1, importe2, fecha1, nombre1) values (" & _
                      vUsu.Codigo & "," & DBSet(Variedad, "N") & "," & DBSet(RS!Codsocio, "N") & "," & _
                      DBSet(RS!codCampo, "N") & "," & DBSet(RS!Numnotac, "N") & "," & _
                      DBSet(FechaEnt, "F") & "," & DBSet(Mens, "T") & ")"
                conn.Execute SQL
            End If
        End If
        
        RS.MoveNext
    Wend
    Set RS = Nothing
    
    pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    lblProgres(2).Caption = ""

    ComprobarErrores = b
    Exit Function

eComprobarErrores:
    ComprobarErrores = False
End Function


Private Function CargarEntradas() As Boolean
Dim SQL As String
Dim SQL1 As String
Dim Sql3 As String
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Precio As Currency
Dim Transporte As Currency
Dim Kilos As Long

Dim AlbarAnt As Long
Dim KilosAlbar As Long
Dim KilosNetAnt As Long
Dim VarieAnt As Long
Dim CalidAnt As Long
Dim longitud As Long

Dim campo As Variant
Dim cadMen As String

Dim Variedad As String
Dim TipoEntr As Byte
Dim Recolect As Byte
Dim KilosNet As Long

Dim Fecha As String
Dim Hora As String


    On Error GoTo eCargarEntradas
    
    CargarEntradas = False
    
    
    lblProgres(1).Caption = "Cargando Entradas"
    
    SQL = "select count(*) from tmpentradas order by numnotac"
    longitud = TotalRegistros(SQL)
    
    pb1.visible = True
    Me.pb1.Max = longitud
    Me.Refresh
    Me.pb1.Value = 0
    
    SQL = "select * from tmpentradas order by numnotac"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        Me.pb1.Value = Me.pb1.Value + 1
        lblProgres(2).Caption = "Nro.Nota " & DBLet(RS!Numnotac, "N")
        Me.Refresh
        
        Variedad = Format(RS!codprodu, "00") & Format(RS!codvarie, "00")
            
        SQL = "insert into rentradas (numnotac,fechaent,horaentr,codvarie,codsocio,codcampo,tipoentr,recolect,codtrans,"
        SQL = SQL & "codcapat,codtarif,kilosbru,numcajo1,numcajo2,numcajo3,numcajo4,numcajo5,taracaja1,taracaja2,"
        SQL = SQL & "taracaja3,taracaja4,taracaja5,taravehi,kilosnet,nropesada,numlinea,transportadopor) values "
    
        campo = 0
        campo = DevuelveValor("select codcampo from rcampos where nrocampo = " & DBSet(RS!codCampo, "N") & " and codsocio=" & DBSet(RS!Codsocio, "N") & " and codvarie=" & DBSet(Variedad, "N"))
    
        ' fecha y hora en formato de mysql
        Fecha = "20" & Mid(RS!FechaEnt, 7, 2) & "-" & Mid(RS!FechaEnt, 4, 2) & "-" & Mid(RS!FechaEnt, 1, 2)
        Hora = Fecha & " " & Format(Now, "hh:mm:ss")
    
        SQL = SQL & "(" & DBSet(RS!Numnotac, "N") & ","
        SQL = SQL & DBSet(Fecha, "F") & ","
        SQL = SQL & DBSet(Hora, "FH") & ","
        SQL = SQL & DBSet(Variedad, "N") & ","
        SQL = SQL & DBSet(RS!Codsocio, "N") & ","
        SQL = SQL & DBSet(campo, "N") & ","
        
        Select Case DBLet(RS!TipoEntr, "T")
            Case "N"
                TipoEntr = 0
            Case "V"
                TipoEntr = 1
            Case "R"
                TipoEntr = 0
            Case "I"
                TipoEntr = 2
        End Select
        
        Select Case DBLet(RS!Recolect, "T")
            Case "C"
                Recolect = 0
            Case "S"
                Recolect = 1
        End Select
        
        KilosNet = DBLet(RS!KilosBru, "N") - _
                  (Round2(DBLet(RS!numcajo1, "N") * vParamAplic.PesoCaja1, 0) + _
                   Round2(DBLet(RS!numcajo2, "N") * vParamAplic.PesoCaja2, 0) + _
                   Round2(DBLet(RS!numcajo3, "N") * vParamAplic.PesoCaja3, 0))
        
        
        SQL = SQL & DBSet(TipoEntr, "N") & "," ' tipoentr 0=normal
        SQL = SQL & DBSet(Recolect, "N") & "," ' recolect 1=socio
        SQL = SQL & ValorNulo & "," 'transportista
        SQL = SQL & ValorNulo & "," 'capataz
        SQL = SQL & ValorNulo & "," 'tarifa
        SQL = SQL & DBSet(RS!KilosBru, "N") & ","
        SQL = SQL & DBSet(RS!numcajo1, "N") & ","
        SQL = SQL & DBSet(RS!numcajo2, "N") & ","
        SQL = SQL & DBSet(RS!numcajo3, "N") & ","
        SQL = SQL & ValorNulo & "," ' numcajo4
        SQL = SQL & ValorNulo & "," ' numcajo5
        SQL = SQL & DBSet(Round2(DBLet(RS!numcajo1, "N") * vParamAplic.PesoCaja1, 0), "N") & ","
        
        If DBLet(RS!numcajo2, "N") <> 0 Then
            SQL = SQL & DBSet(Round2(DBLet(RS!numcajo2, "N") * vParamAplic.PesoCaja2, 0), "N") & ","
        Else
            SQL = SQL & ValorNulo & ","
        End If
        If DBLet(RS!numcajo3, "N") <> 0 Then
            SQL = SQL & DBSet(Round2(DBLet(RS!numcajo3, "N") * vParamAplic.PesoCaja3, 0), "N") & ","
        Else
            SQL = SQL & ValorNulo & ","
        End If
        
        SQL = SQL & ValorNulo & ","
        SQL = SQL & ValorNulo & ","
        SQL = SQL & ValorNulo & "," ' taravehi
        SQL = SQL & DBSet(KilosNet, "N") & "," ' kilos netos
        SQL = SQL & ValorNulo & "," ' nro de pesada
        SQL = SQL & ValorNulo & "," ' nro de linea
        SQL = SQL & "0)" ' transportado por cooperativa
        
'        SQL = SQL & DBSet(RS!KilosNet, "N") & ","
'        SQL = SQL & ValorNulo & ","
'        SQL = SQL & DBSet(Transporte, "N") & ","
'        SQL = SQL & ValorNulo & ","
'        SQL = SQL & ValorNulo & ","
'        SQL = SQL & ValorNulo & ","
'        SQL = SQL & "0," 'tiporecol 0=horas 1=destajo no admite valor nulo
'        SQL = SQL & ValorNulo & ","
'        SQL = SQL & ValorNulo & ","
'        SQL = SQL & DBSet(RS!NumAlbar, "N") & ","
'        SQL = SQL & DBSet(RS!fecalbar, "F") & ",0)"
'
        If Not EsVariedadGrupo5(Variedad) And Not EsVariedadGrupo6(Variedad) Then
            conn.Execute SQL
        End If
        
        RS.MoveNext
    Wend
    
    Set RS = Nothing

    pb1.visible = False
    lblProgres(1).Caption = ""
    lblProgres(2).Caption = ""

    CargarEntradas = True
    Exit Function
    
eCargarEntradas:
    MuestraError Err.Number, "Cargar entradas", Err.Description
End Function

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim SQL As String
Dim Sql2 As String
Dim vClien As CSocio
' añadido
Dim Mens As String
Dim numfactu As String
Dim numser As String
Dim Fecha As Date
Dim vCont As CTiposMov
Dim tipoMov As String

    b = True
    DatosOk = b

End Function
