VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAlmzTrasCampos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traspaso Campos "
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6480
   Icon            =   "frmAlmzTrasCampos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   6480
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
      Height          =   4725
      Left            =   -60
      TabIndex        =   3
      Top             =   -90
      Width           =   6555
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Text5"
         Top             =   1035
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   0
         Top             =   1035
         Width           =   750
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4905
         TabIndex        =   2
         Top             =   3780
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   1
         Top             =   3780
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   210
         TabIndex        =   4
         Top             =   2730
         Width           =   6030
         _ExtentX        =   10636
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1140
         Top             =   3990
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "doc"
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1290
         MouseIcon       =   "frmAlmzTrasCampos.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
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
         Index           =   0
         Left            =   330
         TabIndex        =   8
         Top             =   1050
         Width           =   630
      End
      Begin VB.Label lblProgres 
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   3120
         Width           =   6195
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   5
         Top             =   3480
         Width           =   6195
      End
   End
End
Attribute VB_Name = "frmAlmzTrasCampos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' PROGRAMA DE TRASPASO DE ENTRADAS DE BASCULA DE ALMAZARA
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
Private WithEvents frmConce As frmConceConta 'conceptos de contabilidad
Attribute frmConce.VB_VarHelpID = -1
Private WithEvents frmTDia As frmDiaConta 'diarios de contabilidad
Attribute frmTDia.VB_VarHelpID = -1
Private WithEvents frmVar As frmComVar 'Variedades de comercial
Attribute frmVar.VB_VarHelpID = -1


'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim indCodigo As Integer 'indice para txtCodigo
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Cad As String
Dim cadTabla As String

Dim vContad As Long

Dim PrimeraVez As Boolean

Dim Socio As String
Dim Municipio As String
Dim Poligono As String
Dim Parcela As String
Dim Subparcela As String
Dim Variedad As String
Dim Superficie As String
Dim NroCampo As Long

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim Sql As String
Dim I As Byte
Dim cadwhere As String
Dim b As Boolean
Dim NomFic As String
Dim Cadena As String
Dim cadena1 As String

On Error GoTo eError

    If Not DatosOk Then Exit Sub
    
    Me.CommonDialog1.Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist


    Me.CommonDialog1.DefaultExt = "csv"
    'cadena = Format(CDate(txtcodigo(0).Text), FormatoFecha)
    CommonDialog1.FilterIndex = 1
    Me.CommonDialog1.FileName = "campos.csv"

    
    Me.CommonDialog1.CancelError = True
    
    Me.CommonDialog1.ShowOpen
    
    If Me.CommonDialog1.FileName <> "" Then
        InicializarVbles
            '========= PARAMETROS  =============================
        'Añadir el parametro de Empresa
        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1


        If ProcesarFichero(Me.CommonDialog1.FileName) Then
            
            MsgBox "Proceso realizado correctamente.", vbExclamation
            Pb1.visible = False
            lblProgres(0).Caption = ""
            lblProgres(1).Caption = ""
            
            cadTabla = "tmpinformes"
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            
            Sql = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo
            
            If TotalRegistros(Sql) <> 0 Then
                MsgBox "Han habido errores en el Traspaso de Campos. ", vbExclamation
                cadTitulo = "Errores en el Traspaso de Campos"
                cadNombreRPT = "rErroresTrasCampos.rpt"
                
                cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                
                LlamarImprimir
    
            End If
            cmdCancel_Click
        Else
            MsgBox "No se ha podido realizar el proceso.", vbExclamation
        End If
    
    Else
        MsgBox "No ha seleccionado ningún fichero", vbExclamation
        Exit Sub
    End If
             
             
eError:
    If Err.Number = 32755 Then Exit Sub ' le han dado a cancelar

'    If Err.Number <> 0 Or Not b Then
'        conn.RollbackTrans
'        MsgBox "No se ha podido realizar el proceso. LLame a Ariadna.", vbExclamation
'    Else
'        conn.CommitTrans
'    End If
'
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

    For H = 2 To 2
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H


         
    FrameCobrosVisible True, H, W
    Pb1.visible = False
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
'   Me.Width = w + 70
'   Me.Height = h + 350
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DesBloqueoManual ("TRASPOST")
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
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

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Sql As String
    
    b = True
 
    If txtcodigo(2).Text = "" Then
        MsgBox "Debe introducir el código de variedad. Reintroduzca.", vbExclamation
        b = False
        PonerFoco txtcodigo(2)
    Else
        Sql = "select count(*) from variedades where codvarie = " & DBSet(txtcodigo(2).Text, "N")
        If TotalRegistros(Sql) = 0 Then
            MsgBox "Código de variedad no existe. Reintroduzca.", vbExclamation
            b = False
            PonerFoco txtcodigo(2)
        End If
    End If
 
    DatosOk = b

End Function



Private Function RecuperaFichero() As Boolean
Dim NF As Integer

    RecuperaFichero = False
    NF = FreeFile
    Open App.Path For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    Line Input #NF, Cad
    Close #NF
    If Cad <> "" Then RecuperaFichero = True
    
End Function


Private Function ProcesarFichero(nomFich As String) As Boolean
Dim NF As Long
Dim Cad As String
Dim I As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim Numreg As Long
Dim Sql As String
Dim Sql1 As String
Dim total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean
Dim NomFic As String

    
    ProcesarFichero = False
    
    InicializarTabla
    
    NF = FreeFile
    
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, Cad ' saltamos la primera linea
    Line Input #NF, Cad
    I = 1
    
    lblProgres(0).Caption = "Procesando Fichero: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
        
    NroCampo = DevuelveValor("select max(codcampo) from rcampos")
        
    b = True
    While Not EOF(NF) And b
        I = I + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(Cad)
        lblProgres(1).Caption = "Linea " & I
        Me.Refresh
        
        NroCampo = NroCampo + 1
        
        If Cad <> ";;;;;;;;;" Then b = InsertarLinea(Cad)
        
        Line Input #NF, Cad
    Wend
    Close #NF
    
    If Cad <> "" And b Then
        NroCampo = NroCampo + 1
        
        If Cad <> ";;;;;;;;;" Then b = InsertarLinea(Cad)
    End If
    
    ProcesarFichero = b
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    
End Function
                
Private Function ProcesarFichero2(nomFich As String) As Boolean
Dim NF As Long
Dim Cad As String
Dim I As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim Numreg As Long
Dim Sql As String
Dim Sql1 As String
Dim total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim b As Boolean

    On Error GoTo eProcesarFichero2
    
    ProcesarFichero2 = False
    
    NF = FreeFile
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, Cad
    Line Input #NF, Cad
    I = 1
    
    lblProgres(0).Caption = "Insertando en Tabla temporal: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0

    b = True

    While Not EOF(NF) And b
        I = I + 1
        
        
        Me.Pb1.Value = Me.Pb1.Value + Len(Cad)
        lblProgres(1).Caption = "Linea " & I
        Me.Refresh
        
        If Cad <> ";;;;;;;;;" Then b = ComprobarRegistro(Cad)
        
        Line Input #NF, Cad
    Wend
    Close #NF
    
    If Cad <> "" Then
        I = I + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(Cad)
        lblProgres(1).Caption = "Linea " & I
        Me.Refresh
        
        If Cad <> ";;;;;;;;;" Then b = ComprobarRegistro(Cad)
    
    End If
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

    ProcesarFichero2 = b
    Exit Function

eProcesarFichero2:
    ProcesarFichero2 = False
End Function
                
            
Private Function ComprobarRegistro(Cad As String) As Boolean
Dim Sql As String
Dim Mens As String
Dim Cadena As String

    On Error GoTo eComprobarRegistro

    ComprobarRegistro = True

    CargarVariables Cad

    ' comprobamos que me han puesto los datos de busqueda de parcela
    If Poligono = "" Or Parcela = "" Or Subparcela = "" Then
        Mens = "Datos de poligono/parcela/subparcela incorrectos"
        Sql = "insert into tmpinformes (codusu, importe1,  " & _
              "importe2, nombre2, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Socio, "N") & ","
        Sql = Sql & "0," & DBSet(Cadena, "T") & "," & DBSet(Mens, "T") & ")"
        
        conn.Execute Sql
    Else
        Cadena = Format(CCur(Poligono), "0000") & "-" & Format(CCur(Parcela), "0000") & "-" & Subparcela
    End If
    
    
    'Comprobamos que el socio existe
    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "rsocios", "codsocio", "codsocio", Socio, "N")
    If Sql = "" Then
        Mens = "No existe el socio"
        Sql = "insert into tmpinformes (codusu, importe1,  " & _
              "importe2, nombre2, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Socio, "N") & ","
        Sql = Sql & "0," & DBSet(Cadena, "T") & "," & DBSet(Mens, "T") & ")"
        
        conn.Execute Sql
    End If
    
    'Comprobamos que la partida existe
    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "rpartida", "codparti", "codparti", Municipio, "N")
    If Sql = "" Then
        Mens = "No existe la partida " & Municipio
        Sql = "insert into tmpinformes (codusu, importe1,  " & _
              "importe2, nombre2, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Socio, "N") & ","
        Sql = Sql & "0," & DBSet(Cadena, "T") & "," & DBSet(Mens, "T") & ")"
        
        conn.Execute Sql
    End If
    
    
    
    ' comprobamos que el campo no está creado ya
    If ComprobarCero(Poligono) <> 0 And ComprobarCero(Parcela) <> 0 And ComprobarCero(Subparcela) <> 0 Then
        Sql = "select codcampo from rcampos where poligono = " & DBSet(Poligono, "N")
        Sql = Sql & " and parcela = " & DBSet(Parcela, "N")
        Sql = Sql & " and subparcela = " & DBSet(Subparcela, "N")


        If DevuelveValor(Sql) <> 0 Then
            Mens = "El campo existe Nº." & DevuelveValor(Sql)
            Sql = "insert into tmpinformes (codusu, importe1, " & _
                  "importe2, nombre2, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(Socio, "N") & ","
            Sql = Sql & "1," & DBSet(Cadena, "T") & "," & DBSet(Mens, "T") & ")"

            conn.Execute Sql
        End If
    End If
    
eComprobarRegistro:
    If Err.Number <> 0 Then
        ComprobarRegistro = False
    End If
End Function

            
Private Function InsertarLinea(Cad As String) As Boolean
Dim Sql As String
Dim CodZona As String
Dim vSuperficie As Currency
Dim HayError As Boolean
Dim Mens As String
Dim Cadena As String

    On Error GoTo EInsertarLinea

    InsertarLinea = True
    
    
    CargarVariables Cad
    
    
    HayError = False
    
    ' comprobaciones para poder insertar
    
    Cadena = ""
    ' comprobamos que me han puesto los datos de busqueda de parcela
    If Poligono = "" Or Parcela = "" Or Subparcela = "" Then
        Mens = "Datos de poligono/parcela/subparcela incorrectos"
        Sql = "insert into tmpinformes (codusu, importe1,  " & _
              "importe2, nombre2, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Socio, "N") & ","
        Sql = Sql & "0," & DBSet(Cadena, "T") & "," & DBSet(Mens, "T") & ")"
        
        conn.Execute Sql
        
        HayError = True
    Else
        Cadena = Format(CCur(Poligono), "0000") & "-" & Format(CCur(Parcela), "0000") & "-" & Subparcela
    End If
    
    
    'Comprobamos que el socio existe
    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "rsocios", "codsocio", "codsocio", Socio, "N")
    If Sql = "" Then
        Mens = "No existe el socio"
        Sql = "insert into tmpinformes (codusu, importe1,  " & _
              "importe2, nombre2, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Socio, "N") & ","
        Sql = Sql & "0," & DBSet(Cadena, "T") & "," & DBSet(Mens, "T") & ")"
        
        conn.Execute Sql
        
        HayError = True
    End If
    
    'Comprobamos que la partida existe
    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "rpartida", "codparti", "codparti", Municipio, "N")
    If Sql = "" Then
        Mens = "No existe la partida " & Municipio
        Sql = "insert into tmpinformes (codusu, importe1,  " & _
              "importe2, nombre2, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Socio, "N") & ","
        Sql = Sql & "0," & DBSet(Cadena, "T") & "," & DBSet(Mens, "T") & ")"
        
        conn.Execute Sql
        
        HayError = True
    End If
    
    If HayError Then Exit Function
    
    ' comprobamos que el campo no está creado ya
    If ComprobarCero(Poligono) <> 0 And ComprobarCero(Parcela) <> 0 And ComprobarCero(Subparcela) <> 0 Then
        Sql = "select codcampo from rcampos where poligono = " & DBSet(Poligono, "N")
        Sql = Sql & " and parcela = " & DBSet(Parcela, "N")
        Sql = Sql & " and subparce = " & DBSet(Subparcela, "N")


        If DevuelveValor(Sql) <> 0 Then Exit Function
    End If
    
    
    
    
    CodZona = DevuelveValor("select codzonas from rpartida where codparti = " & DBSet(Municipio, "N"))
    vSuperficie = Round2(Superficie * vParamAplic.Faneca, 4)
    
    
    ' insertamos en la tabla de rhisfruta
    Sql = "insert into rcampos (codcampo, codsocio, codpropiet, codvarie, codparti, "
    Sql = Sql & "codzonas, fecaltas, supsigpa, supcoope, supcatas, supculti, codsitua, "
    Sql = Sql & "poligono, parcela, subparce, asegurado, tipoparc, recintos, nrocampo, recolect) VALUES ("
    Sql = Sql & DBSet(NroCampo, "N") & ","
    Sql = Sql & DBSet(Socio, "N") & ","
    Sql = Sql & DBSet(Socio, "N") & ","
    Sql = Sql & DBSet(txtcodigo(2).Text, "N") & ","
    Sql = Sql & DBSet(Municipio, "N") & ","
    Sql = Sql & DBSet(CodZona, "N") & ","
    Sql = Sql & DBSet(Now, "F") & ","
    Sql = Sql & DBSet(vSuperficie, "N") & "," ' superficie en hectareas
    Sql = Sql & DBSet(vSuperficie, "N") & ","
    Sql = Sql & DBSet(vSuperficie, "N") & ","
    Sql = Sql & DBSet(vSuperficie, "N") & ","
    Sql = Sql & "0," ' situacion
    Sql = Sql & DBSet(Poligono, "N") & ","
    Sql = Sql & DBSet(Parcela, "N") & ","
    Sql = Sql & DBSet(Subparcela, "T") & ","
    Sql = Sql & "0,0,0,"
    Sql = Sql & DBSet(NroCampo, "N") & ","
    Sql = Sql & "0)"
    
    
    conn.Execute Sql
    Exit Function
    
EInsertarLinea:
    If Err.Number <> 0 Then
        InsertarLinea = False
        MsgBox "Error en Insertar Linea " & Err.Description, vbExclamation
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
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub InicializarTabla()
Dim Sql As String
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    
    conn.Execute Sql
End Sub

Private Sub AbrirFrmVariedad(indice As Integer)
    indCodigo = indice
    Set frmVar = New frmComVar
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub


Private Sub CargarVariables(Cad As String)
    
        Socio = ""
        Municipio = ""
        Poligono = ""
        Parcela = ""
        Subparcela = ""
        Variedad = ""
        Superficie = ""

        Socio = RecuperaValorNew(Cad, ";", 1)
        Municipio = RecuperaValorNew(Cad, ";", 3)
        Poligono = RecuperaValorNew(Cad, ";", 4)
        Parcela = RecuperaValorNew(Cad, ";", 5)
        Subparcela = RecuperaValorNew(Cad, ";", 6)
        Superficie = RecuperaValorNew(Cad, ";", 7)
        
    
    
End Sub


Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 2 'VARIEDADES
            AbrirFrmVariedad (Index)
    End Select
    PonerFoco txtcodigo(Index)
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtcodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 2: KEYBusqueda KeyAscii, 2 'variedad
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
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 2 'variedades
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
    
    
    End Select
End Sub

