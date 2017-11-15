VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAlmzTrasBascula 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traspaso Entradas Báscula"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6480
   Icon            =   "frmAlmzTrasBascula.frx":0000
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
      TabIndex        =   2
      Top             =   -90
      Width           =   6555
      Begin VB.Frame Frame1 
         Caption         =   "Datos para Selección"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1545
         Left            =   240
         TabIndex        =   3
         Top             =   690
         Width           =   5955
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
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   540
            Width           =   3150
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo "
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
            Left            =   870
            TabIndex        =   7
            Top             =   570
            Width           =   495
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
         Left            =   5085
         TabIndex        =   1
         Top             =   3960
         Width           =   1065
      End
      Begin VB.CommandButton cmdAceptar 
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
         Left            =   3900
         TabIndex        =   0
         Top             =   3960
         Width           =   1065
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
Attribute VB_Name = "frmAlmzTrasBascula"
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
Private WithEvents frmMens As frmMensajes 'para la creacion del campo
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmMens2 As frmMensajes 'para la visualizacion previa del fichero
Attribute frmMens2.VB_VarHelpID = -1

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
Dim cad As String
Dim cadTabla As String

Dim vContad As Long

Dim PrimeraVez As Boolean


Dim NumNota As String
Dim Producto As String
Dim Variedad As String
Dim Socio As String
Dim Bruto As String
Dim Neto As String
Dim CPobla As String
Dim CPostal As String
Dim FechaEnt As String
Dim Poligono As String
Dim Parcela As String
Dim Subparcela As String
Dim Tara As String
Dim NroMuestra As String
Dim HoraEnt As String
Dim HayError As Boolean


Dim PP As String
Dim VV As String


Dim campo As String
Dim Continuar As Boolean

Dim SociosNoExisten As String
Dim VariedadesNoExisten As String


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me

End Sub

Private Sub cmdAceptar_Click()
Dim SQL As String
Dim I As Byte
Dim cadWHERE As String
Dim B As Boolean
Dim NomFic As String
Dim cadena As String
Dim cadena1 As String

On Error GoTo eError

    If Not DatosOK Then Exit Sub
    
    '[Monica]22/10/2015: nuevo traspaso para ABN
    If vParamAplic.Cooperativa = 1 Then
        cmdAceptarABN
        Exit Sub
    End If
    
    Me.CommonDialog1.Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist


    If vParamAplic.Cooperativa = 1 Then 'VALSUR
        Select Case Me.Combo1(0).ListIndex
            Case 0
                Me.CommonDialog1.DefaultExt = "asc"
                'cadena = Format(CDate(txtcodigo(0).Text), FormatoFecha)
                CommonDialog1.FilterIndex = 1
                Me.CommonDialog1.FileName = "albaran.asc"
    
            Case 1, 2
                Me.CommonDialog1.DefaultExt = "TXT"
                'cadena = Format(CDate(txtcodigo(0).Text), FormatoFecha)
                CommonDialog1.FilterIndex = 1
                Me.CommonDialog1.FileName = "tickets.txt"
        End Select
    Else ' Caso de Moixent
        Select Case Me.Combo1(0).ListIndex
            Case 0
                Me.CommonDialog1.DefaultExt = "TXT"
                'cadena = Format(CDate(txtcodigo(0).Text), FormatoFecha)
                CommonDialog1.FilterIndex = 1
                Me.CommonDialog1.FileName = "pesadas.txt"
        End Select
    End If
    
    Me.CommonDialog1.CancelError = True
    
    Me.CommonDialog1.ShowOpen
    
    If Me.CommonDialog1.FileName <> "" Then
        InicializarVbles
        InicializarTabla
            '========= PARAMETROS  =============================
        'Añadir el parametro de Empresa
        CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1

          
        If ProcesarFichero2(Me.CommonDialog1.FileName) Then
                cadTabla = "tmpinformes"
                cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                
                SQL = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo
                
                If TotalRegistros(SQL) <> 0 Then
'                If HayRegParaInforme(cadTABLA, cadSelect) Then
                    MsgBox "Hay errores en el Traspaso de Báscula. Debe corregirlos previamente.", vbExclamation
                    cadTitulo = "Errores de Traspaso de Báscula"
                    cadNombreRPT = "rErroresTrasBascula.rpt"
                    
                    Select Case vParamAplic.Cooperativa
                        Case 1 ' valsur
                            CadParam = CadParam & "pDescrip=""Coop/Pobl""|"
                            numParam = numParam + 1
                        Case 3 ' Moixent
                            CadParam = CadParam & "pDescrip=""Pol/Parc""|"
                            numParam = numParam + 1
                    End Select
                    
                    LlamarImprimir
                    Exit Sub
                Else
                    conn.BeginTrans
                    B = ProcesarFichero(Me.CommonDialog1.FileName)
                End If
        End If
    Else
        MsgBox "No ha seleccionado ningún fichero", vbExclamation
        Exit Sub
    End If
             
             
eError:
    If Err.Number = 32755 Then Exit Sub ' le han dado a cancelar

    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        MsgBox "No se ha podido realizar el proceso. LLame a Ariadna.", vbExclamation
    Else
        conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        Pb1.visible = False
        lblProgres(0).Caption = ""
        lblProgres(1).Caption = ""
'        BorrarArchivo Me.CommonDialog1.FileName
'        BorrarArchivo Replace(LCase(Me.CommonDialog1.FileName), "ventas", "totaliza")
'        If vParamAplic.Cooperativa = 1 Then
'        ' solo en el caso de alzira se graba en la srecau
'            BorrarArchivo Replace(LCase(Me.CommonDialog1.FileName), "ventas", "caja")
'            BorrarArchivo Replace(LCase(Me.CommonDialog1.FileName), "ventas", "totales")
'        End If
        cmdCancel_Click
    End If
    
End Sub


Private Sub cmdAceptarABN()

Dim SQL As String
Dim I As Byte
Dim cadWHERE As String
Dim B As Boolean
Dim NomFic As String
Dim cadena As String
Dim cadena1 As String

On Error GoTo eError

    If Not DatosOK Then Exit Sub
    
    Me.CommonDialog1.Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist

    Me.CommonDialog1.DefaultExt = "csv"
    CommonDialog1.FilterIndex = 1
    Me.CommonDialog1.FileName = "albaranes.csv"
    
    
    Me.CommonDialog1.CancelError = True
    
    Me.CommonDialog1.ShowOpen
    
    If Me.CommonDialog1.FileName <> "" Then
        InicializarVbles
            '========= PARAMETROS  =============================
        'Añadir el parametro de Empresa
        CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1

        
        If CargaInicialABN(Me.CommonDialog1.FileName) Then
            Set frmMens2 = New frmMensajes
            
            frmMens2.OpcionMensaje = 63
            frmMens2.Show vbModal
        
            Set frmMens = Nothing
            
            If Not Continuar Then Exit Sub
            'If MsgBox("¿ Desea continuar con la carga del fichero ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
        Else
            MsgBox "No se ha podido realizar la carga inicial.", vbExclamation
        
        End If

        If ComprobarSociosVariedades(Me.CommonDialog1.FileName) Then
            If SociosNoExisten <> "" Then
                MsgBox "Los siguientes socios no existen, creelos y vuelva a importar: " & vbCrLf & vbCrLf & Mid(SociosNoExisten, 1, Len(SociosNoExisten) - 2), vbExclamation
                Exit Sub
            End If
            If VariedadesNoExisten <> "" Then
                MsgBox "Las siguientes variedades no existen, creelas y vuelva a importar: : " & vbCrLf & vbCrLf & Mid(VariedadesNoExisten, 1, Len(VariedadesNoExisten) - 2), vbExclamation
                Exit Sub
            End If
        End If

        If ProcesarFicheroABN(Me.CommonDialog1.FileName) Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
            Pb1.visible = False
            lblProgres(0).Caption = ""
            lblProgres(1).Caption = ""
            
            cadTabla = "tmpinformes"
            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            
            SQL = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo
            
            If TotalRegistros(SQL) <> 0 Then
                MsgBox "Han habido errores en el Traspaso de Báscula. ", vbExclamation
                cadTitulo = "Errores en el Traspaso de Báscula"
                cadNombreRPT = "rErroresTrasEntBascula.rpt"
                
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
End Sub
    


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Combo1(0).SetFocus
        Combo1(0).ListIndex = 0
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

    'IMAGES para busqueda
'     txtcodigo(0).Text = Format(Now - 1, "dd/mm/yyyy")

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, H, W
    Pb1.visible = False
    
    CargaCombo
    
'    Frame1.visible = (vParamAplic.Cooperativa = 1)
'    Frame1.Enabled = (vParamAplic.Cooperativa = 1)
    
    
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

Private Function DatosOK() As Boolean
Dim B As Boolean
Dim SQL As String
    
    B = True
    
    DatosOK = B

End Function



Private Function RecuperaFichero() As Boolean
Dim NF As Integer

    RecuperaFichero = False
    NF = FreeFile
    Open App.Path For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    Line Input #NF, cad
    Close #NF
    If cad <> "" Then RecuperaFichero = True
    
End Function


Private Function ProcesarFichero(nomFich As String) As Boolean
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
Dim NomFic As String

    ProcesarFichero = False
    NF = FreeFile
    
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, cad
    I = 0
    
    lblProgres(0).Caption = "Procesando Fichero: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
        
    B = True
    While Not EOF(NF) And B
        I = I + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & I
        Me.Refresh
        
        B = InsertarLinea(cad)
        
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" And B Then
        B = InsertarLinea(cad)
    End If
    
    ProcesarFichero = B
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

End Function

Private Function CargaInicialABN(nomFich As String) As Boolean
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
Dim NomFic As String

    On Error GoTo eProcesarFicheroABN



    CargaInicialABN = False
    NF = FreeFile
    
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, cad
    I = 0
    
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    
    lblProgres(0).Caption = "Carga inicial fichero: " & nomFich
    longitud = FileLen(nomFich)
    
        
    ' salto la primera linea que es la cabecera
    Line Input #NF, cad
    lblProgres(1).Caption = "Linea " & I
    Me.Refresh
    I = 1
    
        
    B = True
    While Not EOF(NF) And B
        I = I + 1
        
        lblProgres(1).Caption = "Linea " & I
        Me.Refresh
        
        cad = cad & ";"
        If Mid(cad, 1, 6) <> ";;;;;;" Then B = InsertarLineaPreviaABN(cad)
        
        If B Then
            If I > 20 Then
                CargaInicialABN = True
                Close #NF
                lblProgres(0).Caption = ""
                lblProgres(1).Caption = ""
                Exit Function
            End If
        End If
                
        
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" And B Then
        cad = cad & ";"
        If Mid(cad, 1, 6) <> ";;;;;;" Then B = InsertarLineaPreviaABN(cad)
    End If
    
    
    CargaInicialABN = B
    
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    
eProcesarFicheroABN:
    If Err.Number <> 0 Or Not B Then
    Else
    End If
 

End Function

Private Function InsertarLineaPreviaABN(cad As String) As Boolean
Dim SQL As String
Dim cadena As String

    On Error GoTo eInsertarLineaPreviaABN

    InsertarLineaPreviaABN = True
    
    CargarVariables cad
    
    ' insertamos la entrada
    cadena = vUsu.Codigo & "," & NumNota & "," & DBSet(FechaEnt, "F") & "," & DBSet(Socio, "N") & "," & DBSet(Variedad, "N") & "," & DBSet(Poligono, "N", "S") & "," & DBSet(Parcela, "N", "S") & "," & DBSet(Subparcela, "N", "S")
    cadena = cadena & "," & DBSet(ComprobarCero(Bruto) - ComprobarCero(Tara), "N")
    
    SQL = "insert into tmpinformes (codusu, importe1, fecha1, importe2, importe3, importe4, importe5, nombre1, importeb1) values "
    SQL = SQL & "(" & cadena & ")"
    conn.Execute SQL
    
    Exit Function
    
eInsertarLineaPreviaABN:
    InsertarLineaPreviaABN = False
    MuestraError Err.Number, "Insertar Linea Previa", Err.Description
End Function


Private Function ProcesarFicheroABN(nomFich As String) As Boolean
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
Dim NomFic As String

    On Error GoTo eProcesarFicheroABN


    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    conn.BeginTrans

    ProcesarFicheroABN = False
    NF = FreeFile
    
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, cad
    I = 0
    
    lblProgres(0).Caption = "Procesando Fichero: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
        
    ' salto la primera linea que es la cabecera
    Line Input #NF, cad
    Me.Pb1.Value = Me.Pb1.Value + Len(cad)
    lblProgres(1).Caption = "Linea " & I
    Me.Refresh
    I = 1
    
    
    ' procendencia de la entrada
    Select Case Combo1(0).ListIndex
        Case 0 ' bolbaite
            CPobla = 3
        Case 1 ' anna
            CPobla = 1
        Case 2 ' navarres
            CPobla = 5
    End Select
    CPostal = DevuelveDesdeBDNew(cAgro, "rcoope", "codposta", "codcoope", CPobla, "N")
        
        
    B = True
    While Not EOF(NF) And B
        I = I + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & I
        Me.Refresh
        
        cad = cad & ";"
        If Mid(cad, 1, 6) <> ";;;;;;" Then B = InsertarLineaABN(cad)
        
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" And B Then
        cad = cad & ";"
        If Mid(cad, 1, 6) <> ";;;;;;" Then B = InsertarLineaABN(cad)
    End If
    
    ProcesarFicheroABN = B
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    
eProcesarFicheroABN:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
    Else
        conn.CommitTrans
    End If
End Function
                
                
Private Function ComprobarSociosVariedades(nomFich As String) As Boolean
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
Dim NomFic As String

    On Error GoTo eComprobarSociosVariedades


    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    ComprobarSociosVariedades = False
    NF = FreeFile
    
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, cad
    I = 0
    
    lblProgres(0).Caption = "Procesando Fichero: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
        
    ' salto la primera linea que es la cabecera
    Line Input #NF, cad
    Me.Pb1.Value = Me.Pb1.Value + Len(cad)
    lblProgres(1).Caption = "Linea " & I
    Me.Refresh
    I = 1
    
    SociosNoExisten = ""
    VariedadesNoExisten = ""
        
    B = True
    While Not EOF(NF) And B
        I = I + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & I
        Me.Refresh
        
        cad = cad & ";"
        B = CompruebaSociosVariedades(cad)
        
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" And B Then
        cad = cad & ";"
        B = CompruebaSociosVariedades(cad)
    End If
    
    ComprobarSociosVariedades = B
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""
    
eComprobarSociosVariedades:
    If Err.Number <> 0 Or Not B Then
        MuestraError Err.Number, "Comprobar socios variedades", Err.Description
    End If
End Function
                
                
                
                
                
Private Function ProcesarFichero2(nomFich As String) As Boolean
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

    On Error GoTo eProcesarFichero2
    
    ProcesarFichero2 = False
    
    NF = FreeFile
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, cad
    I = 0
    
    lblProgres(0).Caption = "Insertando en Tabla temporal: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0

    B = True

    While Not EOF(NF) And B
        I = I + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & I
        Me.Refresh
        
        B = ComprobarRegistro(cad)
        
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" Then
        I = I + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & I
        Me.Refresh
        
        B = ComprobarRegistro(cad)
    
    End If
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

    ProcesarFichero2 = B
    Exit Function

eProcesarFichero2:
    ProcesarFichero2 = False
End Function
                
            
Private Function ComprobarRegistro(cad As String) As Boolean
Dim SQL As String
Dim Mens As String
Dim cadena As String

    On Error GoTo eComprobarRegistro

    ComprobarRegistro = True

    CargarVariables cad

    'Comprobamos fechas
    If Not EsFechaOK(FechaEnt) Then
        Mens = "Fecha incorrecta"
        SQL = "insert into tmpinformes (codusu, importe1, fecha1, importe2, importe3, importe4, " & _
              "nombre2, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(NumNota, "N") & "," & DBSet(FechaEnt, "F") & ","
        SQL = SQL & DBSet(Socio, "N") & "," & DBSet(Variedad, "N") & "," & _
              DBSet(Bruto, "N") & "," & DBSet(CPobla, "T") & "," & DBSet(Mens, "T") & ")"
        
        conn.Execute SQL
    End If
    
    
    'Comprobamos que existe la variedad
    SQL = ""
    SQL = DevuelveDesdeBDNew(cAgro, "variedades", "codvarie", "codvarie", Variedad, "N")
    If SQL = "" Then
        Mens = "No existe la variedad"
        SQL = "insert into tmpinformes (codusu, importe1, fecha1, importe2, importe3, " & _
              "importe4, nombre2, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(NumNota, "N") & "," & DBSet(FechaEnt, "F") & ","
        SQL = SQL & DBSet(Socio, "N") & "," & DBSet(Variedad, "N") & "," & _
              DBSet(Bruto, "N") & "," & DBSet(CPobla, "T") & "," & DBSet(Mens, "T") & ")"
              
        conn.Execute SQL
    End If
    
' han creado la variedad correspondiente
'
'    If Combo1(0).ListIndex = 1 Or Combo1(0).ListIndex = 2 Then
'        If CCur(PP) < 60 Or CCur(PP) > 64 Then
'            Mens = "Producto Erróneo"
'            SQL = "insert into tmpinformes (codusu, importe1, fecha1, importe2, importe3, " & _
'                  "importe4, nombre1) values (" & _
'                  vUsu.Codigo & "," & DBSet(NumNota, "N") & "," & DBSet(FechaEnt, "F") & ","
'            SQL = SQL & DBSet(Socio, "N") & "," & DBSet(Variedad, "N") & "," & _
'                  DBSet(Bruto, "N") & "," & DBSet(Mens, "T") & ")"
'
'            conn.Execute SQL
'        End If
'
'    End If
    
    'Comprobamos que el socio existe
    SQL = ""
    SQL = DevuelveDesdeBDNew(cAgro, "rsocios", "codsocio", "codsocio", Socio, "N")
    If SQL = "" Then
        Mens = "No existe el socio"
        SQL = "insert into tmpinformes (codusu, importe1, fecha1, importe2, importe3, " & _
              "importe4, nombre2, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(NumNota, "N") & "," & DBSet(FechaEnt, "F") & ","
        SQL = SQL & DBSet(Socio, "N") & "," & DBSet(Variedad, "N") & "," & _
                DBSet(Bruto, "N") & "," & DBSet(CPobla, "T") & "," & DBSet(Mens, "T") & ")"
        
        conn.Execute SQL
    End If
    
    '[Monica]22/12/2011: Nuevo control para todos en el que comprobamos que el socio no esté dado de baja
    If Not (EstaSocioDeAlta(Socio) And EstaSocioDeAltaSeccion(Socio, vParamAplic.SeccionAlmaz)) Then
        Mens = "Socio dado de baja"
        SQL = "insert into tmpinformes (codusu, importe1, fecha1, importe2, importe3, " & _
              "importe4, nombre2, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(NumNota, "N") & "," & DBSet(FechaEnt, "F") & ","
        SQL = SQL & DBSet(Socio, "N") & "," & DBSet(Variedad, "N") & "," & _
                DBSet(Bruto, "N") & "," & DBSet(CPobla, "T") & "," & DBSet(Mens, "T") & ")"
        
        conn.Execute SQL
    End If
    
    'Comprobamos que no exista el numero de nota
    SQL = ""
    SQL = DevuelveDesdeBDNew(cAgro, "rhisfruta", "numalbar", "numalbar", NumNota, "N")
    If SQL <> "" Then
        Mens = "Existe el Nro de nota"
        SQL = "insert into tmpinformes (codusu, importe1, fecha1, importe2, importe3, " & _
              "importe4, nombre2, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(NumNota, "N") & "," & DBSet(FechaEnt, "F") & ","
        SQL = SQL & DBSet(Socio, "N") & "," & DBSet(Variedad, "N") & "," & _
                DBSet(Bruto, "N") & "," & DBSet(CPobla, "T") & "," & DBSet(Mens, "T") & ")"
        
        conn.Execute SQL
    End If
    
    'Comprobamos que el codigo de cooperativa existe
    If vParamAplic.Cooperativa = 1 Then
        If Combo1(0).ListIndex = 0 Then
            SQL = ""
            SQL = DevuelveDesdeBDNew(cAgro, "rpueblos", "codpobla", "codpobla", CPobla, "T")
            Mens = "No existe la Cooperativa"
        Else
            SQL = ""
            SQL = DevuelveDesdeBDNew(cAgro, "rcoope", "codcoope", "codcoope", CPobla, "N")
            Mens = "No existe la Poblacion"
        End If
        
        If SQL = "" Then
            SQL = "insert into tmpinformes (codusu, importe1, fecha1, importe2, importe3, " & _
                  "importe4, nombre2, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(NumNota, "N") & "," & DBSet(FechaEnt, "F") & ","
            SQL = SQL & DBSet(Socio, "N") & "," & DBSet(Variedad, "N") & "," & _
                    DBSet(Bruto, "N") & "," & DBSet(CPobla, "T") & "," & DBSet(Mens, "T") & ")"
            
            conn.Execute SQL
        End If
    End If
    
    ' en el caso de moixent comprobamos que si nos han dado poligono y parcela, exista un campo
    ' si el poligono y parcela es cero no hacemos comprobacion
' 18/11/2009 de momento
'    If vParamAplic.Cooperativa = 3 Then
'        If ComprobarCero(Poligono) <> 0 And ComprobarCero(Parcela) <> 0 Then
'            SQL = "select codcampo from rcampos where codsocio = " & DBSet(Socio, "N")
'            SQL = SQL & " and codvarie = " & DBSet(Variedad, "N")
'            SQL = SQL & " and poligono = " & DBSet(Poligono, "N")
'            SQL = SQL & " and parcela = " & DBSet(Parcela, "N")
'
'            Cadena = Format(CCur(Poligono), "0000") & "-" & Format(CCur(Parcela), "0000")
'
'            If DevuelveValor(SQL) = 0 Then
'                Mens = "No existe el Campo"
'                SQL = "insert into tmpinformes (codusu, importe1, fecha1, importe2, importe3, " & _
'                      "importe4, nombre2, nombre1) values (" & _
'                      vUsu.Codigo & "," & DBSet(NumNota, "N") & "," & DBSet(FechaEnt, "F") & ","
'                SQL = SQL & DBSet(Socio, "N") & "," & DBSet(Variedad, "N") & "," & _
'                        DBSet(Bruto, "N") & "," & DBSet(Cadena, "T") & "," & DBSet(Mens, "T") & ")"
'
'                Conn.Execute SQL
'            End If
'        End If
'    End If
    
eComprobarRegistro:
    If Err.Number <> 0 Then
        ComprobarRegistro = False
    End If
End Function

            
Private Function InsertarLinea(cad As String) As Boolean
Dim NumLin As String
Dim codpro As String
Dim articulo As String
Dim Familia As String
Dim Precio As String
Dim ImpDes As String
Dim CodIva As String
Dim B As Boolean
Dim Codclave As String
Dim SQL As String

Dim Import As Currency

Dim CPostal As String

Dim Sql1 As String

Dim Mens As String
Dim numlinea As Long

Dim codsoc As String
Dim campo As String

    On Error GoTo EInsertarLinea

    InsertarLinea = True
    
    
    CargarVariables cad
    
    
    If Combo1(0).ListIndex = 0 Then
        CPostal = CPobla
    Else
        CPostal = ""
        CPostal = DevuelveDesdeBDNew(cAgro, "rcoope", "codposta", "codcoope", CPobla, "N")
    End If
    
    
    ' insertamos en la tabla de rhisfruta
    SQL = "insert into rhisfruta ("
    SQL = SQL & "`numalbar`,`fecalbar`,`codvarie`,`codsocio`,`codcampo`,`tipoentr`,"
    SQL = SQL & "`recolect`,`kilosbru`,`numcajon`,`kilosnet`,`imptrans`,`impacarr`,"
    SQL = SQL & "`imprecol`,`imppenal`,`impreso`,`impentrada`,`cobradosn`,`prestimado`,"
    SQL = SQL & "`codpobla`,`nromuestraalmz` ) VALUES ("
    SQL = SQL & DBSet(NumNota, "N") & ","
    SQL = SQL & DBSet(FechaEnt, "F") & ","
    SQL = SQL & DBSet(Variedad, "N") & ","
    SQL = SQL & DBSet(Socio, "N") & ","
    
    If vParamAplic.Cooperativa = 1 Then ' valsur no sabe el campo
        SQL = SQL & ValorNulo & ","
    Else ' caso de moixent
'[Monica]13/12/2011: en el caso de mogente tampoco podemos saber el campo por el poligono y parcela con lo cual suprimo esto
'        If CCur(Poligono) <> 0 And CCur(Parcela) <> 0 Then
'            Sql1 = "select codcampo from rcampos where codsocio = " & DBSet(Socio, "N")
'            Sql1 = Sql1 & " and codvarie = " & DBSet(Variedad, "N")
'            Sql1 = Sql1 & " and poligono = " & DBSet(Poligono, "N")
'            Sql1 = Sql1 & " and parcela = " & DBSet(Parcela, "N")
'
'            Campo = DevuelveValor(Sql1)
'
'            ' 18/11/2009 de momento si no encuentro el campo no lo meto
'            If Campo = 0 Then
'                SQL = SQL & ValorNulo & ","
'            Else
'                SQL = SQL & DBSet(Campo, "N") & ","
'            End If
'        Else
            SQL = SQL & ValorNulo & ","
'        End If
    End If
    
    SQL = SQL & "0,0,"
    SQL = SQL & DBSet(Bruto, "N") & ","
    SQL = SQL & "0," ' numero de cajones
    SQL = SQL & DBSet(Bruto, "N") & ","
    SQL = SQL & "0,0,0,0,0,0,0,0,"
    SQL = SQL & DBSet(CPostal, "T") & ","
    
    If vParamAplic.Cooperativa = 1 Then ' valsur no tiene nro.de muestra
        SQL = SQL & ValorNulo & ")"
    Else
        SQL = SQL & DBSet(NroMuestra, "N") & ")"
    End If
    
    conn.Execute SQL
    Exit Function
    
EInsertarLinea:
    If Err.Number <> 0 Then
        InsertarLinea = False
        MsgBox "Error en Insertar Linea " & Err.Description, vbExclamation
    End If
End Function
            
Private Function InsertarLineaABN(cad As String) As Boolean
Dim NumLin As String
Dim B As Boolean
Dim SQL As String

Dim Sql1 As String

Dim Mens As String
Dim numlinea As Long

Dim vError As Boolean
Dim vNota As Long
Dim cadena As String

    On Error GoTo EInsertarLinea

    InsertarLineaABN = True
    
    CargarVariables cad
    
    HayError = False
    
     ' comprobaciones para poder insertar la entrada
    cadena = ""
    ' comprobamos que me han puesto los datos de busqueda de parcela
    If Poligono = "" And Parcela = "" And Subparcela = "" Then
        Mens = "No hay datos de campo"
        SQL = "insert into tmpinformes (codusu, importe1,  " & _
              "importe2, nombre2, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(NumNota, "N") & ","
        SQL = SQL & "0," & DBSet(cadena, "T") & "," & DBSet(Mens, "T") & ")"

        conn.Execute SQL

    Else
        '[Monica]14/11/2017: hay veces que nos viene "" pongo comprobar cero para poligono y parcela
        cadena = Format(CCur(ComprobarCero(Poligono)), "0000") & "-" & Format(CCur(ComprobarCero(Parcela)), "0000") & "-" & Subparcela
    End If
    
' de momento lo quito pq hay una comprobacion previa que impide hacer nada si no existen los socios y variedades
'    'Comprobamos que el socio existe
'    Sql = ""
'    Sql = DevuelveDesdeBDNew(cAgro, "rsocios", "codsocio", "codsocio", Socio, "N")
'    If Sql = "" Then
'        Mens = "No existe el socio " & Socio
'        Sql = "insert into tmpinformes (codusu, importe1,  " & _
'              "importe2, nombre2, nombre1) values (" & _
'              vUsu.Codigo & "," & DBSet(NumNota, "N") & ","
'        Sql = Sql & "0," & DBSet(Cadena, "T") & "," & DBSet(Mens, "T") & ")"
'
'        conn.Execute Sql
'
'        HayError = True
'    End If
'
'    'Comprobamos que la variedad existe
'    Sql = ""
'    Sql = DevuelveDesdeBDNew(cAgro, "variedades", "codvarie", "codvarie", Variedad, "N")
'    If Sql = "" Then
'        Mens = "No existe la variedad " & Variedad
'        Sql = "insert into tmpinformes (codusu, importe1,  " & _
'              "importe2, nombre2, nombre1) values (" & _
'              vUsu.Codigo & "," & DBSet(NumNota, "N") & ","
'        Sql = Sql & "0," & DBSet(Cadena, "T") & "," & DBSet(Mens, "T") & ")"
'
'        conn.Execute Sql
'
'        HayError = True
'    End If
'
'    If HayError Then Exit Function
    
    '[Monica]02/11/2016: quito la condicion de que tengan valores
    ' comprobamos que el campo existe
'    If ComprobarCero(poligono) <> 0 And ComprobarCero(Parcela) <> 0  And ComprobarCero(Subparcela) <> 0 Then
    SQL = "select codcampo from rcampos where (1=1) "
    If ComprobarCero(Poligono) <> 0 And ComprobarCero(Parcela) <> 0 Then
        SQL = SQL & " and poligono = " & DBSet(Poligono, "N")
        SQL = SQL & " and parcela = " & DBSet(Parcela, "N")
        If ComprobarCero(Subparcela) <> 0 Then SQL = SQL & " and subparce = " & DBSet(Subparcela, "N")

        'si no existe el campo lo creamos
        If DevuelveValor(SQL) = 0 Then
            Set frmMens = New frmMensajes
            frmMens.cadena = Socio & "|" & Variedad & "|" & Poligono & "|" & Parcela & "|" & Subparcela & "|"
            frmMens.OpcionMensaje = 62
            frmMens.Show vbModal
            Set frmMens = Nothing
        Else
            campo = DevuelveValor(SQL)
        End If
    Else
        campo = 0
    End If
    
    If HayError Then Exit Function
    
    ' al nro de nota le sumo por delante la cooperativa
    Select Case Combo1(0).ListIndex
        Case 0 'bolbaite
            vNota = 3000000 + NumNota
        Case 1 ' anna
            vNota = 1000000 + NumNota
        Case 2 ' navarres
            vNota = 6000000 + NumNota
    End Select
    
    ' Comprobamos que la entrada no exista ya
    SQL = "select count(*) from rhisfruta where numalbar = " & DBSet(vNota, "N")
    If TotalRegistros(SQL) <> 0 Then
        HayError = True
    End If
    
    If HayError Then
        SQL = "update rhisfruta set fecalbar = " & DBSet(FechaEnt, "F")
        SQL = SQL & ", codvarie = " & DBSet(Variedad, "N")
        SQL = SQL & ", codsocio = " & DBSet(Socio, "N")
        SQL = SQL & ", codcampo = " & DBSet(campo, "N")
        SQL = SQL & ", kilosbru = " & DBSet(Bruto, "N")
        SQL = SQL & ", kilosnet = " & DBSet(Neto, "N")
        SQL = SQL & ", codpobla = " & DBSet(CPostal, "N")
        SQL = SQL & " where numalbar = " & DBSet(vNota, "N")
        
        conn.Execute SQL
        
        Exit Function
    End If
    
    
    ' insertamos en la tabla de rhisfruta
    SQL = "insert into rhisfruta ("
    SQL = SQL & "`numalbar`,`fecalbar`,`codvarie`,`codsocio`,`codcampo`,`tipoentr`,"
    SQL = SQL & "`recolect`,`kilosbru`,`numcajon`,`kilosnet`,`imptrans`,`impacarr`,"
    SQL = SQL & "`imprecol`,`imppenal`,`impreso`,`impentrada`,`cobradosn`,`prestimado`,"
    SQL = SQL & "`codpobla`,`nromuestraalmz` ) VALUES ("
    SQL = SQL & DBSet(vNota, "N") & ","
    SQL = SQL & DBSet(FechaEnt, "F") & ","
    SQL = SQL & DBSet(Variedad, "N") & ","
    SQL = SQL & DBSet(Socio, "N") & ","
    
    'campo
    SQL = SQL & DBSet(campo, "N") & ","
    
    SQL = SQL & "0,0,"
    SQL = SQL & DBSet(Bruto, "N") & ","
    SQL = SQL & "0," ' numero de cajones
    SQL = SQL & DBSet(Neto, "N") & ","
    SQL = SQL & "0,0,0,0,0,0,0,0,"
    SQL = SQL & DBSet(CPostal, "T") & ","
    SQL = SQL & ValorNulo & ")"
    
    conn.Execute SQL
    Exit Function
    
EInsertarLinea:
    If Err.Number <> 0 Then
        InsertarLineaABN = False
        MsgBox "Error en Insertar Linea " & Err.Description, vbExclamation
    End If
End Function
            
            
            
            
Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = CadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub InicializarTabla()
Dim SQL As String
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    
    conn.Execute SQL
End Sub



Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim I As Byte
    
    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For I = 0 To Combo1.Count - 1
        Combo1(I).Clear
    Next I
    
    If vParamAplic.Cooperativa = 1 Then ' caso de valsur tendran tres opciones
        'tipo de fichero
        Combo1(0).AddItem "Traspaso Entradas Bolbaite"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 0
        Combo1(0).AddItem "Traspaso Entradas Anna"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 1
        Combo1(0).AddItem "Traspaso Entradas Navarrés"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    Else ' en el caso de mogente tendran una sola opcion
        Combo1(0).AddItem "Traspaso Entradas"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    End If
    
    
End Sub


Private Sub CargarVariables(cad As String)
            
    
    Select Case vParamAplic.Cooperativa
    Case 1 ' CASO de VALSUR
        NumNota = ""
        FechaEnt = ""
        HoraEnt = ""
        Bruto = ""
        Variedad = ""
        Socio = ""
        Poligono = ""
        Parcela = ""
        Subparcela = ""
        Tara = ""
        
        NumNota = RecuperaValorNew(cad, ";", 1)
        FechaEnt = RecuperaValorNew(cad, ";", 2)
        HoraEnt = RecuperaValorNew(cad, ";", 3)
        Bruto = RecuperaValorNew(cad, ";", 5)
        Variedad = RecuperaValorNew(cad, ";", 6)
        Socio = RecuperaValorNew(cad, ";", 7)
        Poligono = RecuperaValorNew(cad, ";", 16)
        Parcela = RecuperaValorNew(cad, ";", 18)
        Subparcela = RecuperaValorNew(cad, ";", 19)
        Tara = RecuperaValorNew(cad, ";", 21)
    
        Neto = Round2(CCur(ComprobarCero(Bruto)) - CCur(ComprobarCero(Tara)), 0)
    
    Case 3 ' CASO de MOIXENT
        NumNota = ""
        PP = ""
        VV = ""
        Variedad = ""
        Socio = ""
        Bruto = ""
        CPobla = "46640"
        Poligono = ""
        Parcela = ""
        
        FechaEnt = ""
        
        NumNota = Mid(cad, 1, 5)
        NumNota = CStr(CCur(NumNota) + 9000000)
        
        Socio = Mid(cad, 6, 9)
        VV = Mid(cad, 53, 9)
        Variedad = Format(CCur(VV), "000000")
        Bruto = Mid(cad, 23, 7)
        Poligono = Mid(cad, 146, 8) '[Monica]07/12/2011: antes 136
        Parcela = Mid(cad, 154, 8)  '[Monica]07/12/2011: antes 144
        
        FechaEnt = Mid(cad, 15, 8)
        FechaEnt = Mid(FechaEnt, 7, 2) & "/" & Mid(FechaEnt, 5, 2) & "/" & Mid(FechaEnt, 1, 4)
    
        NroMuestra = Mid(cad, 228, 6) '[Monica]07/12/2011: antes no venia
    End Select
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim cadena As String
Dim Mens As String
Dim SQL As String

    If CadenaSeleccion = "" Then
        cadena = Format(ComprobarCero(Poligono), "0000") & "-" & Format(ComprobarCero(Parcela), "0000") & "-" & Subparcela
    
        Mens = "No se creó el Campo "
        SQL = "insert into tmpinformes (codusu, importe1,  " & _
              "importe2, nombre2, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Socio, "N") & ","
        SQL = SQL & "0," & DBSet(cadena, "T") & "," & DBSet(Mens, "T") & ")"
        
        conn.Execute SQL
        
        HayError = True
    
    
    Else
        campo = CadenaSeleccion
    End If
        
End Sub

Private Sub frmMens2_DatoSeleccionado(CadenaSeleccion As String)
    Continuar = (CadenaSeleccion = "OK")
End Sub




            
Private Function CompruebaSociosVariedades(cad As String) As Boolean
Dim NumLin As String
Dim B As Boolean
Dim SQL As String

Dim Sql1 As String

Dim Mens As String
Dim numlinea As Long

Dim vError As Boolean
Dim vNota As Long
Dim cadena As String


    CompruebaSociosVariedades = True
    
    CargarVariables cad
    
    
    'Comprobamos que el socio existe
    SQL = ""
    SQL = DevuelveDesdeBDNew(cAgro, "rsocios", "codsocio", "codsocio", Socio, "N")
    If SQL = "" Then
        SQL = "select count(*) from tmpinformes where codigo1 = " & Socio & " and codusu = " & vUsu.Codigo
        If TotalRegistros(SQL) = 0 Then
            SQL = "insert into tmpinformes (codusu,codigo1) values ("
            SQL = SQL & DBSet(vUsu.Codigo, "N") & "," & DBSet(Socio, "N") & ")"
            conn.Execute SQL
            
            SociosNoExisten = SociosNoExisten & Socio & ", "
        End If
    End If
    
    'Comprobamos que la variedad existe
    SQL = ""
    SQL = DevuelveDesdeBDNew(cAgro, "variedades", "codvarie", "codvarie", Variedad, "N")
    If SQL = "" Then
        SQL = "select count(*) from tmpinformes where importe1 = " & Variedad & " and codusu = " & vUsu.Codigo
        If TotalRegistros(SQL) = 0 Then
            SQL = "insert into tmpinformes (codusu,importe1) values ("
            SQL = SQL & DBSet(vUsu.Codigo, "N") & "," & DBSet(Variedad, "N") & ")"
            conn.Execute SQL
            
            VariedadesNoExisten = VariedadesNoExisten & Variedad & ", "
        End If
    End If


End Function

