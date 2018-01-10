VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAlmzTrasRendimiento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traspaso Rendimiento de Entradas ADV"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6480
   Icon            =   "frmAlmzTrasRendimiento.frx":0000
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
         Top             =   3780
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
         Top             =   3780
         Width           =   1065
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   210
         TabIndex        =   3
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
      Begin VB.Label Label2 
         Caption         =   "¿ Desea continuar ?"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2010
         TabIndex        =   7
         Top             =   1680
         Width           =   2835
      End
      Begin VB.Label Label1 
         Caption         =   "Proceso que actualiza el Rendimiento en las entradas de Almazara."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   765
         Left            =   270
         TabIndex        =   6
         Top             =   840
         Width           =   6105
      End
      Begin VB.Label lblProgres 
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   3120
         Width           =   6195
      End
      Begin VB.Label lblProgres 
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   3480
         Width           =   6195
      End
   End
End
Attribute VB_Name = "frmAlmzTrasRendimiento"
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


Dim Socio As String
Dim Rendimiento As String
Dim Variedad As String
Dim FechaDesde As String
Dim FechaHasta As String

Dim Muestra As String
Dim FechaRdto As String
Dim Humedad As String
Dim Acidez As String


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me

End Sub

Private Sub cmdAceptar_Click()

    If vParamAplic.Cooperativa = 3 Then
        ProcesoRendimientoMoixent
        Exit Sub
    Else
        ' En ppio este proceso lo gastaba Valsur
        ProcesoRendimiento
        Exit Sub
    End If
    
    
End Sub

Private Sub ProcesoRendimiento()
Dim SQL As String
Dim I As Byte
Dim cadWHERE As String
Dim B As Boolean
Dim NomFic As String
Dim cadena As String
Dim cadena1 As String

On Error GoTo eError

    Me.CommonDialog1.Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist


    Me.CommonDialog1.DefaultExt = "TXT"
    'cadena = Format(CDate(txtcodigo(0).Text), FormatoFecha)
    CommonDialog1.FilterIndex = 1
    Me.CommonDialog1.FileName = "rendim"
    
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
                    MsgBox "Hay errores en el Traspaso de Rendimiento. Debe corregirlos previamente.", vbExclamation
                    cadTitulo = "Errores de Traspaso Rendimiento"
                    cadNombreRPT = "rErroresTrasRdto.rpt"
                    
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


Private Sub ProcesoRendimientoMoixent()
Dim SQL As String
Dim I As Byte
Dim cadWHERE As String
Dim B As Boolean
Dim NomFic As String
Dim cadena As String
Dim cadena1 As String

On Error GoTo eError

    Me.CommonDialog1.Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist

    Me.CommonDialog1.DefaultExt = "PRN"
    'cadena = Format(CDate(txtcodigo(0).Text), FormatoFecha)
    CommonDialog1.FilterIndex = 1
    Me.CommonDialog1.FileName = "exportacion"
    
    Me.CommonDialog1.CancelError = True
    
    Me.CommonDialog1.ShowOpen
    
    If Me.CommonDialog1.FileName <> "" Then
        InicializarVbles
        InicializarTabla
            '========= PARAMETROS  =============================
        'Añadir el parametro de Empresa
        CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1

          
        If ProcesarFicheroMoixent2(Me.CommonDialog1.FileName) Then
                cadTabla = "tmpinformes"
                cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                
                SQL = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo
                
                If TotalRegistros(SQL) <> 0 Then
'                If HayRegParaInforme(cadTABLA, cadSelect) Then
                    MsgBox "Hay errores en el Traspaso de Rendimiento. Debe corregirlos previamente.", vbExclamation
                    cadTitulo = "Errores de Traspaso Rendimiento"
                    cadNombreRPT = "rErroresTrasRdto2.rpt"
                    
                    LlamarImprimir
                    Exit Sub
                Else
                    conn.BeginTrans
                    B = ProcesarFicheroMoixent(Me.CommonDialog1.FileName)
                End If
        '[Monica]13/01/2015: si hay error en la comprobacion que no haga nada
        Else
            conn.BeginTrans
                
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

    'IMAGES para busqueda
'     txtcodigo(0).Text = Format(Now - 1, "dd/mm/yyyy")

    '###Descomentar
'    CommitConexion
         
    If vParamAplic.Cooperativa = 3 Then
        Label1.Caption = Label1.Caption & vbCrLf & "Inserta en el mantenimiento de Rendimientos"
        DoEvents
    End If
         
         
    FrameCobrosVisible True, H, W
    Pb1.visible = False
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
'   Me.Width = w + 70
'   Me.Height = h + 350

'    cmdAceptar_Click

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
    DoEvents
        
    B = True
    While Not EOF(NF) And B
        I = I + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & I
        Me.Refresh
        DoEvents
        
        B = ActualizarLinea(cad)
        
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" Then
        B = ActualizarLinea(cad)
    End If
    
    ProcesarFichero = B
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

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
    DoEvents
    

    B = True

    While Not EOF(NF) And B
        I = I + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & I
        Me.Refresh
        DoEvents
        
        B = ComprobarRegistro(cad)
        
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" Then
        I = I + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & I
        Me.Refresh
        DoEvents
        
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
    If Not EsFechaOK(FechaDesde) Then
        Mens = "Fecha Desde incorrecta"
        SQL = "insert into tmpinformes (codusu, fecha1, fecha2, importe1, importe2, importe3, nombre1) " & _
              " values (" & _
              vUsu.Codigo & "," & DBSet(FechaDesde, "F") & "," & DBSet(FechaHasta, "F") & ","
        SQL = SQL & DBSet(Socio, "N") & "," & DBSet(Rendimiento, "N") & "," & _
              DBSet(Variedad, "N") & "," & DBSet(Mens, "T") & ")"
        
        conn.Execute SQL
    End If
    
    'Comprobamos fechas
    If Not EsFechaOK(FechaHasta) Then
        Mens = "Fecha Hasta incorrecta"
        SQL = "insert into tmpinformes (codusu, fecha1, fecha2, importe1, importe2, importe3, nombre1) " & _
              " values (" & _
              vUsu.Codigo & "," & DBSet(FechaDesde, "F") & "," & DBSet(FechaHasta, "F") & ","
        SQL = SQL & DBSet(Socio, "N") & "," & DBSet(Rendimiento, "N") & "," & _
              DBSet(Variedad, "N") & "," & DBSet(Mens, "T") & ")"
        
        conn.Execute SQL
    End If
    
    'Comprobamos que existe el socio
    SQL = ""
    SQL = DevuelveDesdeBDNew(cAgro, "rsocios", "nomsocio", "codsocio", Socio, "N")
    If SQL = "" Then
        Mens = "No existe el socio"
        SQL = "insert into tmpinformes (codusu, fecha1, fecha2, importe1, importe2, importe3, nombre1) " & _
              " values (" & _
              vUsu.Codigo & "," & DBSet(FechaDesde, "F") & "," & DBSet(FechaHasta, "F") & ","
        SQL = SQL & DBSet(Socio, "N") & "," & DBSet(Rendimiento, "N") & "," & _
              DBSet(Variedad, "N") & "," & DBSet(Mens, "T") & ")"
              
        conn.Execute SQL
    End If
    
    
    'Comprobamos que existe la variedad
    SQL = ""
    SQL = DevuelveDesdeBDNew(cAgro, "variedades", "codvarie", "codvarie", Variedad, "N")
    If SQL = "" Then
        Mens = "No existe la variedad"
        SQL = "insert into tmpinformes (codusu, fecha1, fecha2, importe1, importe2, importe3, nombre1) " & _
              " values (" & _
              vUsu.Codigo & "," & DBSet(FechaDesde, "F") & "," & DBSet(FechaHasta, "F") & ","
        SQL = SQL & DBSet(Socio, "N") & "," & DBSet(Rendimiento, "N") & "," & _
              DBSet(Variedad, "N") & "," & DBSet(Mens, "T") & ")"
              
        conn.Execute SQL
    End If
    
    
eComprobarRegistro:
    If Err.Number <> 0 Then
        ComprobarRegistro = False
    End If
End Function

            
Private Function ActualizarLinea(cad As String) As Boolean
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

    On Error GoTo EActualizarLinea

    ActualizarLinea = True
    
    CargarVariables cad
    
    ' actualizamos los registros de la tabla de rhisfruta
    SQL = "update rhisfruta set prestimado = " & DBSet(Rendimiento, "N")
    SQL = SQL & " where codsocio = " & DBSet(Socio, "N")
    SQL = SQL & " and codvarie = " & DBSet(Variedad, "N")
    SQL = SQL & " and fecalbar >= " & DBSet(FechaDesde, "F")
    SQL = SQL & " and fecalbar <= " & DBSet(FechaHasta, "F")
    SQL = SQL & " and codvarie in (select variedades.codvarie " ' la variedad sea del grupo de almazara
    SQL = SQL & " from variedades, productos where variedades.codprodu = productos.codprodu "
    SQL = SQL & " and productos.codgrupo = 5)"
     
    
    conn.Execute SQL
    Exit Function
    
EActualizarLinea:
    If Err.Number <> 0 Then
        ActualizarLinea = False
        MsgBox "Error en Actualizar Linea " & Err.Description, vbExclamation
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


Private Sub CargarVariables(cad As String)
            
        Socio = ""
        Rendimiento = ""
        Variedad = ""
        FechaDesde = ""
        FechaHasta = ""
        
        Socio = Mid(cad, 1, 6)
        Rendimiento = Mid(cad, 7, 6)
        Variedad = Mid(cad, 13, 3)
        FechaDesde = Mid(cad, 16, 10)
        FechaHasta = Mid(cad, 26, 10)

End Sub


Private Sub CargarVariablesMoixent(cad As String)
            
        Muestra = ""
        FechaRdto = ""
        
        Rendimiento = ""
        Acidez = ""
        Humedad = ""
        
        Muestra = Mid(cad, 1, 8)
        
        '[Monica]13/01/2014: el fichero ya no trae la fecha de muestra solo nro muestra y rdto
        FechaRdto = Format(Now, "dd/mm/yyyy") ' Mid(cad, 9, 19)
        
        Acidez = "0" ' Mid(cad, 46, 13)
        '[Monica]13/01/2014: lo modificamos
        Humedad = "0" ' Mid(cad, 39, 11)
        Rendimiento = Mid(cad, 15, 15) 'Mid(cad, 29, 10)
        

End Sub


Private Function ProcesarFicheroMoixent2(nomFich As String) As Boolean
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

    On Error GoTo eProcesarFicheroMoixent2
    
    ProcesarFicheroMoixent2 = False
    
    NF = FreeFile
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, cad
    I = 1
    
    lblProgres(0).Caption = "Insertando en Tabla temporal: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
    DoEvents
        
    Me.Pb1.Value = Me.Pb1.Value + Len(cad)
    lblProgres(1).Caption = "Linea " & I
    Me.Refresh
    DoEvents

    B = True

    While Not EOF(NF) And B
        Line Input #NF, cad
        
        I = I + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & I
        Me.Refresh
        DoEvents
        
        B = ComprobarRegistroMoixent(cad)
        
    Wend
    Close #NF
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

    ProcesarFicheroMoixent2 = B
    Exit Function

eProcesarFicheroMoixent2:
    ProcesarFicheroMoixent2 = False
End Function
            
Private Function ComprobarRegistroMoixent(cad As String) As Boolean
Dim SQL As String
Dim Mens As String
Dim cadena As String

    On Error GoTo eComprobarRegistro

    ComprobarRegistroMoixent = True

    CargarVariablesMoixent cad

    'Comprobamos la fecha de rendimiento
    If Not EsFechaOK(Mid(FechaRdto, 1, 10)) Then
        Mens = "Fecha Hasta incorrecta"
        SQL = "insert into tmpinformes (codusu, fecha1, fecha2, importe1, importe2, importe3, nombre1) " & _
              " values (" & _
              vUsu.Codigo & "," & DBSet(Mid(FechaRdto, 1, 10), "F") & "," & ValorNulo & ","
        SQL = SQL & DBSet(Muestra, "N") & "," & DBSet(Rendimiento, "N") & "," & _
              ValorNulo & "," & DBSet(Mens, "T") & ")"
        
        conn.Execute SQL
    End If
    
    '[Monica]13/01/2015: comprobamos que el numero de muestra es numerico
    ' Comprobamos que el nro de muestra sea numerico
    If Not IsNumeric(Muestra) Then
        Mens = "Muestra no numerico"
        SQL = "insert into tmpinformes (codusu, fecha1, fecha2, importe1, importe2, importe3, nombre1) " & _
              " values (" & _
              vUsu.Codigo & "," & DBSet(Mid(FechaRdto, 1, 10), "F") & "," & ValorNulo & ","
        SQL = SQL & DBSet(0, "N") & "," & DBSet(Rendimiento, "N") & "," & _
              ValorNulo & "," & DBSet(Mens, "T") & ")"
              
        conn.Execute SQL
    Else
    
        
        'Comprobamos que exista el nro de muestra
        SQL = ""
        SQL = DevuelveDesdeBDNew(cAgro, "rhisfruta", "numalbar", "nromuestraalmz", Muestra, "N")
        If SQL = "" Then
            Mens = "No existe Nro.Muestra"
            SQL = "insert into tmpinformes (codusu, fecha1, fecha2, importe1, importe2, importe3, nombre1) " & _
                  " values (" & _
                  vUsu.Codigo & "," & DBSet(Mid(FechaRdto, 1, 10), "F") & "," & ValorNulo & ","
            SQL = SQL & DBSet(Muestra, "N") & "," & DBSet(Rendimiento, "N") & "," & _
                  ValorNulo & "," & DBSet(Mens, "T") & ")"
                  
            conn.Execute SQL
        End If
        
        'Comprobamos que no exista el nro de muestra albaran fecha en la tabla de rendimiento
        SQL = ""
        SQL = "select count(*) from rrendim where nromuestra = " & DBSet(Muestra, "N") & " and fecha = " & DBSet(FechaRdto, "FH") ' DevuelveDesdeBDNew(cAgro, "rrendim", "nromuestra", "nromuestra", Muestra, "N", , "fecha", Mid(FechaRdto, 1, 10), "F")
        If TotalRegistros(SQL) > 0 Then
            Mens = "Existe Muestra en Rendimientos"
            SQL = "insert into tmpinformes (codusu, fecha1, fecha2, importe1, importe2, importe3, nombre1) " & _
                  " values (" & _
                  vUsu.Codigo & "," & DBSet(Mid(FechaRdto, 1, 10), "F") & "," & ValorNulo & ","
            SQL = SQL & DBSet(Muestra, "N") & "," & DBSet(Rendimiento, "N") & "," & _
                  ValorNulo & "," & DBSet(Mens, "T") & ")"
                  
            conn.Execute SQL
        End If
        
    End If
eComprobarRegistro:
    If Err.Number <> 0 Then
        ComprobarRegistroMoixent = False
    End If
End Function


Private Function ProcesarFicheroMoixent(nomFich As String) As Boolean
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

    ProcesarFicheroMoixent = False
    NF = FreeFile
    
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, cad
    I = 1
    
    lblProgres(0).Caption = "Procesando Fichero: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    DoEvents
    Me.Pb1.Value = 0
        
    Me.Pb1.Value = Me.Pb1.Value + Len(cad)
    lblProgres(1).Caption = "Linea " & I
    Me.Refresh
    DoEvents
        
    B = True
    While Not EOF(NF) And B
        Line Input #NF, cad
        
        I = I + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & I
        Me.Refresh
        DoEvents
        
        B = ActualizarLineaMoixent(cad)
        
    Wend
    Close #NF
    
    
    ProcesarFicheroMoixent = B
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

End Function



Private Function ActualizarLineaMoixent(cad As String) As Boolean
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

    On Error GoTo EActualizarLinea

    ActualizarLineaMoixent = True
    
    CargarVariablesMoixent cad
    
    ' actualizamos los registros de la tabla de rhisfruta
    SQL = "update rhisfruta set prestimado = " & DBSet(Rendimiento, "N")
    SQL = SQL & " where nromuestraalmz = " & DBSet(Muestra, "N")
     
    
    conn.Execute SQL
    
    ' insertamos en la tabla de rendimiento
    
    SQL = "insert ignore into rrendim (nromuestra, fecha, numalbar, acidez, humedad, rendimiento) "
    SQL = SQL & " select nromuestraalmz, " & DBSet(Mid(FechaRdto, 1, 10) & " " & Mid(FechaRdto, 12, 8), "FH") & ", numalbar, "
    SQL = SQL & DBSet(Acidez, "N") & "," & DBSet(Humedad, "N") & "," & DBSet(Rendimiento, "N") & " from rhisfruta "
    SQL = SQL & " where nromuestraalmz = " & DBSet(Muestra, "N")
    
    conn.Execute SQL
    
    Exit Function
    
EActualizarLinea:
    If Err.Number <> 0 Then
        ActualizarLineaMoixent = False
        MsgBox "Error en Actualizar Linea " & Err.Description, vbExclamation
    End If
End Function

