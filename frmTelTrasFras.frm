VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTelTrasFras 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Traspaso Facturas de Telefonia"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6825
   Icon            =   "frmTelTrasFras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   6825
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
      Height          =   4665
      Left            =   150
      TabIndex        =   2
      Top             =   120
      Width           =   6555
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   2295
         MaxLength       =   1
         TabIndex        =   6
         Tag             =   "Letra Serie Telefonia|T|S|||rparam|letraserietel|||"
         Top             =   1215
         Width           =   465
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   570
         Top             =   3390
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "doc"
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4905
         TabIndex        =   1
         Top             =   3780
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3690
         TabIndex        =   0
         Top             =   3780
         Width           =   975
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
      Begin VB.Label Label1 
         Caption         =   "Letra Serie Facturas"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   67
         Left            =   360
         TabIndex        =   7
         Top             =   1245
         Width           =   1650
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
Attribute VB_Name = "frmTelTrasFras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' PROGRAMA DE TRASPASO FACTURAS DE TELEFONIA PARA VALSUR
' basado en frmTrasPoste de gasolinera
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
Dim Cad As String
Dim cadTabla As String

Dim vContad As Long

Dim PrimeraVez As Boolean

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me

End Sub

Private Sub cmdAceptar_Click()
Dim Sql As String
Dim i As Byte
Dim cadWHERE As String
Dim b As Boolean
Dim NomFic As String
Dim cadena As String
Dim cadena1 As String

On Error GoTo eError


    If Not DatosOk Then Exit Sub
    
    Me.CommonDialog1.DefaultExt = "TXT"
    CommonDialog1.FilterIndex = 1
    Me.CommonDialog1.FileName = "*.txt"
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
                
            Sql = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo
                
            If TotalRegistros(Sql) <> 0 Then
                MsgBox "Hay errores en el Traspaso de Facturas Telefonia. Debe corregirlos previamente.", vbExclamation
                cadTitulo = "Errores de Traspaso de Facturas"
                cadNombreRPT = "rErroresTrasTel.rpt"
                LlamarImprimir
                Exit Sub
            Else
                conn.BeginTrans
                b = ProcesarFichero(Me.CommonDialog1.FileName)
            End If
        Else
            MsgBox "No se ha procesado ningún fichero. Revise.", vbExclamation
            Exit Sub
        End If
    Else
        MsgBox "No ha seleccionado ningún fichero", vbExclamation
        Exit Sub
    End If
             
             
eError:
    If Err.Number <> 0 Or Not b Then
        conn.RollbackTrans
        MsgBox "No se ha podido realizar el proceso. LLame a Ariadna.", vbExclamation
    Else
        conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        Pb1.visible = False
        lblProgres(0).Caption = ""
        lblProgres(1).Caption = ""
'        BorrarArchivo Me.CommonDialog1.FileName
        cmdCancel_Click
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco Text1(17)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, H, W
    Pb1.visible = False
'    Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Text1(17).Text = vParamAplic.LetraSerieTel
    
    Me.cmdCancel.Cancel = True
'   Me.Width = w + 70
'   Me.Height = h + 350
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DesBloqueoManual ("TRASPOST")
End Sub



Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    ' ***************** configurar els LostFocus dels camps de la capçalera *****************
    Select Case Index
        Case 17 'Letra de serie
            If Text1(Index).Text <> "" Then Text1(Index).Text = UCase(Text1(Index).Text)
            
    End Select
        ' ***************************************************************************
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
'            Case 17: KEYBusqueda KeyAscii, 3 'forma pago
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
'    imgBuscar_Click (indice)
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

 

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Sql As String
   b = True

    If vParamAplic.Seccionhorto = "" Then
        MsgBox "No se introducido la seccion de Horto en parámetros. Revise.", vbExclamation
        DatosOk = False
        Exit Function
    End If

   If Text1(17).Text = "" And b Then
        MsgBox "La letra de serie debe tener un valor. Reintroduzca.", vbExclamation
        b = False
        PonerFoco Text1(17)
    End If
 
    DatosOk = b
End Function


Private Function ProcesarFichero(nomFich As String) As Boolean
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
Dim NomFic As String

    ProcesarFichero = False
    NF = FreeFile
    
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, Cad
    i = 0
    
    lblProgres(0).Caption = "Procesando Fichero: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
        
    b = True
    While Not EOF(NF)
        i = i + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(Cad)
        lblProgres(1).Caption = "Linea " & i
        Me.Refresh
        
        Cad = Replace(Cad, Chr(9), "|")
        b = InsertarLinea(Cad)
        
        If b = False Then
            ProcesarFichero = False
            Exit Function
        End If
        
        Line Input #NF, Cad
    Wend
    Close #NF
    
    If Cad <> "" Then
        Cad = Replace(Cad, Chr(9), "|")
        b = InsertarLinea(Cad)

        If b = False Then
            ProcesarFichero = False
            Exit Function
        End If
    End If
    
    ProcesarFichero = b
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

End Function
                
Private Function ProcesarFichero2(nomFich As String) As Boolean
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

    On Error GoTo eProcesarFichero2
    
    ProcesarFichero2 = False
    
    NF = FreeFile
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, Cad
    i = 0
    
    lblProgres(0).Caption = "Insertando en Tabla temporal: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
    ' PROCESO DEL FICHERO VENTAS.TXT

    b = True

    While Not EOF(NF) And b
        i = i + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(Cad)
        lblProgres(1).Caption = "Linea " & i
        Me.Refresh
        Cad = Replace(Cad, Chr(9), "|")
        b = ComprobarRegistro(Cad)
        
        Line Input #NF, Cad
    Wend
    Close #NF
    
    If Cad <> "" Then
        i = i + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(Cad)
        lblProgres(1).Caption = "Linea " & i
        Me.Refresh
        Cad = Replace(Cad, Chr(9), "|")
        b = ComprobarRegistro(Cad)
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

Dim c_BaseImpo As Currency
Dim c_CuotaIva As Currency
Dim c_TotalFac As Currency

Dim Mens As String

Dim Fecha As String
Dim codsoc As String
Dim numfactu As String
Dim baseimpo As String
Dim CuotaIva As String
Dim TotalFac As String
Dim Digito As String

    On Error GoTo eComprobarRegistro

    ComprobarRegistro = True

    Fecha = RecuperaValor(Cad, 1)
    codsoc = RecuperaValor(Cad, 4)
    numfactu = RecuperaValor(Cad, 2)
    numfactu = Replace(numfactu, "-", "|") & "|"
    Digito = RecuperaValor(numfactu, 1)
    numfactu = RecuperaValor(numfactu, 4)
    numfactu = Format((CInt(Digito) * 1000000) + CLng(numfactu), "0000000")
    
    
    baseimpo = RecuperaValor(Cad, 6)
    CuotaIva = RecuperaValor(Cad, 7)
    TotalFac = RecuperaValor(Cad, 8)
    
    c_BaseImpo = CCur(TransformaPuntosComas(baseimpo))
    c_CuotaIva = CCur(TransformaPuntosComas(CuotaIva))
    c_TotalFac = CCur(TransformaPuntosComas(TotalFac))
    
    
    'Comprobamos fechas
    If Not EsFechaOK(Fecha) Then
        Mens = "Fecha incorrecta"
        Sql = "insert into tmpinformes (codusu, fecha1, importe1, importe2, importe3, " & _
              "importe4, importe5, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Fecha, "F") & _
              "," & DBSet(codsoc, "N") & "," & _
              DBSet(numfactu, "N") & "," & _
              DBSet(c_BaseImpo, "N") & "," & _
              DBSet(c_CuotaIva, "N") & "," & _
              DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ")"
        
        conn.Execute Sql
    End If
    
    
    'Comprobamos que el socio existe
    If codsoc <> "" Then
        Sql = ""
        Sql = DevuelveDesdeBDNew(cAgro, "rsocios", "codsocio", "codsocio", codsoc, "N")
        If Sql = "" Then
            Mens = "No existe el Socio"
            Sql = "insert into tmpinformes (codusu, fecha1, importe1, importe2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(Fecha, "F") & _
                  "," & DBSet(codsoc, "N") & "," & _
                  DBSet(numfactu, "N") & "," & _
                  DBSet(c_BaseImpo, "N") & "," & _
                  DBSet(c_CuotaIva, "N") & "," & _
                  DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ")"

            conn.Execute Sql
        End If
    End If
    
    ' comprobamos que el socio es de la seccion de horto
    If codsoc <> "" Then
        Sql = "select count(*) from rsocios_seccion where codsocio = " & DBSet(codsoc, "N")
        Sql = Sql & " and codsecci = " & vParamAplic.Seccionhorto
        If TotalRegistros(Sql) = 0 Then
            Mens = "No existe el Socio en Horto"
            Sql = "insert into tmpinformes (codusu, fecha1, importe1, importe2, importe3, " & _
                  "importe4, importe5, nombre1) values (" & _
                  vUsu.Codigo & "," & DBSet(Fecha, "F") & _
                  "," & DBSet(codsoc, "N") & "," & _
                  DBSet(numfactu, "N") & "," & _
                  DBSet(c_BaseImpo, "N") & "," & _
                  DBSet(c_CuotaIva, "N") & "," & _
                  DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ")"

            conn.Execute Sql
        End If
    End If
    
    
    'Comprobamos que la factura no existe
    Sql = "select count(*) from rtelmovil where numserie = " & DBSet(Text1(17).Text, "T")
    Sql = Sql & " and numfactu = " & DBSet(numfactu, "N")
    Sql = Sql & " and fecfactu = " & DBSet(Fecha, "F")
    
    If TotalRegistros(Sql) > 0 Then
        Mens = "Existe la factura"
        Sql = "insert into tmpinformes (codusu, fecha1, importe1, importe2, importe3, " & _
              "importe4, importe5, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Fecha, "F") & _
              "," & DBSet(codsoc, "N") & "," & _
              DBSet(numfactu, "N") & "," & _
              DBSet(c_BaseImpo, "N") & "," & _
              DBSet(c_CuotaIva, "N") & "," & _
              DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ")"
        
        conn.Execute Sql
    End If
    
    
    
eComprobarRegistro:
    If Err.Number <> 0 Then
        ComprobarRegistro = False
    End If
End Function
            
            
            
Private Function InsertarLinea(Cad As String) As Boolean
Dim c_BaseImpo As Currency
Dim c_CuotaIva As Currency
Dim c_TotalFac As Currency

Dim Mens As String

Dim Fecha As String
Dim codsoc As String
Dim numfactu As String
Dim baseimpo As String
Dim CuotaIva As String
Dim TotalFac As String
Dim Digito As String
Dim Sql As String


    On Error GoTo EInsertarLinea

    InsertarLinea = True

    Fecha = RecuperaValor(Cad, 1)
    codsoc = RecuperaValor(Cad, 4)
    numfactu = RecuperaValor(Cad, 2)
    numfactu = Replace(numfactu, "-", "|") & "|"
    Digito = RecuperaValor(numfactu, 1)
    numfactu = RecuperaValor(numfactu, 4)
    numfactu = Format((CInt(Digito) * 1000000) + CLng(numfactu), "0000000")
    
    
    baseimpo = RecuperaValor(Cad, 6)
    CuotaIva = RecuperaValor(Cad, 7)
    TotalFac = RecuperaValor(Cad, 8)
    
    c_BaseImpo = CCur(TransformaPuntosComas(baseimpo))
    c_CuotaIva = CCur(TransformaPuntosComas(CuotaIva))
    c_TotalFac = CCur(TransformaPuntosComas(TotalFac))
    
    
    ' insertamos en la tabla de telefonia
    
    Sql = "INSERT INTO rtelmovil (numserie, numfactu, fecfactu, codsocio, baseimpo, cuotaiva, " & _
          "totalfac, intconta) VALUES (" & DBSet(Text1(17).Text, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(Fecha, "F") & "," & _
           DBSet(codsoc, "N") & "," & DBSet(c_BaseImpo, "N") & "," & DBSet(c_CuotaIva, "N") & "," & _
           DBSet(c_TotalFac, "N") & ",0)"
    
    conn.Execute Sql
        
 
EInsertarLinea:
    If Err.Number <> 0 Then
        InsertarLinea = False
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
Dim Sql As String
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    
    conn.Execute Sql
End Sub


