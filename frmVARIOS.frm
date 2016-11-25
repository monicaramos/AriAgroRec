VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmVARIOS 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6690
   Icon            =   "frmVARIOS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   -30
      TabIndex        =   22
      Top             =   0
      Width           =   6645
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4200
         TabIndex        =   24
         Top             =   930
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5280
         TabIndex        =   23
         Top             =   945
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   225
         Left            =   390
         TabIndex        =   26
         Top             =   750
         Width           =   2625
      End
      Begin VB.Label Label3 
         Caption         =   "Recálculo de nuevas calidades de destrio y pixat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   330
         TabIndex        =   25
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.Frame FrameGeneraPreciosMasiva 
      Height          =   5310
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1980
         MaxLength       =   10
         TabIndex        =   1
         Top             =   2475
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1980
         MaxLength       =   10
         TabIndex        =   0
         Top             =   2070
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5190
         TabIndex        =   3
         Top             =   4545
         Width           =   975
      End
      Begin VB.CommandButton CmdAcepGen 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   4110
         TabIndex        =   2
         Top             =   4530
         Width           =   975
      End
      Begin VB.CommandButton Command10 
         Height          =   440
         Left            =   7860
         Picture         =   "frmVARIOS.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command9 
         Height          =   440
         Left            =   7860
         Picture         =   "frmVARIOS.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   630
         TabIndex        =   12
         Top             =   3990
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Albarán : "
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   660
         TabIndex        =   11
         Top             =   3660
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   1050
         TabIndex        =   10
         Top             =   2100
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   39
         Left            =   1050
         TabIndex        =   9
         Top             =   2520
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Albarán"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   50
         Left            =   600
         TabIndex        =   8
         Top             =   1650
         Width           =   540
      End
      Begin VB.Label Calculo 
         Caption         =   "Recálculo de Importes de Transporte"
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
         Left            =   630
         TabIndex        =   7
         Top             =   405
         Width           =   5775
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6075
      Top             =   4770
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameRecalculoDC 
      Height          =   2670
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   6615
      Begin VB.CommandButton Command3 
         Height          =   440
         Left            =   7860
         Picture         =   "frmVARIOS.frx":0620
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command2 
         Height          =   440
         Left            =   7860
         Picture         =   "frmVARIOS.frx":092A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton CmdAcepRec 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   1
         Left            =   3990
         TabIndex        =   15
         Top             =   1590
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel2 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5130
         TabIndex        =   14
         Top             =   1590
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio : "
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   660
         TabIndex        =   21
         Top             =   1920
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label Label1 
         Caption         =   "Recálculo de Dígito de Control"
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
         Left            =   630
         TabIndex        =   20
         Top             =   405
         Width           =   5775
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   7
         Left            =   600
         TabIndex        =   19
         Top             =   1650
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   630
         TabIndex        =   18
         Top             =   3990
         Width           =   45
      End
   End
End
Attribute VB_Name = "frmVARIOS"
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


Private WithEvents frmArtADV As frmADVArticulos 'articulos de adv
Attribute frmArtADV.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
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
Dim Tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String


Dim PrimeraVez As Boolean
Dim Contabilizada As Byte


Dim vSeccion As CSeccion

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me

End Sub


Private Sub cmdAceptar_Click()
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim KilosDestrio As Currency
Dim KilosPixat As Currency
Dim Maxima As Currency
Dim Kilos As Currency

Dim KilosDes As Currency

Dim v1 As Currency
Dim v2 As Currency
Dim TotalKilos As Long
Dim RstaGlobal As Boolean



    On Error GoTo EAceptar

    conn.BeginTrans

    Sql = "select * from rhisfruta_aux  order by numalbar "
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Label4.Caption = "Albaran " & Format(DBLet(Rs!numalbar, "N"), "0000000")
        
        KilosDestrio = Round(Rs!KilosNet * Rs!pordestrio / 100, 0)
        KilosPixat = Round(Rs!KilosNet * Rs!PORPIXAT / 100, 0)
        
        KilosDes = DevuelveValor("select kilosnet from rhisfruta_clasif where numalbar = " & DBSet(Rs!numalbar, "N") & " and codvarie = " & DBSet(Rs!CodVarie, "N") & " and codcalid = 5")
        
        
        Sql2 = "select * from rhisfruta_clasif where numalbar = " & DBSet(Rs!numalbar, "N") & " order by codvarie, codcalid "
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
        v1 = Rs!KilosNet - KilosDes
        v2 = Rs!KilosNet - KilosDestrio - KilosPixat
        TotalKilos = v2
        
        
        While Not Rs2.EOF
            RstaGlobal = False
            Select Case Rs2!codcalid
                Case 5
                    Kilos = KilosDestrio
                Case 16
                    Kilos = KilosPixat
                Case Else
                    Kilos = Round((Rs2!KilosNet * v2) / v1, 0)
                    
                    RstaGlobal = True
            End Select
            
           
            Sql3 = " where numalbar = " & DBSet(Rs!numalbar, "N")
            Sql3 = Sql3 & " and codvarie = " & DBSet(Rs!CodVarie, "N")
            Sql3 = Sql3 & " and codcalid = " & DBSet(Rs2!codcalid, "N")
            Rs2.MoveNext
            
            If Rs2.EOF Then
            
            Else
                If Rs2!codcalid = 16 Then
                    If TotalKilos < 0 Then TotalKilos = 0
                    Kilos = TotalKilos
                    TotalKilos = 0
                Else
                    If RstaGlobal Then
                        If TotalKilos - Kilos < 0 Then
                            Kilos = TotalKilos
                            TotalKilos = 0
                        Else
                            TotalKilos = TotalKilos - Kilos
                        End If
                    End If
                End If
            End If
            
            Sql3 = "update rhisfruta_clasif set kilosnet = " & DBSet(Kilos, "N") & Sql3
            conn.Execute Sql3
       
        Wend
        Set Rs2 = Nothing
            
        Rs.MoveNext
    Wend
    Set Rs = Nothing

    conn.CommitTrans
    MsgBox "Proceso realizado correctamente", vbExclamation
    Exit Sub

EAceptar:
    conn.RollbackTrans
    MuestraError Err.Number, "ERROR", Err.Description

End Sub

Private Sub CmdAcepRec_Click(Index As Integer)
Dim Sql As String
Dim Rs As ADODB.Recordset
                
Dim Ent As String ' Entidad
Dim Suc As String ' Oficina
Dim DC As String ' Digitos de control
Dim i, i2, i3, i4 As Integer
Dim NumCC As String ' Número de cuenta propiamente dicho
Dim CC As String
Dim cadResult As String
Dim DDCC As Integer
Dim NFich As Integer

    On Error GoTo eError


    conn.BeginTrans
    
    NFich = FreeFile
    Open App.Path & "\Resultados" & Format(Now, "yyyymmdd hhmmss") & ".txt" For Output As #NFich
     
    Label2(4).visible = True
    DoEvents
       
    
    Sql = "select codsocio,codbanco,codsucur,digcontr,cuentaba from rsocios where cuentaba <> '8888888888' order by codsocio "
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText


    cadResult = ""

    While Not Rs.EOF
        Label2(4).Caption = "Socio : " & Format(DBLet(Rs!Codsocio), "000000")
        DoEvents
    
        If IsNumeric(DBLet(Rs!CuentaBa)) And IsNumeric(DBLet(Rs!CodBanco)) And IsNumeric(DBLet(Rs!CodSucur)) Then
            
                If Not IsNumeric(DBLet(Rs!digcontr)) Then
                    DDCC = 0
                Else
                    DDCC = DBLet(Rs!digcontr)
                End If
                
            
                CC = Format(DBLet(Rs!CodBanco), "0000") & Format(DBLet(Rs!CodSucur), "0000") & Format(DDCC, "00") & Format(DBLet(Rs!CuentaBa), "0000000000")
                
                If Len(CC) = 20 Then  '-- Las cuentas deben contener 20 dígitos en total
                    If Not Comprueba_CC(CC) Then
                        
                        '-- Calculamos el primer dígito de control
                        i = Val(Mid(CC, 1, 1)) * 4
                        i = i + Val(Mid(CC, 2, 1)) * 8
                        i = i + Val(Mid(CC, 3, 1)) * 5
                        i = i + Val(Mid(CC, 4, 1)) * 10
                        i = i + Val(Mid(CC, 5, 1)) * 9
                        i = i + Val(Mid(CC, 6, 1)) * 7
                        i = i + Val(Mid(CC, 7, 1)) * 3
                        i = i + Val(Mid(CC, 8, 1)) * 6
                        i2 = Int(i / 11)
                        i3 = i - (i2 * 11)
                        i4 = 11 - i3
                        Select Case i4
                            Case 11
                                i4 = 0
                            Case 10
                                i4 = 1
                        End Select
                        
                        DC = i4
                        
                        '-- Calculamos el segundo dígito de control
                        i = Val(Mid(CC, 11, 1)) * 1
                        i = i + Val(Mid(CC, 12, 1)) * 2
                        i = i + Val(Mid(CC, 13, 1)) * 4
                        i = i + Val(Mid(CC, 14, 1)) * 8
                        i = i + Val(Mid(CC, 15, 1)) * 5
                        i = i + Val(Mid(CC, 16, 1)) * 10
                        i = i + Val(Mid(CC, 17, 1)) * 9
                        i = i + Val(Mid(CC, 18, 1)) * 7
                        i = i + Val(Mid(CC, 19, 1)) * 3
                        i = i + Val(Mid(CC, 20, 1)) * 6
                        i2 = Int(i / 11)
                        i3 = i - (i2 * 11)
                        i4 = 11 - i3
                        Select Case i4
                            Case 11
                                i4 = 0
                            Case 10
                                i4 = 1
                        End Select
                        
                        DC = DC & i4
                        
                        Sql = "update rsocios set digcontr = " & DBSet(DC, "N") & " where codsocio = " & DBSet(Rs!Codsocio, "N")
                        conn.Execute Sql
                    
                        cadResult = cadResult & DBLet(Rs!Codsocio, "N") & "-"
                    
                        Print #NFich, "Socio : " & Format(DBLet(Rs!Codsocio), "000000") & " DC Anterior " & DBLet(Rs!digcontr) & " - Nuevo " & DC
                    
                    End If
                End If
            
        End If
    
        Rs.MoveNext
    Wend

    If cadResult <> "" Then
        
        cadResult = Mid(cadResult, 1, Len(cadResult) - 1)
    
    
        Set Rs = Nothing
        
        conn.CommitTrans
        
        MsgBox "Proceso realizado correctamente. Se han modificado los DC de los socios: " & vbCrLf & vbCrLf & cadResult, vbExclamation
    
    Else
        conn.CommitTrans
        MsgBox "No se han encontrado registros erroneos.", vbExclamation
        
    End If
    
    Label2(4).visible = False
    DoEvents
    
    Close #NFich
    Exit Sub

eError:
    conn.RollbackTrans
    MuestraError Err.Number, "Error en socio", Err.Description
    

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub CmdCancel2_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtcodigo(1)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

    Tabla = "rhisfruta"
    
    ' Necesitamos la conexion a la contabilidad de la seccion de adv
    ' para sacar los porcentajes de iva de los articulos y calcular
    ' los datos de la factura
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
'    Me.cmdCancel(indFrame).Cancel = True
'    Me.Width = w + 70
'    Me.Height = h + 350
End Sub


Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " {variedades.codvarie} in (" & CadenaSeleccion & ")"
        Sql2 = " {variedades.codvarie} in [" & CadenaSeleccion & "]"
    Else
        Sql = " {variedades.codvarie} = -1 "
    End If
    If Not AnyadirAFormula(cadSelect, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub


Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0 'VARIEDADES
            AbrirFrmArticuloADV (Index)
    
    End Select
    PonerFoco txtcodigo(indCodigo)
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
            Case 0: KEYBusqueda KeyAscii, 0 'variedad desde
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub


'Private Sub txtCodigo_LostFocus(Index As Integer)
'Dim Cad As String, cadTipo As String 'tipo cliente
'
'    'Quitar espacios en blanco por los lados
'    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
''    If txtCodigo(Index).Text = "" Then Exit Sub
'
'    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
'    'mostrar mensajes ni hacer nada
'    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
'
'    Select Case Index
'        Case 0 'VARIEDADES
'            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedades", "nomvarie", "codvarie", "N")
'            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
'
'        Case 1, 2 'FECHAS
'            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
'
'        Case 3, 4 'PRECIOS
'            PonerFormatoDecimal txtCodigo(Index), 8
'
'    End Select
'End Sub


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

Private Sub AbrirFrmArticuloADV(indice As Integer)
    indCodigo = indice
    Set frmArtADV = New frmADVArticulos
    frmArtADV.DatosADevolverBusqueda = "0|1|"
    frmArtADV.Show vbModal
    Set frmArtADV = Nothing
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


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Sql As String
Dim Sql2 As String
Dim vClien As cSocio
' añadido
Dim Mens As String
Dim numfactu As String
Dim numser As String
Dim Fecha As Date
Dim Cad As String

    b = True
    
    If b Then
        If (txtcodigo(1).Text = "" Or txtcodigo(2).Text = "") Then
            MsgBox "El rango de albaranes debe de tener un valor. Reintroduzca.", vbExclamation
            b = False
        End If
    End If
    
    DatosOk = b

End Function




Private Function ActualizarRegistros(cTabla As String, cWhere As String) As Boolean
'Actualizar la marca de impreso
Dim Sql As String

    On Error GoTo eActualizarRegistros

    ActualizarRegistros = False
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "update " & QuitarCaracterACadena(cTabla, "_1") & " set impreso = 1 "
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    conn.Execute Sql
    
    ActualizarRegistros = True
    Exit Function
    
eActualizarRegistros:
    MuestraError Err.Number, "Actualizando registros", Err.Description
End Function



Public Function GeneraRegistros(vDesde As String, vHasta As String) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' ariges.scafac --> conta.cabfact
' ariges.slifac --> conta.linfact
'Actualizar la tabla ariges.scafac.inconta=1 para indicar que ya esta contabilizada
Dim b As Boolean
Dim cadMen As String
Dim Sql As String
Dim NumF As Currency
Dim vFactADV As CFacturaADV
Dim vSocio As cSocio
Dim Rs As ADODB.Recordset
Dim cadErr As String


Dim Desde As Long
Dim Hasta As Long
Dim i As Long

    On Error GoTo EContab

    conn.BeginTrans
    
    b = True
    
    Desde = CLng(vDesde)
    Hasta = CLng(vHasta)
    
    Sql = "select numalbar from rhisfruta where tipoentr <> 1 and numalbar >=  " & DBSet(Desde, "N")
    Sql = Sql & " and numalbar <= " & DBSet(Hasta, "N")
    Sql = Sql & " order by numalbar "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF And b
        Label2(2).Caption = Rs!numalbar
        DoEvents
    
        b = CalculoGastosTransporte(Rs!numalbar, cadErr)
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
'antes
'    I = vDesde
'    While I <= vHasta And b
'
'            If DevuelveValor("select tipoentr from rhisfruta where numalbar =" & I) <> 1 Then
'                '[Monica]27/04/2010 Calculo de costes de transporte, si es por tarifas y la entrada no es de venta campo
'                b = CalculoGastosTransporte(I, cadErr)
'            End If
'        I = I + 1
'    Wend
    
EContab:
    
    If Err.Number <> 0 Then
        b = False
        MuestraError Err.Number, "Modificando Registros", Err.Description & " " & cadErr
    End If
    If b Then
        conn.CommitTrans
        GeneraRegistros = True
    Else
        conn.RollbackTrans
        GeneraRegistros = False
    End If
End Function

Private Function CalculoGastosTransporte(Albaran As Long, cadErr As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim PrecTarifaAlm As Currency
Dim PrecTarifaAlm2 As Currency
Dim PrecTarifa As Currency
Dim ImpTrans As Currency
Dim TotImpTrans As Currency
Dim ImpGastoSocio As Currency
Dim NumF As String

On Error GoTo EInsertar
    
    ' calculamos los gastos de transporte para el socio y el importe de gastos de transporte de rhisfruta
    ' a partir de la entradas que ya hemos grabado en la rhisfruta_entradas

    cadErr = ""

    If vParamAplic.Cooperativa <> 2 And vParamAplic.Cooperativa <> 16 Then
        Sql = "select numnotac, rhisfruta_entradas.codtarif, rtarifatra.tipotarifa, sum(kilosnet) as kilos "
    Else
        Sql = "select numnotac, rhisfruta_entradas.codtarif, rtarifatra.tipotarifa, sum(kilostra) as kilos "
    End If
    Sql = Sql & " from rhisfruta_entradas, rtarifatra where numalbar = " & DBSet(Albaran, "N")
    Sql = Sql & " and rtarifatra.tipotarifa <> 2 " 'las tarifas que buscamos son del tipo 1 o 2 (no sin asignar)
    Sql = Sql & " and rtarifatra.codtarif = rhisfruta_entradas.codtarif "
    Sql = Sql & " group by 1, 2, 3 "
    Sql = Sql & " order by 1, 2, 3 "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    PrecTarifaAlm = DevuelveValor("select preciokg from rtarifatra where codtarif = " & vParamAplic.TarifaTRA)
    PrecTarifaAlm2 = DevuelveValor("select preciokg from rtarifatra where codtarif = " & vParamAplic.TarifaTRA2)
    
    ImpTrans = 0
    TotImpTrans = 0
    ImpGastoSocio = 0
    
    While Not Rs.EOF
        PrecTarifa = DevuelveValor("select preciokg from rtarifatra where codtarif = " & DBSet(Rs!Codtarif, "N"))
        
        ImpTrans = Round2(PrecTarifa * DBLet(Rs!Kilos, "N"), 2)
        
        TotImpTrans = TotImpTrans + ImpTrans
        If vParamAplic.Cooperativa <> 2 And vParamAplic.Cooperativa <> 16 Then
            Sql = "update rhisfruta_entradas set imptrans = " & DBSet(ImpTrans, "N")
            Sql = Sql & " where numalbar = " & DBSet(Albaran, "N") & " and numnotac = " & DBSet(Rs!Numnotac, "N")
            
            conn.Execute Sql
        End If
        If DBLet(Rs!tipotarifa, "N") = 0 Then
            ImpGastoSocio = ImpGastoSocio + Round2((DBLet(Rs!Kilos, "N") * (PrecTarifa - PrecTarifaAlm)), 2)
        Else
            ImpGastoSocio = ImpGastoSocio + Round2((DBLet(Rs!Kilos, "N") * (PrecTarifa - PrecTarifaAlm2)), 2)
        End If
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    If vParamAplic.Cooperativa <> 2 And vParamAplic.Cooperativa <> 16 Then
        ' actualizamos cabecera
        Sql = "update rhisfruta set imptrans = " & DBSet(TotImpTrans, "N")
        Sql = Sql & " where numalbar = " & DBSet(Albaran, "N")
        
        conn.Execute Sql
    End If
    
    If ImpGastoSocio <> 0 Then
        ' si existe registro en la tabla rhisfruta_gastos de concepto codgastotra actualizamos el importe
        Sql = "select count(*) from rhisfruta_gastos where numalbar = " & DBSet(Albaran, "N")
        Sql = Sql & " and codgasto = " & DBSet(vParamAplic.CodGastoTRA, "N")
        
        If TotalRegistros(Sql) = 0 Then
        
            NumF = ""
            NumF = SugerirCodigoSiguienteStr("rhisfruta_gastos", "numlinea", "numalbar = " & DBSet(Albaran, "N"))
            ' grabamos un registro en con los gastos del cliente
            Sql = "insert into rhisfruta_gastos (numalbar, numlinea, codgasto, importe) values (" & DBSet(Albaran, "N") & ","
            Sql = Sql & DBSet(NumF, "N") & "," & DBSet(vParamAplic.CodGastoTRA, "N") & "," & DBSet(ImpGastoSocio, "N") & ")"
            
            conn.Execute Sql
            
        Else
        
            ' acualizamos el registro que hay
            Sql = "update rhisfruta_gastos set importe = " & DBSet(ImpGastoSocio, "N")
            Sql = Sql & " where numalbar = " & DBSet(Albaran, "N")
            Sql = Sql & " and codgasto = " & DBSet(vParamAplic.CodGastoTRA, "N")
            
            conn.Execute Sql
        
        End If
    End If
EInsertar:
    If Err.Number <> 0 Then
        CalculoGastosTransporte = False
        cadErr = Err.Description
    Else
        CalculoGastosTransporte = True
    End If
End Function


