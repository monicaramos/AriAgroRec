VERSION 5.00
Begin VB.Form frmImpAridoc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar datos a AriDoc"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5940
   Icon            =   "frmImpAridoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   3810
      Width           =   1695
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   3810
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Carpeta de destino: "
      Height          =   1215
      Left            =   135
      TabIndex        =   3
      Top             =   360
      Width           =   5655
      Begin VB.TextBox txtCarp 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   5
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox txtCarp 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   495
         TabIndex        =   0
         Top             =   585
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Código:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1650
      Left            =   150
      TabIndex        =   7
      Top             =   1650
      Width           =   5640
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Tag             =   "Tipo|N|N|||straba|codsecci||N|"
         Top             =   240
         Width           =   1350
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   9
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1170
         Width           =   1050
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   825
         Width           =   1050
      End
      Begin VB.Label Label8 
         Caption         =   "Tipo Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Left            =   270
         TabIndex        =   14
         Top             =   270
         Width           =   1065
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   15
         Left            =   1590
         TabIndex        =   12
         Top             =   810
         Width           =   465
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   14
         Left            =   1590
         TabIndex        =   11
         Top             =   1170
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   2250
         Picture         =   "frmImpAridoc.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   810
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   2250
         Picture         =   "frmImpAridoc.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   1170
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   270
         TabIndex        =   10
         Top             =   690
         Width           =   1185
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1590
      Left            =   135
      TabIndex        =   15
      Top             =   1710
      Width           =   5640
      Begin VB.CheckBox Check1 
         Caption         =   "Sobre Horas Productivas"
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   20
         Top             =   1080
         Width           =   2130
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1695
         MaxLength       =   10
         TabIndex        =   17
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   405
         Width           =   1050
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   4005
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Tag             =   "Tipo|N|N|||straba|codsecci||N|"
         Top             =   405
         Width           =   1350
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Recibo"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   19
         Top             =   435
         Width           =   1185
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1425
         Picture         =   "frmImpAridoc.frx":0122
         ToolTipText     =   "Buscar fecha"
         Top             =   435
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Sección "
         ForeColor       =   &H00972E0B&
         Height          =   255
         Left            =   3060
         TabIndex        =   18
         Top             =   435
         Width           =   615
      End
   End
   Begin VB.Label lblInf 
      Alignment       =   2  'Center
      Caption         =   "Información del proceso"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3375
      Width           =   5295
   End
End
Attribute VB_Name = "frmImpAridoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Tipo As Byte
    'Tipo:  0 Impresion de facturas
    '       1 Impresion de facturas de adv
    '       2 Impresion de facturas de almazara
    '       3 Impresion de facturas de bodega
    '       4 Impresion de recibos nóminas
    '       5 Impresion de facturas de transporte
    '       6 Impresion de facturas de pozos

Dim DesdeFecha As Date
Dim Hastafecha As Date
Dim frmVis As frmVisReport
Dim impor As ArdImportador

Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1

Private Sub cmdAceptar_Click()
    If Not DatosOk() Then Exit Sub
    '-- Cargar facturas  entre las fechas seleccionadas
    Select Case Tipo
        Case 0 ' facturas
            CargaFacturas Combo1(1).ListIndex, DesdeFecha, Hastafecha
            MsgBox "Proceso finalizado", vbInformation
        Case 1 ' facturas de adv
            CargaFacturasADV DesdeFecha, Hastafecha
            MsgBox "Proceso finalizado", vbInformation
            
        Case 2 ' facturas de almazara
            CargaFacturasBod DesdeFecha, Hastafecha, 0
            MsgBox "Proceso finalizado", vbInformation
        
        Case 3 ' facturas de bodega
            CargaFacturasBod DesdeFecha, Hastafecha, 1
            MsgBox "Proceso finalizado", vbInformation
        
        Case 4 ' recibos nómina
            CargaRecibos DesdeFecha, Hastafecha
            MsgBox "Proceso finalizado", vbInformation
        
        Case 5 ' facturas de transporte
            CargaFacturasTransporte Combo1(1).ListIndex, DesdeFecha, Hastafecha
            MsgBox "Proceso finalizado", vbInformation
        
        Case 6 ' facturas de pozos
            CargaFacturasPozos Combo1(1).ListIndex, DesdeFecha, Hastafecha
            MsgBox "Proceso finalizado", vbInformation
        
        
    End Select
    cmdSalir_Click
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Function DatosOk() As Boolean
    DesdeFecha = CDate(txtCodigo(0).Text)
    Hastafecha = CDate(txtCodigo(1).Text)
    If DesdeFecha > Hastafecha Then
        MsgBox "La fecha desde debe ser menor que la fecha hasta", vbInformation
        Exit Function
    End If
    If txtCarp(1) = "" Then
        MsgBox "Debe seleccionar una carpeta de importación.", vbInformation
        Exit Function
    End If
    DatosOk = True
End Function




Private Sub Combo1_LostFocus(Index As Integer)
   If Index = 1 Then
    Select Case Tipo
        Case 0
            Select Case Combo1(1).ListIndex
                Case 0 ' anticipos
                    Me.txtCarp(0).Text = vParamAplic.CarpetaAnt
                Case 1 ' liquidaciones
                    Me.txtCarp(0).Text = vParamAplic.CarpetaLiq
            End Select
        Case 1
            Combo1(1).ListIndex = 3
            Combo1(1).Enabled = False
            Me.txtCarp(0).Text = vParamAplic.CarpetaADV
        Case 2
            Combo1(1).ListIndex = 4
            Combo1(1).Enabled = False
            Me.txtCarp(0).Text = vParamAplic.CarpetaAlmz
        Case 3
            Combo1(1).ListIndex = 5
            Combo1(1).Enabled = False
            Me.txtCarp(0).Text = vParamAplic.CarpetaBOD
        Case 5
            Combo1(1).ListIndex = 7
            Combo1(1).Enabled = False
            Me.txtCarp(0).Text = vParamAplic.CarpetaTra
            
    End Select
   Else
        Select Case Combo1(0).ListIndex
            Case 0
                Me.txtCarp(0).Text = vParamAplic.CarpetaRecCampo
            Case 1
                Me.txtCarp(0).Text = vParamAplic.CarpetaRecAlmacen
        End Select
   End If
    txtCarp_LostFocus (0)

End Sub

Private Sub Form_Load()
    txtCodigo(0).Text = Date
    txtCodigo(1).Text = Date
    Set impor = New ArdImportador
    
    Set ardDB = New BaseDatos
    ardDB.Tipo = "MYSQL"
    ardDB.abrir "Aridoc", "root", "aritel"
    
    CargaCombo
    
    Frame2.Enabled = (Tipo = 4)
    Frame2.visible = (Tipo = 4)
    
    Frame3.Enabled = (Tipo <> 4)
    Frame3.visible = (Tipo <> 4)
    Combo1(0).ListIndex = 1
    Check1(1).Enabled = False
    Check1(1).visible = False
        
    Select Case Tipo
        Case 0, 6
            Combo1(1).ListIndex = 0
            Combo1(1).Enabled = True
        Case 1
            Combo1(1).ListIndex = 3
            Combo1(1).Enabled = False
        Case 2
            Combo1(1).ListIndex = 4
            Combo1(1).Enabled = False
        Case 3
            Combo1(1).ListIndex = 5
            Combo1(1).Enabled = False
        Case 4
            Combo1(1).ListIndex = 6
            Combo1(1).Enabled = False
        Case 5
            Combo1(1).ListIndex = 7
            Combo1(1).Enabled = False
    End Select
'    PosicionarCombo Me.Combo1(1), Combo1(1).ListIndex
    Select Case Combo1(1).ListIndex
        Case 0:
            Me.txtCarp(0).Text = vParamAplic.CarpetaAnt
        Case 1:
            Me.txtCarp(0).Text = vParamAplic.CarpetaLiq
        Case 3:
            Me.txtCarp(0).Text = vParamAplic.CarpetaADV
        Case 4:
            Me.txtCarp(0).Text = vParamAplic.CarpetaAlmz
        Case 5:
            Me.txtCarp(0).Text = vParamAplic.CarpetaBOD
        Case 6:
            Select Case Combo1(0).ListIndex
                Case 0
                    Me.txtCarp(0).Text = vParamAplic.CarpetaRecCampo
                Case 1
                    Me.txtCarp(0).Text = vParamAplic.CarpetaRecAlmacen
            End Select
            Check1(1).Enabled = True
            Check1(1).visible = True
        Case 7
            Me.txtCarp(0).Text = vParamAplic.CarpetaTra
    End Select
    txtCarp_LostFocus (0)
End Sub

Private Sub imgFec_Click(Index As Integer)
'FEchas
    Dim esq, dalt As Long
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
       
    ' es desplega dalt i cap a la esquerra
    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + 420 + 30

    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
    imgFec(0).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index).Text <> "" Then frmC.NovaData = txtCodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(0).Tag))
    ' ***************************
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub txtCarp_GotFocus(Index As Integer)
    ConseguirFoco txtCarp(Index), 3
End Sub

Private Sub txtCarp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCarp_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCarp_LostFocus(Index As Integer)
Dim Cad As String
    If Index = 0 Then
        If txtCarp(0) <> "" Then
            'txtCarp(1) = impor.nombreCarpeta(CLng(txtCarp(0))) '  CargaPath(CLng(txtCarp(0))) 'impor.nombreCarpeta(CLng(txtCarp(0)))
            Cad = CargaPath(txtCarp(Index))
            txtCarp(1).Text = Mid(Cad, 2, Len(Cad))
        End If
    End If
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'14/02/2007 antes
'    KEYpress KeyAscii
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYFecha KeyAscii, 0 'fecha desde
            Case 1: KEYFecha KeyAscii, 1 'fecha hasta
            Case 2: KEYFecha KeyAscii, 1 'fecha de recibo
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 0, 1, 2 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
    End Select
End Sub


Private Sub CargaFacturas(TipoFact As Byte, DFecha As Date, HFecha As Date)
' TipoFact: 0 = Anticipos
'           1 = Liquidacion

    Dim db As BaseDatos
    Dim SQL As String
    Dim RS As ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Dim I As Long
    Dim FicheroPDF As String
    Dim c1 As String
    Dim c2 As String
    Dim c3 As String
    Dim c4 As String
    Dim f1 As Date
    Dim f3 As Date
    Dim i1 As Currency
    Dim fr As frmVisReport
    
    Dim Variedad As String
    
    Dim TipoFact1 As Byte
    Dim Gastos As Currency
    
On Error GoTo err_CargaFacturas
    
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim numParam As Byte
Dim cadparam As String

    Set db = New BaseDatos
    db.Tipo = "MYSQL"
    
    db.abrir_MYSQL vConfig.SERVER, vUsu.CadenaConexion, "root", "aritel"


'    db.abrir "accArigasol", "", ""
    SQL = "select rfactsoc.*, stipom.letraser " & _
            " from rfactsoc, usuarios.stipom where rfactsoc.fecfactu >= " & db.Fecha(CDate(txtCodigo(0).Text)) & _
            " and rfactsoc.fecfactu <= " & db.Fecha(CDate(txtCodigo(1).Text)) & _
            " and rfactsoc.codtipom = stipom.codtipom " & _
            " and rfactsoc.pasaridoc = 0"
            
    Select Case TipoFact
        Case 0 ' anticipos
            '[Monica]01/04/2011 añadidas 7,9 anticipos de almazara y de bodega
            SQL = SQL & " and stipom.tipodocu in (1,3,7,9)"
        Case 1 ' liquidaciones
            '[Monica]01/04/2011 añadidas 8,10 liquidaciones de almazara y de bodega
            SQL = SQL & " and stipom.tipodocu in (2,4,5,6,8,10)" ' [Monica]26/04/2010: añadidas 5 y 6
        Case 2 ' rectificativas
            SQL = SQL & " and stipom.tipodocu = 11 "
    End Select
            
    Set RS = db.cursor(SQL)
    
    If Not RS.EOF Then
        RS.MoveFirst
        While Not RS.EOF
            I = I + 1
            lblInf.Caption = "Procesando registro " & CStr(I)
            lblInf.Refresh
            '-- Creamos el pdf
            FicheroPDF = App.Path & "\ExpAriDoc.pdf"
            
'18/02/2010: lo quito para que prueben
'            If Not IntentaMatar(FicheroPDF) Then Err.Raise 53
            
            
            Set fr = New frmVisReport
            
            '++monica: seleccionamos que rpt se ha de ejecutar

'            cadParam = "pEmpresa=""AriagroRec""|"
            cadparam = ""
            numParam = 1
            
            If DBLet(RS!CodTipom, "T") = "FLI" Then
                indRPT = 38
            Else
                
                '[Monica]01/04/2011: anticipos y liquidaciones de almazara y bodega
                '                   antes no estaba puesta esta condicion
                If DBLet(Mid(RS!CodTipom, 2, 2), "T") = "LZ" Or _
                   DBLet(Mid(RS!CodTipom, 2, 2), "T") = "NZ" Or _
                   DBLet(Mid(RS!CodTipom, 2, 2), "T") = "LB" Or _
                   DBLet(Mid(RS!CodTipom, 2, 2), "T") = "NB" Then
                    indRPT = 42
                Else
                    indRPT = 23 'Impresion de Factura
                End If
            End If
            If Not PonerParamRPT(indRPT, cadparam, numParam, nomDocu) Then Exit Sub
            '++
            fr.NumeroParametros = numParam
            fr.OtrosParametros = cadparam
            fr.ConSubInforme = True
            fr.Informe = App.Path & "\Informes\" & nomDocu
            fr.FormulaSeleccion = "{rfactsoc.codtipom} = '" & RS!CodTipom & "' and " & _
                                  "{rfactsoc.numfactu} =" & CStr(RS!numfactu) & " and " & _
                                  "{rfactsoc.fecfactu} = Date(" & Format(RS!fecfactu, "yyyy") & _
                                                        "," & Format(RS!fecfactu, "mm") & _
                                                        "," & Format(RS!fecfactu, "dd") & ")"
            fr.FicheroPDF = FicheroPDF
            Load fr 'trabaja sin mostrar el formulario
            Screen.MousePointer = vbDefault

            SQL = "select * from rsocios where codsocio = " & db.numero(RS!Codsocio)
            Set Rs2 = db.cursor(SQL)
            
            '++monica: si hay mas una variedad meter agrupada
            '          si solo hay una meter la variedad
            
            SQL = "select count(distinct codvarie) from rfactsoc_variedad where "
            SQL = SQL & " codtipom = " & DBSet(RS!CodTipom, "T") & " and "
            SQL = SQL & " numfactu = " & CStr(RS!numfactu) & " and "
            SQL = SQL & " fecfactu = " & DBSet(RS!fecfactu, "F")
            
            If TotalRegistros(SQL) > 1 Then
                Variedad = "AGRUPADA"
            Else
                SQL = "select nomvarie from variedades where codvarie in ( "
                SQL = SQL & "select codvarie from rfactsoc_variedad where "
                SQL = SQL & " codtipom = " & DBSet(RS!CodTipom, "T") & " and "
                SQL = SQL & " numfactu = " & CStr(RS!numfactu) & " and "
                SQL = SQL & " fecfactu = " & DBSet(RS!fecfactu, "F") & " ) "
            
                Variedad = DevuelveValor(SQL)
            End If

            Select Case TipoFact
                Case 0 ' anticipos
                    c1 = CargaParametroFac(vParamAplic.C1Anticipo, RS, Rs2, Variedad)
                    c2 = CargaParametroFac(vParamAplic.C2Anticipo, RS, Rs2, Variedad)
                    c3 = CargaParametroFac(vParamAplic.C3Anticipo, RS, Rs2, Variedad)
                    c4 = CargaParametroFac(vParamAplic.C4Anticipo, RS, Rs2, Variedad)
                Case 1
                    c1 = CargaParametroFac(vParamAplic.C1Liquidacion, RS, Rs2, Variedad)
                    c2 = CargaParametroFac(vParamAplic.C2Liquidacion, RS, Rs2, Variedad)
                    c3 = CargaParametroFac(vParamAplic.C3Liquidacion, RS, Rs2, Variedad)
                    c4 = CargaParametroFac(vParamAplic.C4Liquidacion, RS, Rs2, Variedad)
                Case 2 ' rectificativas
                    TipoFact1 = DevuelveValor("select tipodocu from usuarios.stipom where codtipom = " & DBSet(RS!rectif_codtipom, "T"))
                    
                    Select Case TipoFact1 ' tipo de la factura que rectifica
                        Case 1
                            c1 = CargaParametroFac(vParamAplic.C1Anticipo, RS, Rs2, Variedad)
                            c2 = CargaParametroFac(vParamAplic.C2Anticipo, RS, Rs2, Variedad)
                            c3 = CargaParametroFac(vParamAplic.C3Anticipo, RS, Rs2, Variedad)
                            c4 = CargaParametroFac(vParamAplic.C4Anticipo, RS, Rs2, Variedad)
                        Case 2
                            c1 = CargaParametroFac(vParamAplic.C1Liquidacion, RS, Rs2, Variedad)
                            c2 = CargaParametroFac(vParamAplic.C2Liquidacion, RS, Rs2, Variedad)
                            c3 = CargaParametroFac(vParamAplic.C3Liquidacion, RS, Rs2, Variedad)
                            c4 = CargaParametroFac(vParamAplic.C4Liquidacion, RS, Rs2, Variedad)
                    End Select
                    
            End Select
            
            '[Monica]20/04/2011: en el importe total factura no estan incluidos los gastos
            SQL = "select sum(importe) from rfactsoc_gastos where "
            SQL = SQL & " codtipom = " & DBSet(RS!CodTipom, "T") & " and "
            SQL = SQL & " numfactu = " & CStr(RS!numfactu) & " and "
            SQL = SQL & " fecfactu = " & DBSet(RS!fecfactu, "F") & " ) "
            
            Gastos = DevuelveValor(SQL)
            '[Monica]20/04/2011
            
            
            f1 = RS!fecfactu
            i1 = RS!TotalFac + Gastos
            f3 = Now
            If impor.importaFicheroPDF(FicheroPDF, CLng(txtCarp(0)), c1, c2, c3, c4, f1, f3, i1) Then
                'actualizamos el pasaridoc de facturas socios
                SQL = "update rfactsoc set pasaridoc = 1 where codtipom = " & DBSet(RS!CodTipom, "T")
                SQL = SQL & " and numfactu = " & DBSet(RS!numfactu, "N") & " and fecfactu = " & DBSet(RS!fecfactu, "F")
                db.ejecutar SQL
            End If
            
            Unload fr
            Set fr = Nothing
            
            RS.MoveNext
        Wend
    End If
    Exit Sub
err_CargaFacturas:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "CargaFacturas"
    End If
End Sub

Private Sub CargaFacturasADV(DFecha As Date, HFecha As Date)
    
    Dim db As BaseDatos
    Dim SQL As String
    Dim RS As ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Dim I As Long
    Dim FicheroPDF As String
    Dim c1 As String
    Dim c2 As String
    Dim c3 As String
    Dim c4 As String
    Dim f1 As Date
    Dim f3 As Date
    Dim i1 As Currency
    Dim fr As frmVisReport
    
    Dim Variedad As String
    
On Error GoTo err_CargaFacturas
    
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim numParam As Byte
Dim cadparam As String

    Set db = New BaseDatos
    db.Tipo = "MYSQL"
    
    db.abrir_MYSQL vConfig.SERVER, vUsu.CadenaConexion, "root", "aritel"


'    db.abrir "accArigasol", "", ""
    SQL = "select advfacturas.*, stipom.letraser " & _
            " from advfacturas, usuarios.stipom where advfacturas.fecfactu >= " & db.Fecha(CDate(txtCodigo(0).Text)) & _
            " and advfacturas.fecfactu <= " & db.Fecha(CDate(txtCodigo(1).Text)) & _
            " and advfacturas.codtipom = stipom.codtipom " & _
            " and advfacturas.pasaridoc = 0"
    
    Set RS = db.cursor(SQL)
    
    If Not RS.EOF Then
        RS.MoveFirst
        While Not RS.EOF
            I = I + 1
            lblInf.Caption = "Procesando registro " & CStr(I)
            lblInf.Refresh
            '-- Creamos el pdf
            FicheroPDF = App.Path & "\ExpAriDoc.pdf"
            
'18/02/2010: lo quito para que prueben
'            If Not IntentaMatar(FicheroPDF) Then Err.Raise 53
            
            Set fr = New frmVisReport
            
            '++monica: seleccionamos que rpt se ha de ejecutar

'            cadParam = "pEmpresa=""AriagroRec""|"
            cadparam = ""
            numParam = 1
            indRPT = 32 'Impresion de Factura de adv
            If Not PonerParamRPT(indRPT, cadparam, numParam, nomDocu) Then Exit Sub
            '++
            fr.NumeroParametros = numParam
            fr.OtrosParametros = cadparam
            fr.ConSubInforme = False
            fr.Informe = App.Path & "\Informes\" & nomDocu
            fr.FormulaSeleccion = "{advfacturas.codtipom} = '" & RS!CodTipom & "' and " & _
                                  "{advfacturas.numfactu} =" & CStr(RS!numfactu) & " and " & _
                                  "{advfacturas.fecfactu} = Date(" & Format(RS!fecfactu, "yyyy") & _
                                                        "," & Format(RS!fecfactu, "mm") & _
                                                        "," & Format(RS!fecfactu, "dd") & ")"
            fr.FicheroPDF = FicheroPDF
            Load fr 'trabaja sin mostrar el formulario
            Screen.MousePointer = vbDefault

            SQL = "select * from rsocios where codsocio = " & db.numero(RS!Codsocio)
            Set Rs2 = db.cursor(SQL)

            c1 = CargaParametroFacADV(vParamAplic.C1ADV, RS, Rs2)
            c2 = CargaParametroFacADV(vParamAplic.C2ADV, RS, Rs2)
            c3 = CargaParametroFacADV(vParamAplic.C3ADV, RS, Rs2)
            c4 = CargaParametroFacADV(vParamAplic.C4ADV, RS, Rs2)
            
            f1 = RS!fecfactu
            i1 = RS!TotalFac
            f3 = Now
            If impor.importaFicheroPDF(FicheroPDF, CLng(txtCarp(0)), c1, c2, c3, c4, f1, f3, i1) Then
                'actualizamos el pasaridoc de facturas socios
                SQL = "update advfacturas set pasaridoc = 1 where codtipom = " & DBSet(RS!CodTipom, "T")
                SQL = SQL & " and numfactu = " & DBSet(RS!numfactu, "N") & " and fecfactu = " & DBSet(RS!fecfactu, "F")
                db.ejecutar SQL
            End If
            
            Unload fr
            Set fr = Nothing
            
            RS.MoveNext
        Wend
    End If
    Exit Sub
err_CargaFacturas:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "CargaFacturasADV"
    End If
End Sub

Private Sub CargaRecibos(DFecha As Date, HFecha As Date)
    Dim db As BaseDatos
    Dim SQL As String
    Dim RS As ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Dim I As Long
    Dim FicheroPDF As String
    Dim c1 As String
    Dim c2 As String
    Dim c3 As String
    Dim c4 As String
    Dim f1 As Date
    Dim f3 As Date
    Dim i1 As Currency
    Dim fr As frmVisReport
On Error GoTo err_CargaRecibos
    
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim numParam As Byte
Dim cadparam As String

    Set db = New BaseDatos
    db.Tipo = "MYSQL"
    
    db.abrir_MYSQL vConfig.SERVER, vUsu.CadenaConexion, "root", "aritel"


'    db.abrir "accArigasol", "", ""
    SQL = "select horas.codtraba " & _
            " from horas where fecharec = " & db.Fecha(CDate(txtCodigo(2).Text)) & _
            " and horas.pasaridoc = 0 " & _
            " and codtraba in (select codtraba from straba where codsecci = " & Combo1(0).ListIndex & ")" & _
            " group by codtraba "
            
    Set RS = db.cursor(SQL)
    
    
    If Not RS.EOF Then
        RS.MoveFirst
        While Not RS.EOF
            I = I + 1
            lblInf.Caption = "Procesando registro " & CStr(I)
            lblInf.Refresh
            '-- Creamos el pdf
            FicheroPDF = App.Path & "\ExpAriDoc.pdf"

'18/02/2010: lo quito para que prueben
'            If Not IntentaMatar(FicheroPDF) Then Err.Raise 53
            
            Set fr = New frmVisReport
            
            '++monica: seleccionamos que rpt se ha de ejecutar
            cadparam = "pEmpresa=""Ariagro""|"
            numParam = 1
            indRPT = 13 'Impresion de Factura
            If Not PonerParamRPT(indRPT, cadparam, numParam, nomDocu, True) Then Exit Sub
            '++
            cadparam = cadparam & "|pFecha=""" & txtCodigo(2).Text & """|"
            numParam = numParam + 1
            cadparam = cadparam & "|pTitulo=""" & "Recibo Horas " & Combo1(0).Text & """|"
            numParam = numParam + 1
            cadparam = cadparam & "|pHProductivas=" & Check1(1).Value & "|"
            numParam = numParam + 1
            
            
            
            fr.NumeroParametros = numParam
            fr.OtrosParametros = cadparam
            fr.ConSubInforme = False
            fr.Informe = App.Path & "\Informes\" & nomDocu
            fr.FormulaSeleccion = "{horas.codtraba} = " & RS!CodTraba & " and " & _
                                           "{horas.fecharec} = Date(" & Format(CDate(txtCodigo(2).Text), "yyyy") & _
                                                                    "," & Format(CDate(txtCodigo(2).Text), "mm") & _
                                                                    "," & Format(CDate(txtCodigo(2).Text), "dd") & ") and " & _
                                           "{horas.pasaridoc} = 0 "
                                                                    
            fr.FicheroPDF = FicheroPDF
            Load fr 'trabaja sin mostrar el formulario
            Screen.MousePointer = vbDefault
'--monica
'            sql = "select * from clientes where codclien = " & db.numero(RS!CodClien)
'            Set Rs2 = db.cursor(sql)
'            c1 = Rs2!nomclien
'            c2 = Format(RS!numfactu, "0000000") & "-" & RS!letraser
'            c3 = "ARIAGRO"
'            c4 = RS!CodClien
'++monica: c1 a c4 esta parametrizado
            SQL = "select * from straba where codtraba = " & db.numero(RS!CodTraba)
            Set Rs2 = db.cursor(SQL)
            c1 = CargaParametroRec(vParamAplic.C1Recibo, RS, Rs2)
            c2 = CargaParametroRec(vParamAplic.C2Recibo, RS, Rs2)
            c3 = CargaParametroRec(vParamAplic.C3Recibo, RS, Rs2)
            c4 = CargaParametroRec(vParamAplic.C4Recibo, RS, Rs2)
            
'            f1 = RS!fechahora
'            i1 = RS!TotalFac
            f1 = CDate(txtCodigo(2).Text)
            i1 = 0
            f3 = Now
            If impor.importaFicheroPDF(FicheroPDF, CLng(txtCarp(0)), c1, c2, c3, c4, f1, f3, i1) Then
                'actualizamos el pasaridoc de facturas
                SQL = "update horas set pasaridoc = 1 where codtraba = " & DBSet(RS!CodTraba, "N")
                SQL = SQL & " and fecharec = " & db.Fecha(CDate(txtCodigo(2).Text))
    '            SQL = SQL & " and fechahora = " & DBSet(RS!fechahora, "F")
                db.ejecutar SQL
            End If
            
            Unload fr
            Set fr = Nothing
            
            RS.MoveNext
        Wend
    End If
    Exit Sub
err_CargaRecibos:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "CargaRecibos"
    End If
End Sub




Private Sub CargaFacturasBod(DFecha As Date, HFecha As Date, Tipo As Byte)
' Tipo : 0=facturas de almazara
'        1=facturas de bodega

    
    Dim db As BaseDatos
    Dim SQL As String
    Dim RS As ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Dim I As Long
    Dim FicheroPDF As String
    Dim c1 As String
    Dim c2 As String
    Dim c3 As String
    Dim c4 As String
    Dim f1 As Date
    Dim f3 As Date
    Dim i1 As Currency
    Dim fr As frmVisReport
    
    Dim Variedad As String
    
On Error GoTo err_CargaFacturas
    
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim numParam As Byte
Dim cadparam As String
Dim codigoTipom As String

    Set db = New BaseDatos
    db.Tipo = "MYSQL"
    
    db.abrir_MYSQL vConfig.SERVER, vUsu.CadenaConexion, "root", "aritel"

    Select Case Tipo
        Case 0 ' almazara
            codigoTipom = "ZA"
        Case 1 ' bodega
            codigoTipom = "AB"
    End Select
            

'    db.abrir "accArigasol", "", ""
    SQL = "select rbodfacturas.*, stipom.letraser " & _
            " from rbodfacturas, usuarios.stipom where rbodfacturas.fecfactu >= " & db.Fecha(CDate(txtCodigo(0).Text)) & _
            " and rbodfacturas.fecfactu <= " & db.Fecha(CDate(txtCodigo(1).Text)) & _
            " and rbodfacturas.codtipom = stipom.codtipom " & _
            " and rbodfacturas.pasaridoc = 0 " & _
            " and mid(rbodfacturas.codtipom,2,2) = " & DBSet(codigoTipom, "T")
    
    Set RS = db.cursor(SQL)
    
    If Not RS.EOF Then
        RS.MoveFirst
        While Not RS.EOF
            I = I + 1
            lblInf.Caption = "Procesando registro " & CStr(I)
            lblInf.Refresh
            '-- Creamos el pdf
            FicheroPDF = App.Path & "\ExpAriDoc.pdf"
            
'18/02/2010: lo quito para que prueben
'            If Not IntentaMatar(FicheroPDF) Then Err.Raise 53
            
            Set fr = New frmVisReport
            
            '++monica: seleccionamos que rpt se ha de ejecutar

'            cadParam = "pEmpresa=""AriagroRec""|"
            cadparam = ""
            numParam = 1
            indRPT = 41 'Impresion de Factura de retirada de almazara / bodega
            If Not PonerParamRPT(indRPT, cadparam, numParam, nomDocu) Then Exit Sub
            '++
            fr.NumeroParametros = numParam
            fr.OtrosParametros = cadparam
            fr.ConSubInforme = False
            fr.Informe = App.Path & "\Informes\" & nomDocu
            fr.FormulaSeleccion = "{rbodfacturas.codtipom} = '" & RS!CodTipom & "' and " & _
                                  "{rbodfacturas.numfactu} =" & CStr(RS!numfactu) & " and " & _
                                  "{rbodfacturas.fecfactu} = Date(" & Format(RS!fecfactu, "yyyy") & _
                                                        "," & Format(RS!fecfactu, "mm") & _
                                                        "," & Format(RS!fecfactu, "dd") & ")"
            fr.FicheroPDF = FicheroPDF
            Load fr 'trabaja sin mostrar el formulario
            Screen.MousePointer = vbDefault

            
            
            '++monica: si hay mas una variedad meter agrupada
            '          si solo hay una meter la variedad
            
            SQL = "select count(distinct codvarie) from rbodfacturas_lineas where "
            SQL = SQL & " codtipom = " & DBSet(RS!CodTipom, "T") & " and "
            SQL = SQL & " numfactu = " & CStr(RS!numfactu) & " and "
            SQL = SQL & " fecfactu = " & DBSet(RS!fecfactu, "F")
            
            If TotalRegistros(SQL) > 1 Then
                Variedad = "AGRUPADA"
            Else
                SQL = "select nomvarie from variedades where codvarie in ( "
                SQL = SQL & "select codvarie from rbodfacturas_lineas where "
                SQL = SQL & " codtipom = " & DBSet(RS!CodTipom, "T") & " and "
                SQL = SQL & " numfactu = " & CStr(RS!numfactu) & " and "
                SQL = SQL & " fecfactu = " & DBSet(RS!fecfactu, "F") & " ) "
            
                Variedad = DevuelveValor(SQL)
            End If
            
            
            SQL = "select * from rsocios where codsocio = " & db.numero(RS!Codsocio)
            Set Rs2 = db.cursor(SQL)

            Select Case Tipo
                Case 0 ' almazara
                    c1 = CargaParametroFacBOD(vParamAplic.C1Almz, RS, Rs2, Variedad)
                    c2 = CargaParametroFacBOD(vParamAplic.C2Almz, RS, Rs2, Variedad)
                    c3 = CargaParametroFacBOD(vParamAplic.C3Almz, RS, Rs2, Variedad)
                    c4 = CargaParametroFacBOD(vParamAplic.C4Almz, RS, Rs2, Variedad)
                Case 1 ' bodega
                    c1 = CargaParametroFacBOD(vParamAplic.C1BOD, RS, Rs2, Variedad)
                    c2 = CargaParametroFacBOD(vParamAplic.C2BOD, RS, Rs2, Variedad)
                    c3 = CargaParametroFacBOD(vParamAplic.C3BOD, RS, Rs2, Variedad)
                    c4 = CargaParametroFacBOD(vParamAplic.C4BOD, RS, Rs2, Variedad)
            End Select
                        
            
            f1 = RS!fecfactu
            i1 = RS!TotalFac
            f3 = Now
            If impor.importaFicheroPDF(FicheroPDF, CLng(txtCarp(0)), c1, c2, c3, c4, f1, f3, i1) Then
                'actualizamos el pasaridoc de facturas socios
                SQL = "update rbodfacturas set pasaridoc = 1 where codtipom = " & DBSet(RS!CodTipom, "T")
                SQL = SQL & " and numfactu = " & DBSet(RS!numfactu, "N") & " and fecfactu = " & DBSet(RS!fecfactu, "F")
                db.ejecutar SQL
            End If
            
            Unload fr
            Set fr = Nothing
            
            RS.MoveNext
        Wend
    End If
    Exit Sub
err_CargaFacturas:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "CargaFacturasRetirada"
    End If
End Sub


Private Sub CargaFacturasTransporte(TipoFact As Byte, DFecha As Date, HFecha As Date)
' TipoFact: 0 = Anticipos
'           1 = Liquidacion

    Dim db As BaseDatos
    Dim SQL As String
    Dim RS As ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Dim I As Long
    Dim FicheroPDF As String
    Dim c1 As String
    Dim c2 As String
    Dim c3 As String
    Dim c4 As String
    Dim f1 As Date
    Dim f3 As Date
    Dim i1 As Currency
    Dim fr As frmVisReport
    
    Dim Variedad As String
    
    Dim TipoFact1 As Byte
    
On Error GoTo err_CargaFacturas
    
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim numParam As Byte
Dim cadparam As String

    Set db = New BaseDatos
    db.Tipo = "MYSQL"
    
    db.abrir_MYSQL vConfig.SERVER, vUsu.CadenaConexion, "root", "aritel"


'    db.abrir "accArigasol", "", ""
    SQL = "select rfacttra.*, stipom.letraser " & _
            " from rfacttra, usuarios.stipom where rfacttra.fecfactu >= " & db.Fecha(CDate(txtCodigo(0).Text)) & _
            " and rfacttra.fecfactu <= " & db.Fecha(CDate(txtCodigo(1).Text)) & _
            " and rfacttra.codtipom = stipom.codtipom " & _
            " and rfacttra.pasaridoc = 0"
            
    Set RS = db.cursor(SQL)
    
    If Not RS.EOF Then
        RS.MoveFirst
        While Not RS.EOF
            I = I + 1
            lblInf.Caption = "Procesando registro " & CStr(I)
            lblInf.Refresh
            '-- Creamos el pdf
            FicheroPDF = App.Path & "\ExpAriDoc.pdf"
            
'18/02/2010: lo quito para que prueben
'            If Not IntentaMatar(FicheroPDF) Then Err.Raise 53
            
            
            Set fr = New frmVisReport
            
            '++monica: seleccionamos que rpt se ha de ejecutar

'            cadParam = "pEmpresa=""AriagroRec""|"
            cadparam = ""
            numParam = 1
            
            indRPT = 49
            If Not PonerParamRPT(indRPT, cadparam, numParam, nomDocu) Then Exit Sub
            '++
            fr.NumeroParametros = numParam
            fr.OtrosParametros = cadparam
            fr.ConSubInforme = False
            fr.Informe = App.Path & "\Informes\" & nomDocu
            fr.FormulaSeleccion = "{rfacttra.codtipom} = '" & RS!CodTipom & "' and " & _
                                  "{rfacttra.numfactu} =" & CStr(RS!numfactu) & " and " & _
                                  "{rfacttra.fecfactu} = Date(" & Format(RS!fecfactu, "yyyy") & _
                                                        "," & Format(RS!fecfactu, "mm") & _
                                                        "," & Format(RS!fecfactu, "dd") & ")"
            fr.FicheroPDF = FicheroPDF
            Load fr 'trabaja sin mostrar el formulario
            Screen.MousePointer = vbDefault

            SQL = "select * from rtransporte where codtrans = " & db.Texto(RS!codTrans)
            Set Rs2 = db.cursor(SQL)
            
            '++monica: si hay mas una variedad meter agrupada
            '          si solo hay una meter la variedad
            
            SQL = "select count(distinct codvarie) from rfacttra_albaran where "
            SQL = SQL & " codtipom = " & DBSet(RS!CodTipom, "T") & " and "
            SQL = SQL & " numfactu = " & CStr(RS!numfactu) & " and "
            SQL = SQL & " fecfactu = " & DBSet(RS!fecfactu, "F")
            
            If TotalRegistros(SQL) > 1 Then
                Variedad = "AGRUPADA"
            Else
                SQL = "select nomvarie from variedades where codvarie in ( "
                SQL = SQL & "select codvarie from rfacttra_variedad where "
                SQL = SQL & " codtipom = " & DBSet(RS!CodTipom, "T") & " and "
                SQL = SQL & " numfactu = " & CStr(RS!numfactu) & " and "
                SQL = SQL & " fecfactu = " & DBSet(RS!fecfactu, "F") & " ) "
            
                Variedad = DevuelveValor(SQL)
            End If

            c1 = CargaParametroFacTra(vParamAplic.C1Transporte, RS, Rs2, Variedad)
            c2 = CargaParametroFacTra(vParamAplic.C2Transporte, RS, Rs2, Variedad)
            c3 = CargaParametroFacTra(vParamAplic.C3Transporte, RS, Rs2, Variedad)
            c4 = CargaParametroFacTra(vParamAplic.C4Transporte, RS, Rs2, Variedad)
                
'                Case 2 ' rectificativas
'                    TipoFact1 = DevuelveValor("select tipodocu from usuarios.stipom where codtipom = " & DBSet(RS!rectif_codtipom, "T"))
'
'                    Select Case TipoFact1 ' tipo de la factura que rectifica
'                        Case 1
'                            c1 = CargaParametroFac(vParamAplic.C1Anticipo, RS, Rs2, Variedad)
'                            c2 = CargaParametroFac(vParamAplic.C2Anticipo, RS, Rs2, Variedad)
'                            c3 = CargaParametroFac(vParamAplic.C3Anticipo, RS, Rs2, Variedad)
'                            c4 = CargaParametroFac(vParamAplic.C4Anticipo, RS, Rs2, Variedad)
'                        Case 2
'                            c1 = CargaParametroFac(vParamAplic.C1Liquidacion, RS, Rs2, Variedad)
'                            c2 = CargaParametroFac(vParamAplic.C2Liquidacion, RS, Rs2, Variedad)
'                            c3 = CargaParametroFac(vParamAplic.C3Liquidacion, RS, Rs2, Variedad)
'                            c4 = CargaParametroFac(vParamAplic.C4Liquidacion, RS, Rs2, Variedad)
'                    End Select
'
'            End Select
            
            f1 = RS!fecfactu
            i1 = RS!TotalFac
            f3 = Now
            If impor.importaFicheroPDF(FicheroPDF, CLng(txtCarp(0)), c1, c2, c3, c4, f1, f3, i1) Then
                'actualizamos el pasaridoc de facturas socios
                SQL = "update rfacttra set pasaridoc = 1 where codtipom = " & DBSet(RS!CodTipom, "T")
                SQL = SQL & " and numfactu = " & DBSet(RS!numfactu, "N") & " and fecfactu = " & DBSet(RS!fecfactu, "F")
                db.ejecutar SQL
            End If
            
            Unload fr
            Set fr = Nothing
            
            RS.MoveNext
        Wend
    End If
    Exit Sub
err_CargaFacturas:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "CargaFacturas"
    End If
End Sub



Private Sub CargaFacturasPozos(TipoFact As Byte, DFecha As Date, HFecha As Date)
' TipoFact: 0 = Consumo
'           1 = mantenimiento

    Dim db As BaseDatos
    Dim SQL As String
    Dim RS As ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Dim I As Long
    Dim FicheroPDF As String
    Dim c1 As String
    Dim c2 As String
    Dim c3 As String
    Dim c4 As String
    Dim f1 As Date
    Dim f3 As Date
    Dim i1 As Currency
    Dim fr As frmVisReport
    
    Dim Variedad As String
    
    Dim TipoFact1 As Byte
    Dim Gastos As Currency
    
On Error GoTo err_CargaFacturas
    
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim numParam As Byte
Dim cadparam As String

    Set db = New BaseDatos
    db.Tipo = "MYSQL"
    
    db.abrir_MYSQL vConfig.SERVER, vUsu.CadenaConexion, "root", "aritel"


'    db.abrir "accArigasol", "", ""
    SQL = "select rrecibpozos.*, stipom.letraser " & _
            " from rrecibpozos, usuarios.stipom where rrecibpozos.fecfactu >= " & db.Fecha(CDate(txtCodigo(0).Text)) & _
            " and rrecibpozos.fecfactu <= " & db.Fecha(CDate(txtCodigo(1).Text)) & _
            " and rrecibpozos.codtipom = stipom.codtipom " & _
            " and rrecibpozos.pasaridoc = 0"
            
    Select Case TipoFact
        Case 0 ' consumo
            '[Monica]01/04/2011 añadidas 7,9 anticipos de almazara y de bodega
            SQL = SQL & " and rrecibpozos.codtipom in ('RCP')"
        Case 1 ' mantenimiento
            '[Monica]01/04/2011 añadidas 8,10 liquidaciones de almazara y de bodega
            SQL = SQL & " and rrecibpozos.codtipom in ('RMP')" ' [Monica]26/04/2010: añadidas 5 y 6
        Case 2 ' facturas de talla
            SQL = SQL & " and rrecibpozos.codtipom in ('TAL')" ' [Monica]29/06/2012: añadidas facturas de talla
        Case 3 ' contadores
            SQL = SQL & " and rrecibpozos.codtipom in ('RVP')" ' [Monica]27/06/2013: añadidas facturas de contadores
        
    End Select
            
    Set RS = db.cursor(SQL)
    
    If Not RS.EOF Then
        RS.MoveFirst
        While Not RS.EOF
            I = I + 1
            lblInf.Caption = "Procesando registro " & CStr(I)
            lblInf.Refresh
            '-- Creamos el pdf
            FicheroPDF = App.Path & "\ExpAriDoc.pdf"
            
'18/02/2010: lo quito para que prueben
'            If Not IntentaMatar(FicheroPDF) Then Err.Raise 53
            
            
            Set fr = New frmVisReport
            
            '++monica: seleccionamos que rpt se ha de ejecutar

'            cadParam = "pEmpresa=""AriagroRec""|"
            cadparam = ""
            numParam = 1
            
                        
            If DBLet(RS!CodTipom, "T") = "RCP" Then
                indRPT = 46
            Else ' RMP y TAL
                indRPT = 47 'Impresion de Factura
            End If
            
            If Not PonerParamRPT(indRPT, cadparam, numParam, nomDocu) Then Exit Sub
            
            If DBLet(RS!CodTipom, "T") = "TAL" Then nomDocu = Replace(nomDocu, "Mto.", "Tal.")
            If DBLet(RS!CodTipom, "T") = "RVP" Then nomDocu = Replace(nomDocu, "Mto.", "Cont.")
            
            '++
            fr.NumeroParametros = numParam
            fr.OtrosParametros = cadparam
            fr.ConSubInforme = True
            fr.Informe = App.Path & "\Informes\" & nomDocu
            fr.FormulaSeleccion = "{rrecibpozos.codtipom} = '" & RS!CodTipom & "' and " & _
                                  "{rrecibpozos.numfactu} =" & CStr(RS!numfactu) & " and " & _
                                  "{rrecibpozos.fecfactu} = Date(" & Format(RS!fecfactu, "yyyy") & _
                                                        "," & Format(RS!fecfactu, "mm") & _
                                                        "," & Format(RS!fecfactu, "dd") & ")"
            fr.FicheroPDF = FicheroPDF
            Load fr 'trabaja sin mostrar el formulario
            Screen.MousePointer = vbDefault

            SQL = "select * from rsocios where codsocio = " & db.numero(RS!Codsocio)
            Set Rs2 = db.cursor(SQL)
            
            Variedad = ""

            Select Case TipoFact
                Case 0 ' consumo
                    c1 = CargaParametroFac(vParamAplic.C1Anticipo, RS, Rs2, Variedad)
                    c2 = CargaParametroFac(vParamAplic.C2Anticipo, RS, Rs2, Variedad)
                    c3 = CargaParametroFac(vParamAplic.C3Anticipo, RS, Rs2, Variedad)
                    c4 = CargaParametroFac(vParamAplic.C4Anticipo, RS, Rs2, Variedad)
                Case 1 ' mantenimiento
                    c1 = CargaParametroFac(vParamAplic.C1Liquidacion, RS, Rs2, Variedad)
                    c2 = CargaParametroFac(vParamAplic.C2Liquidacion, RS, Rs2, Variedad)
                    c3 = CargaParametroFac(vParamAplic.C3Liquidacion, RS, Rs2, Variedad)
                    c4 = CargaParametroFac(vParamAplic.C4Liquidacion, RS, Rs2, Variedad)
                Case 2 ' facturas de talla
                    c1 = CargaParametroFac(vParamAplic.C1Liquidacion, RS, Rs2, Variedad)
                    c2 = CargaParametroFac(vParamAplic.C2Liquidacion, RS, Rs2, Variedad)
                    c3 = CargaParametroFac(vParamAplic.C3Liquidacion, RS, Rs2, Variedad)
                    c4 = CargaParametroFac(vParamAplic.C4Liquidacion, RS, Rs2, Variedad)
            End Select
            
            f1 = RS!fecfactu
            i1 = RS!TotalFact
            f3 = Now
            If impor.importaFicheroPDF(FicheroPDF, CLng(txtCarp(0)), c1, c2, c3, c4, f1, f3, i1) Then
                'actualizamos el pasaridoc de facturas socios
                SQL = "update rrecibpozos set pasaridoc = 1 where codtipom = " & DBSet(RS!CodTipom, "T")
                SQL = SQL & " and numfactu = " & DBSet(RS!numfactu, "N") & " and fecfactu = " & DBSet(RS!fecfactu, "F")
                db.ejecutar SQL
            End If
            
            Unload fr
            Set fr = Nothing
            
            RS.MoveNext
        Wend
    End If
    Exit Sub
err_CargaFacturas:
    If Err.Number <> 0 Then
        MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "CargaFacturas"
    End If
End Sub






Private Function CargaParametroFac(param As Byte, ByRef RS As ADODB.Recordset, ByRef Rs2 As ADODB.Recordset, NomVar As String) As String
    Select Case param
        Case 0 'facturas
            CargaParametroFac = Format(RS!numfactu, "0000000") & "-" & RS!letraser
        Case 1 'codigo socio
            CargaParametroFac = RS!Codsocio
        Case 2 'nombre socio
            CargaParametroFac = Rs2!nomsocio
        Case 3 'variedad???
            CargaParametroFac = NomVar
        Case Else
            CargaParametroFac = ""
    End Select

End Function

Private Function CargaParametroFacADV(param As Byte, ByRef RS As ADODB.Recordset, ByRef Rs2 As ADODB.Recordset) As String
Dim SQL As String
Dim rs3 As ADODB.Recordset
Dim db As BaseDatos

    Set db = New BaseDatos
    db.Tipo = "MYSQL"

    db.abrir_MYSQL vConfig.SERVER, vUsu.CadenaConexion, "root", "aritel"

    Select Case param
        Case 0 'factura
            CargaParametroFacADV = Format(RS!numfactu, "0000000") & "-" & RS!letraser
        Case 1 'codigo socio
            CargaParametroFacADV = RS!Codsocio
        Case 2 'nombre socio
            CargaParametroFacADV = Rs2!nomsocio
        Case 3 'destino
            CargaParametroFacADV = "ARIAGROREC"
        Case 4 'procedencia
            CargaParametroFacADV = "ARIAGROREC"
        Case Else
            CargaParametroFacADV = ""
    End Select
End Function

Private Function CargaParametroRec(param As Byte, ByRef RS As ADODB.Recordset, ByRef Rs2 As ADODB.Recordset) As String
    Select Case param
        Case 0 'facturas
'            CargaParametroRec = Format(RS!numfactu, "0000000") & "-" & RS!letraser
            CargaParametroRec = RS!CodTraba
        Case 1 'codigo trabajador
            CargaParametroRec = Rs2!NomTraba
        Case 2 'nombre trabajador
            CargaParametroRec = "ARIAGROREC"
        Case 3 'procedencia
            CargaParametroRec = "ARIAGROREC"
        Case Else
            CargaParametroRec = ""
    End Select

End Function

Private Function CargaParametroFacBOD(param As Byte, ByRef RS As ADODB.Recordset, ByRef Rs2 As ADODB.Recordset, NomVar As String) As String
Dim SQL As String
Dim rs3 As ADODB.Recordset
Dim db As BaseDatos

    Set db = New BaseDatos
    db.Tipo = "MYSQL"

    db.abrir_MYSQL vConfig.SERVER, vUsu.CadenaConexion, "root", "aritel"

    Select Case param
        Case 0 'factura
            CargaParametroFacBOD = Format(RS!numfactu, "0000000") & "-" & RS!letraser
        Case 1 'codigo socio
            CargaParametroFacBOD = RS!Codsocio
        Case 2 'nombre socio
            CargaParametroFacBOD = Rs2!nomsocio
        Case 3 'variedades
            CargaParametroFacBOD = NomVar
        Case Else
            CargaParametroFacBOD = ""
    End Select
End Function


Private Function CargaParametroFacTra(param As Byte, ByRef RS As ADODB.Recordset, ByRef Rs2 As ADODB.Recordset, NomVar As String) As String
    Select Case param
        Case 0 'facturas
            CargaParametroFacTra = Format(RS!numfactu, "0000000") & "-" & RS!letraser
        Case 1 'codigo transportista
            CargaParametroFacTra = RS!codTrans
        Case 2 'nombre socio
            CargaParametroFacTra = Rs2!nomtrans
        Case 3 'variedad???
            CargaParametroFacTra = NomVar
        Case Else
            CargaParametroFacTra = ""
    End Select

End Function



Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim I As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
'    For I = 0 To Combo1.Count - 1
'        Combo1(I).Clear
'    Next I
    Combo1(0).Clear
    
    Combo1(0).AddItem "Campo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Almacén"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1

    If Tipo = 6 Then
        Combo1(1).AddItem "Consumo"
        Combo1(1).ItemData(Combo1(1).NewIndex) = 0
        Combo1(1).AddItem "Mantenimiento"
        Combo1(1).ItemData(Combo1(1).NewIndex) = 1
        Combo1(1).AddItem "Talla"
        Combo1(1).ItemData(Combo1(1).NewIndex) = 2
        Combo1(1).AddItem "Contador"
        Combo1(1).ItemData(Combo1(1).NewIndex) = 3
    Else
        Combo1(1).AddItem "Anticipo"
        Combo1(1).ItemData(Combo1(1).NewIndex) = 0
        Combo1(1).AddItem "Liquidacion"
        Combo1(1).ItemData(Combo1(1).NewIndex) = 1
        Combo1(1).AddItem "Rectificativa"
        Combo1(1).ItemData(Combo1(1).NewIndex) = 2
        
        If Tipo > 0 Then
            Combo1(1).AddItem "ADV"
            Combo1(1).ItemData(Combo1(1).NewIndex) = 3
            Combo1(1).AddItem "Almazara"
            Combo1(1).ItemData(Combo1(1).NewIndex) = 4
            Combo1(1).AddItem "Bodega"
            Combo1(1).ItemData(Combo1(1).NewIndex) = 5
            Combo1(1).AddItem "Nóminas"
            Combo1(1).ItemData(Combo1(1).NewIndex) = 6
            Combo1(1).AddItem "Transporte"
            Combo1(1).ItemData(Combo1(1).NewIndex) = 7
        End If
    End If
End Sub


Private Function CargaPath(Codigo As Integer) As String
Dim Nod As Node
Dim J As Integer
Dim I As Integer
Dim C As String
Dim campo1 As String
Dim padre As String
Dim A As String

Dim SQL As String
Dim RS As ADODB.Recordset

    'distinto del cargapath de parametros de aplicacion

    SQL = "select nombre, padre from carpetas where codcarpeta = " & DBSet(Codigo, "N")
    Set RS = ardDB.cursor(SQL)

    If Not RS.EOF Then
        C = "\" & RS!Nombre
        If RS!padre > 0 Then
            C = CargaPath(CInt(RS!padre)) & C
        End If
    End If
    
    CargaPath = C
End Function



Private Function IntentaMatar(FicheroPDF As String) As Boolean
Dim I As Integer

    On Error Resume Next
    I = 1
    IntentaMatar = False
    Do
        If Dir(FicheroPDF, vbArchive) <> "" Then
            Kill FicheroPDF
            If Err.Number <> 0 Then
                Err.Clear
                I = I + 1
            Else
                IntentaMatar = True
                I = 6
            End If
        Else
            IntentaMatar = True
            I = 6
        End If
    Loop Until I < 5 Or IntentaMatar = True
    
    
End Function






