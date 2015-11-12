VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form frmActualizar2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualizar diario"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   Icon            =   "frmActualizar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameResultados 
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   5115
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         Top             =   4200
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3495
         Left            =   60
         TabIndex        =   10
         Top             =   360
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   6165
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nº Asien"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha Entrada"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "Obteniendo resultados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   1920
         Width           =   3735
      End
      Begin VB.Label Label8 
         Caption         =   "Errores:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   7
         Top             =   60
         Width           =   855
      End
   End
   Begin VB.Frame frame1Asiento 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin ComCtl2.Animation Animation1 
         Height          =   735
         Left            =   600
         TabIndex        =   4
         Top             =   1800
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1296
         _Version        =   327681
         FullWidth       =   241
         FullHeight      =   49
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label9 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Width           =   3975
      End
      Begin VB.Label lblAsiento 
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Asiento :"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   3855
      End
   End
End
Attribute VB_Name = "frmActualizar2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public OpcionActualizar As Byte
    '1.- Actualizar 1 asiento
    '2.- Desactualiza pero NO insertes en apuntes
    '3.- Desactualizar asiento desde hco
    
    'Si el asiento es de una factura entonces NUMSERIE tendra "FRACLI" o "FRAPRO"
    ' con lo cual habra que poner su factura asociada a NULL
    
    '4.- Si es para enviar datos a impresora
    '5.- Actualiza mas de 1 asiento
    
    '6.- Integra 1 factura
    '7.- Elimina factura integrada . DesINTEGRA   . C L I E N T E S
    '8.- Integra 1 factura PROVEEDORES
    '9.- Elimina factura integrada . Desintegra.    P R O V E E D O R E S
    
    '10 .- Integracion masiva facturas clientes
    '11 .- Integracion masiva facturas Proveedores
    
    
    '12 .- Recalcular saldos desde hlinapu

    '13 .- IMPRIMIR asientos errores
    
    
    '---------------- DE TESORERIA
    '20
    
Public Numasiento As Long
Public FechaAsiento As Date
Public numdiari As Integer
Public numserie As String




Private cuenta As String
Private ImporteD As Currency
Private ImporteH As Currency
Private CCost As String
'Y estas son privadas
Private mes As Integer
Private Anyo As Integer
Dim Fecha As String  'TENDRA la fecha ya formateada en yyy-mm-dd
Dim PrimeraVez As Boolean
Dim SQL As String
Dim RS As Recordset

Dim INC As Long

Dim NE As Integer
Dim ErroresAbiertos As Boolean
Dim NumErrores As Long

Dim ItmX As ListItem  'Para mostra errores masivos

Private Sub AñadeError(ByRef Mensaje As String)
On Error Resume Next
'Escribimos en el fichero
If Not ErroresAbiertos Then
    NE = FreeFile
    ErroresAbiertos = True
    Open App.Path & "\ErrActua.txt" For Output As NE
    If Err.Number <> 0 Then
        MsgBox " Error abriendo fichero errores", vbExclamation
        Err.Clear
    End If
End If
Print #NE, Mensaje
If Err.Number <> 0 Then
    Err.Clear
    NumErrores = -20000
Else
    NumErrores = NumErrores + 1
End If
End Sub



'Private Function CadenaImporte(VaAlDebe As Boolean, ByRef Importe As Currency, ElImporteEsCero As Boolean) As String
'Dim CadImporte As String
'
''Si va al debe, pero el importe es negativo entonces va al haber a no ser que la contabilidad admita importes negativos
'    If Importe < 0 Then
'        If Not vParam.abononeg Then
'            VaAlDebe = Not VaAlDebe
'            Importe = Abs(Importe)
'        End If
'    End If
'    ElImporteEsCero = (Importe = 0)
'    CadImporte = TransformaComasPuntos(CStr(Importe))
'    If VaAlDebe Then
'        CadenaImporte = CadImporte & ",NULL"
'    Else
'        CadenaImporte = "NULL," & CadImporte
'    End If
'End Function

Private Sub CargaProgres(Valor As Integer)
Me.ProgressBar1.Max = Valor
Me.ProgressBar1.Value = 0
End Sub





Private Sub IncrementaProgres(Veces As Integer)
On Error Resume Next
Me.ProgressBar1.Value = Me.ProgressBar1.Value + (Veces * INC)
If Err.Number <> 0 Then Err.Clear
Me.Refresh
End Sub




Private Sub CargaListAsiento()

NE = FreeFile
If Dir(App.Path & "\ErrActua.txt") = "" Then
    'MsgBox "Los errores han sido eliminados. Imposible ver errores. Modulo: CargaLisAsiento"
    Exit Sub
End If

Me.frameResultados.visible = True
'Los encabezados
ListView1.ColumnHeaders.Clear
ListView1.ColumnHeaders.Add , , "Diario", 800
ListView1.ColumnHeaders.Add , , "Fecha", 1000
ListView1.ColumnHeaders.Add , , "Nº Asie.", 1000
ListView1.ColumnHeaders.Add , , "Error", 3000


Open App.Path & "\ErrActua.txt" For Input As #NE
While Not EOF(NE)
    Line Input #NE, cuenta
    Set ItmX = ListView1.ListItems.Add
    ItmX.Text = RecuperaValor(cuenta, 1)
    ItmX.SubItems(1) = RecuperaValor(cuenta, 2)
    ItmX.SubItems(2) = RecuperaValor(cuenta, 3)
    ItmX.SubItems(3) = RecuperaValor(cuenta, 4)
Wend
Close #NE
End Sub







Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
Dim bol As Boolean
If PrimeraVez Then
    PrimeraVez = False
    Me.Refresh
    bol = False
    
    'TEnemos que eliminar el archivo de errores
    If OpcionActualizar = 20 Then
        EliminarArchivoErrores
        
    End If
    Select Case OpcionActualizar
    Case 1
        If ActualizaAsiento Then ProcesoCorrecto = True
        bol = True
    Case 2, 3
        
        bol = True
    Case 4, 13

        numserie = ""
    Case 6, 8

        bol = True
    Case 7, 9

        bol = True
    Case 10, 11
    
    Case 20
        'COBROS, pagos
        lblAsiento.Caption = "Actualizando registros"
        lblAsiento.Refresh
        If ObtenerRegistrosParaActualizar Then bol = True
        
      
    End Select
    If bol Then Unload Me
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_DblClick()
    CargaListAsiento
End Sub

Private Sub Form_Load()
Dim b As Boolean
    ErroresAbiertos = False
    limpiar Me
 
    PrimeraVez = True
    Me.frameResultados.visible = False
    NumErrores = 0
    ListView1.ListItems.Clear
    Select Case OpcionActualizar
    Case 1, 2, 3, 20     'Pagos, cobros tambien
        Label1.Caption = "Nº Asiento"
        Me.lblAsiento.Caption = Numasiento
        INC = 10  'Incremento para el proggress
        If OpcionActualizar = 1 Then
            Label9.Caption = "Actualizar"
        Else
            Label9.Caption = "Integracion tesoreria"
        End If
        'Tamaño
        Me.Height = 3200
        b = True
    Case 4, 5, 13
        Me.Height = 4665

        If OpcionActualizar <> 5 Then
 

        Else
            'La opcion 5: Actualizar

        End If
        b = False
    Case 6, 7, 8, 9
        '// Estamos en Facturas
        Label1.Caption = "Nº factura"
        If OpcionActualizar < 8 Then
            Label1.Caption = Label1.Caption & " Cliente"
        Else
            Label1.Caption = Label1.Caption & " Proveedor"
        End If
        Me.lblAsiento.Caption = numserie & Numasiento
        INC = 10  'Incremento para el proggress
        If OpcionActualizar = 6 Or OpcionActualizar = 8 Then
            Label9.Caption = "Integrar Factura"
        Else
            Label9.Caption = "Eliminar Factura"
        End If
        Me.Caption = "Actualizar facturas"
        'Tamaño
        Me.Height = 3315
        b = True
    Case 10, 11

    Case 12

    End Select
    Me.frame1Asiento.visible = b
    Me.Animation1.visible = b
End Sub





Private Function ActualizaAsiento() As Boolean
    Dim bol As Boolean
    Dim Donde As String
    On Error GoTo EActualizaAsiento
    
    'Obtenemos el mes y el año
    mes = Month(FechaAsiento)
    Anyo = Year(FechaAsiento)
    Fecha = Format(FechaAsiento, FormatoFecha)
    
    'Comprobamos que no existe en historico
    If AsientoExiste(True) Then
        MsgBox "El asiento ya existe. Fecha: " & Fecha & "     Nº: " & Numasiento, vbExclamation
        Exit Function
    End If
    
    'Aqui bloquearemos
    
'    conn.BeginTrans
    bol = ActualizaElASiento(Donde)
    
EActualizaAsiento:
        If Err.Number <> 0 Then
            SQL = "Actualiza Asiento." & vbCrLf & "----------------------------" & vbCrLf
            SQL = SQL & Donde
            If OpcionActualizar = 1 Then
                MuestraError Err.Number, SQL, Err.Description
            Else
                SQL = Donde & " -> " & Err.Description
                SQL = Mid(SQL, 1, 200)
                InsertaError SQL
            End If
            bol = False
        End If
        If bol Then
            ActualizaAsiento = True
        Else
            ActualizaAsiento = False
        End If
End Function


Private Function ActualizaElASiento(ByRef A_Donde As String) As Boolean



    ActualizaElASiento = False
    
    'Insertamos en cabeceras
    A_Donde = "Insertando datos en historico cabeceras asiento"
    If Not InsertarCabecera Then Exit Function
    IncrementaProgres 1
    
    'Insertamos en lineas
    A_Donde = "Insertando datos en historico lineas asiento"
    If Not InsertarLineas Then Exit Function
    IncrementaProgres 2
    
    
    
    'Modificar saldos
    A_Donde = "Calculando Lineas y saldos "
    If Not CalcularLineasYSaldos(False) Then Exit Function
    
    
    'Borramos cabeceras y lineas del asiento
    A_Donde = "Borrar cabeceras y lineas en asientos"
    If Not BorrarASiento(False) Then Exit Function
    IncrementaProgres 2
    ActualizaElASiento = True
End Function


Private Function InsertarCabecera() As Boolean
On Error Resume Next

    SQL = "INSERT INTO hcabapu (numdiari, fechaent, numasien, obsdiari) SELECT numdiari,fechaent,numasien,obsdiari from cabapu where "
    SQL = SQL & " numdiari =" & numdiari
    SQL = SQL & " AND fechaent='" & Fecha & "'"
    SQL = SQL & " AND numasien=" & Numasiento

    ConnConta.Execute SQL

    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        InsertarCabecera = False
    Else
        InsertarCabecera = True
    End If
End Function


Private Function InsertarCabeceraApuntes() As Boolean
On Error Resume Next

    SQL = "INSERT INTO cabapu (numdiari, fechaent, numasien, obsdiari) SELECT numdiari,fechaent,numasien,obsdiari from hcabapu where "
    SQL = SQL & " numdiari =" & numdiari
    SQL = SQL & " AND fechaent='" & Fecha & "'"
    SQL = SQL & " AND numasien=" & Numasiento

    conn.Execute SQL

    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        InsertarCabeceraApuntes = False
    Else
        InsertarCabeceraApuntes = True
    End If
End Function



Private Function AsientoExiste(EnHistorico As Boolean) As Boolean
    AsientoExiste = True
    SQL = "SELECT numdiari from "
    If EnHistorico Then
        SQL = SQL & "hcabapu"
    Else
        'k existe en introduccion de apuntes
        SQL = SQL & "cabapu"
    End If
    SQL = SQL & " WHERE numdiari =" & numdiari
    SQL = SQL & " AND fechaent='" & Fecha & "'"
    SQL = SQL & " AND numasien=" & Numasiento
    Set RS = New ADODB.Recordset
    RS.Open SQL, ConnConta, adOpenKeyset, adLockOptimistic, adCmdText
    If RS.EOF Then AsientoExiste = False
    RS.Close
    Set RS = Nothing
End Function


Private Function CalcularLineasYSaldos(EsDesdeRecalcular As Boolean) As Boolean
Dim Reparto As Boolean
Dim T As String

    Dim RL As Recordset
    Set RL = New ADODB.Recordset
    
    
    'Ahora
    SQL = "SELECT timporteD AS SD, timporteH AS SH, codmacta"
    SQL = SQL & "  FROM"
    If EsDesdeRecalcular Then
        SQL = SQL & " hlinapu"
    Else
        SQL = SQL & " linapu"
    End If
    'SQL = SQL & " GROUP BY codmacta, numdiari, fechaent, numasien"
    SQL = SQL & " WHERE (((numdiari)= " & numdiari
    SQL = SQL & ") AND ((fechaent)='" & Fecha & "'"
    SQL = SQL & ") AND ((numasien)=" & Numasiento
    SQL = SQL & "));"
   
    Set RL = New ADODB.Recordset
    RL.Open SQL, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RL.EOF
        cuenta = RL!Codmacta
        If IsNull(RL!sD) Then
            ImporteD = 0
        Else
            'ImporteD = RL!tImporteD
            ImporteD = RL!sD
        End If
        If IsNull(RL!sH) Then
            ImporteH = 0
        Else
            'ImporteH = RL!tImporteH
            ImporteH = RL!sH
        End If
        
        If Not CalcularSaldos Then
            RL.Close
            Exit Function
        End If
        
        'Sig
        RL.MoveNext
    Wend
    RL.Close
    IncrementaProgres 3
    If Not vEmpresa.TieneAnalitica Then
        'NO tiene analitica
        CalcularLineasYSaldos = True
        Exit Function
    End If
    
    
    '------------------------------------------
    '       ANALITICA     -> Modificado para 2 de Julio, para subcentros de reparto
    
    If EsDesdeRecalcular Then
        T = "h"
    Else
        T = ""
    End If
    

    SQL = "SELECT timporteD AS SD, timporteH AS SH, codmacta,"
    SQL = SQL & " fechaent, numdiari, numasien, " & T & "linapu.codccost, idsubcos"
    SQL = SQL & " FROM " & T & "linapu,cabccost WHERE cabccost.codccost=" & T & "linapu.codccost"
    'SQL = SQL & " GROUP BY codmacta, fechaent, numdiari, numasien, codccost"
    SQL = SQL & " AND numdiari=" & numdiari
    SQL = SQL & " AND fechaent='" & Fecha & "'"
    SQL = SQL & " AND numasien=" & Numasiento
    SQL = SQL & " AND " & T & "linapu.codccost Is Not Null;"
    
    
    
    
    
'    SQL = "SELECT Sum(timporteD) AS SD, Sum(timporteH) AS SH, codmacta,"
'    SQL = SQL & " fechaent, numdiari, numasien, " & T & "linapu.codccost, idsubcos"
'    SQL = SQL & " FROM " & T & "linapu,cabccost WHERE cabccost.codccost=" & T & "linapu.codccost"
'    SQL = SQL & " GROUP BY codmacta, fechaent, numdiari, numasien, codccost"
'    SQL = SQL & " HAVING (((numdiari)=" & NumDiari
'    SQL = SQL & ") AND ((fechaent)='" & Fecha & "'"
'    SQL = SQL & " ) AND ((numasien)=" & NumAsiento
'    SQL = SQL & ") AND ((codccost) Is Not Null));"
'
'
    
    
    
    
    
    
    
    RL.Open SQL, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RL.EOF
        cuenta = RL!Codmacta
        CCost = RL!CodCCost
        ImporteD = DBLet(RL!sD, "N")
        ImporteH = DBLet(RL!sH, "N")
        Reparto = (RL!idsubcos = 1)
        If Not CalcularSaldosAnal Then
            RL.Close
            Exit Function
        End If
        If Reparto Then
            If Not HacerReparto(True) Then
                RL.Close
                Exit Function
            End If
        End If
        'Sig
        RL.MoveNext
    Wend
    RL.Close
    IncrementaProgres 2
    CalcularLineasYSaldos = True
End Function




Private Function HacerReparto(Actualizar As Boolean) As Boolean
Dim RR As ADODB.Recordset
Dim AD As Currency
Dim AH As Currency
Dim TD As Currency
Dim TH As Currency
Dim b As Boolean

    HacerReparto = False
    TD = ImporteD
    TH = ImporteH
    AD = 0
    AH = 0
    Set RR = New ADODB.Recordset
    SQL = "Select * from linccost WHERE codccost = '" & CCost & "'"
    RR.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RR.EOF
        'Cargamos los porcentajes
        CCost = RR!subccost
        ImporteD = (RR!porccost) / 100
        ImporteH = ImporteD
        'Importe porcentajeado
        ImporteD = Round(ImporteD * TD, 2)
        ImporteH = Round(ImporteH * TH, 2)
        'Movemos al sguiente
        RR.MoveNext
        'Por si acaso los decimales quedan sueltos entonces
        'Los valores para el ultimo subcentro de reaparto se obtienen por diferencias
        'con el acumulado
        If RR.EOF Then
            ImporteD = TD - AD
            ImporteH = TH - AH
        Else
            'Acumulo
            AD = AD + ImporteD
            AH = AH + ImporteH
        End If
        If Actualizar Then
            b = CalcularSaldosAnal
        Else
            b = CalcularSaldosAnalDesactualizar
        End If
        If Not b Then
            RR.Close
            Exit Function
        End If
    Wend
    RR.Close
    HacerReparto = True
End Function









'/////////////////////////////////////////////////
'//
'//
'//     Calcula los saldos del asiento desde las facturas
'//     Estoes, el asiento esta ya en hco, con lo cual las tablas son de hco
Private Function CalcularLineasYSaldosFacturas() As Boolean
    Dim Reparto As Boolean
    Dim RL As Recordset
    Set RL = New ADODB.Recordset
    
    CalcularLineasYSaldosFacturas = False
    '------------------------------------------
    'SALDOS
'    CalcularLineasYSaldos = False
'    SQL = "SELECT timporteD , timporteH , codmacta"
'    SQL = SQL & " From linapu"
'    SQL = SQL & " WHERE linapu.numdiari = " & NumDiari
'    SQL = SQL & " AND linapu.fechaent='" & Fecha & "'"
'    SQL = SQL & " AND linapu.numasien=" & NumAsiento
'    SQL = SQL & ";"
'
    
    'Antiguo: 27 Febrero
'    SQL = "SELECT Sum(hlinapu.timporteD) AS SD, Sum(hlinapu.timporteH) AS SH, hlinapu.codmacta"
'    SQL = SQL & " , hlinapu.numdiari, hlinapu.fechaent, hlinapu.numasien"
'    SQL = SQL & " From hlinapu"
'    SQL = SQL & " GROUP BY hlinapu.codmacta, hlinapu.numdiari, hlinapu.fechaent, hlinapu.numasien"
'    SQL = SQL & " HAVING (((hlinapu.numdiari)= " & NumDiari
'    SQL = SQL & ") AND ((hlinapu.fechaent)='" & Fecha & "'"
'    SQL = SQL & ") AND ((hlinapu.numasien)=" & NumAsiento
'    SQL = SQL & "));"
    
    
    'Abril 2004. Objetivo : QUITAR GROUP BY
    SQL = "SELECT hlinapu.timporteD AS SD, hlinapu.timporteH AS SH, hlinapu.codmacta"
    'SQL = SQL & " , hlinapu.numdiari, hlinapu.fechaent, hlinapu.numasien"
    SQL = SQL & " From hlinapu"
    SQL = SQL & " WHERE (((hlinapu.numdiari)= " & numdiari
    SQL = SQL & ") AND ((hlinapu.fechaent)='" & Fecha & "'"
    SQL = SQL & ") AND ((hlinapu.numasien)=" & Numasiento
    SQL = SQL & "));"
    
   
    Set RL = New ADODB.Recordset
    RL.Open SQL, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RL.EOF
        cuenta = RL!Codmacta
        If IsNull(RL!sD) Then
            ImporteD = 0
        Else
            'ImporteD = RL!tImporteD
            ImporteD = RL!sD
        End If
        If IsNull(RL!sH) Then
            ImporteH = 0
        Else
            'ImporteH = RL!tImporteH
            ImporteH = RL!sH
        End If
        
        If Not CalcularSaldos Then
            RL.Close
            Exit Function
        End If
        
        'Sig
        RL.MoveNext
    Wend
    RL.Close
    IncrementaProgres 3
    If Not vEmpresa.TieneAnalitica Then
        'NO tiene analitica
        CalcularLineasYSaldosFacturas = True
        Exit Function
    End If
    
    
    '------------------------------------------
    '       ANALITICA
    SQL = "SELECT hlinapu.timporteD AS SD, hlinapu.timporteH AS SH, hlinapu.codmacta,"
    SQL = SQL & " hlinapu.fechaent, hlinapu.numdiari, hlinapu.numasien, hlinapu.codccost,idsubcos"
    SQL = SQL & " From hlinapu,cabccost WHERE cabccost.codccost=hlinapu.codccost"
    SQL = SQL & " AND hlinapu.numdiari =" & numdiari
    SQL = SQL & " AND hlinapu.fechaent='" & Fecha & "'"
    SQL = SQL & " AND hlinapu.numasien=" & Numasiento
    SQL = SQL & " AND hlinapu.codccost Is Not Null;"
    RL.Open SQL, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RL.EOF
        cuenta = RL!Codmacta
        CCost = RL!CodCCost
        ImporteD = DBLet(RL!sD, "N")
        ImporteH = DBLet(RL!sH, "N")
        Reparto = (RL!idsubcos = 1)
        If Not CalcularSaldosAnal Then
            RL.Close
            Exit Function
        End If
        'Sig
        
        If Reparto Then
            If Not HacerReparto(True) Then
                RL.Close
                Exit Function
            End If
        End If
        RL.MoveNext
    Wend
    RL.Close
    IncrementaProgres 2
    CalcularLineasYSaldosFacturas = True
End Function




Private Function InsertarLineas() As Boolean
On Error Resume Next
    SQL = "INSERT INTO hlinapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, ampconce, timporteD, timporteH, codccost, ctacontr, idcontab,punteada)"
    SQL = SQL & " SELECT numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, ampconce, timporteD, timporteH, codccost, ctacontr, idcontab,punteada From linapu"
    SQL = SQL & " WHERE numasien = " & Numasiento
    SQL = SQL & " AND numdiari = " & numdiari
    SQL = SQL & " AND fechaent='" & Fecha & "'"
    ConnConta.Execute SQL
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        InsertarLineas = False
    Else
        InsertarLineas = True
    End If
End Function


Private Function InsertarLineasApuntes() As Boolean
On Error Resume Next
    SQL = "INSERT INTO linapu (numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, ampconce, timporteD, timporteH, codccost, ctacontr, idcontab,punteada)"
    SQL = SQL & " SELECT numdiari, fechaent, numasien, linliapu, codmacta, numdocum, codconce, ampconce, timporteD, timporteH, codccost, ctacontr, idcontab,punteada From hlinapu"
    SQL = SQL & " WHERE numasien = " & Numasiento
    SQL = SQL & " AND numdiari = " & numdiari
    SQL = SQL & " AND fechaent='" & Fecha & "'"
    conn.Execute SQL
    If Err.Number <> 0 Then
         'Hay error , almacenamos y salimos
        InsertarLineasApuntes = False
    Else
        InsertarLineasApuntes = True
    End If
End Function


Private Function CalcularSaldos() As Boolean
    Dim I As Integer
    CalcularSaldos = False
    For I = vEmpresa.numnivel To 1 Step -1
        If Not CalcularSaldos1Nivel(I) Then Exit Function
    Next I
    CalcularSaldos = True
End Function



Private Function CalcularSaldos1Nivel(Nivel As Integer) As Boolean
    Dim ImpD As Double
    Dim ImpH As Double
    Dim TD As String
    Dim TH As String
    Dim cta As String
    Dim I As Integer
    
    
    CalcularSaldos1Nivel = False
    I = DigitosNivel(Nivel)
    If I < 0 Then Exit Function
    
    cta = Mid(cuenta, 1, I)
    SQL = "Select Impmesde,impmesha from hsaldos where "
    SQL = SQL & " Codmacta = '" & cta & "' AND Anopsald = " & Anyo & " AND mespsald = " & mes
    Set RS = New ADODB.Recordset
    RS.Open SQL, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RS.EOF Then
        I = 0   'Nuevo
        ImpD = 0
        ImpH = 0
    Else
        I = 1
        ImpD = RS.Fields(0)
        ImpH = RS.Fields(1)
    End If
    RS.Close
    
    'Acumulamos
    ImpD = ImpD + ImporteD
    ImpH = ImpH + ImporteH
    
    TD = TransformaComasPuntos(CStr(ImpD))
    TH = TransformaComasPuntos(CStr(ImpH))
    If I = 0 Then
        'Nueva insercion
        SQL = "INSERT INTO hsaldos VALUES('" & cta & "'," & Anyo & "," & mes & "," & TD & "," & TH & ")"
        Else
        SQL = "UPDATE hsaldos SET Impmesde=" & TD & ", Impmesha = " & TH
        SQL = SQL & " WHERE Codmacta = '" & cta & "' AND Anopsald = " & Anyo & " AND mespsald = " & mes
    End If
    ConnConta.Execute SQL
    CalcularSaldos1Nivel = True
End Function

'-------------------------------------------------------
'-------------------------------------------------------
'ANALITICA
'-------------------------------------------------------
'-------------------------------------------------------

Private Function CalcularSaldosAnal() As Boolean
    
    CalcularSaldosAnal = CalcularSaldos1NivelAnal(vEmpresa.numnivel)

End Function

Private Function CalcularSaldosAnalDesactualizar() As Boolean
    'Dim i As Integer
    'CalcularSaldosAnalDesactualizar = False
    'For i = vEmpresa.numnivel To 1 Step -1
    CalcularSaldosAnalDesactualizar = CalcularSaldos1NivelAnalDesactualizar(vEmpresa.numnivel)

End Function

Private Function CalcularSaldos1NivelAnal(Nivel As Integer) As Boolean
    Dim ImpD As Currency
    Dim ImpH As Currency
    Dim TD As String
    Dim TH As String
    Dim cta As String
    Dim I As Integer
    
    
    CalcularSaldos1NivelAnal = False
    I = DigitosNivel(Nivel)
    If I < 0 Then Exit Function
    
    cta = Mid(cuenta, 1, I)
    SQL = "Select debccost,habccost from hsaldosanal where "
    SQL = SQL & " codccost='" & CCost & "' AND"
    SQL = SQL & " Codmacta = '" & cta & "' AND anoccost = " & Anyo & " AND mesccost = " & mes
    Set RS = New ADODB.Recordset
    RS.Open SQL, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RS.EOF Then
        I = 0   'Nuevo
        ImpD = 0
        ImpH = 0
    Else
        I = 1
        ImpD = RS.Fields(0)
        ImpH = RS.Fields(1)
    End If
    RS.Close
    'Acumulamos
    ImpD = ImpD + ImporteD
    ImpH = ImpH + ImporteH
    TD = TransformaComasPuntos(CStr(ImpD))
    TH = TransformaComasPuntos(CStr(ImpH))
    If I = 0 Then
        'Nueva insercion
        SQL = "INSERT INTO hsaldosanal(codccost,codmacta,anoccost,mesccost,debccost,habccost)"
        SQL = SQL & " VALUES('" & CCost & "','" & cta & "'," & Anyo & "," & mes & "," & TD & "," & TH & ")"
        Else
        SQL = "UPDATE hsaldosanal SET debccost=" & TD & ", habccost = " & TH
        SQL = SQL & " WHERE Codmacta = '" & cta & "' AND Anoccost = " & Anyo & " AND mesccost = " & mes
        SQL = SQL & " AND codccost = '" & CCost & "';"
    End If
    conn.Execute SQL
    CalcularSaldos1NivelAnal = True
End Function



Private Function CalcularSaldos1NivelAnalDesactualizar(Nivel As Integer) As Boolean
    Dim ImpD As Currency
    Dim ImpH As Currency
    Dim TD As String
    Dim TH As String
    Dim cta As String
    Dim I As Integer
    
    CalcularSaldos1NivelAnalDesactualizar = False
    I = DigitosNivel(Nivel)
    If I < 0 Then Exit Function
    
    cta = Mid(cuenta, 1, I)
    SQL = "Select debccost,habccost from hsaldosanal where "
    SQL = SQL & " codccost='" & CCost & "' AND"
    SQL = SQL & " Codmacta = '" & cta & "' AND anoccost = " & Anyo & " AND mesccost = " & mes
    Set RS = New ADODB.Recordset
    RS.Open SQL, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RS.EOF Then
        MsgBox "Error grave. No habia saldos en analitica para la cuenta: " & cta
        RS.Close
        Exit Function
    Else
        I = 1
        ImpD = RS.Fields(0)
        ImpH = RS.Fields(1)
    End If
    RS.Close
    'Acumulamos
    ImpD = ImpD - ImporteD 'Con respecto a ACTUALIZAR CAMBIA EL SIGNO
    ImpH = ImpH - ImporteH
    TD = TransformaComasPuntos(CStr(ImpD))
    TH = TransformaComasPuntos(CStr(ImpH))
    If ImpD = 0 And ImpH = 0 Then
        'Nueva insercion
        SQL = "DELETE FROM hsaldosanal"
        Else
        SQL = "UPDATE hsaldosanal SET debccost=" & TD & ", habccost = " & TH
    End If
    SQL = SQL & " WHERE Codmacta = '" & cta & "' AND Anoccost = " & Anyo & " AND mesccost = " & mes
    SQL = SQL & " AND codccost = '" & CCost & "';"
    conn.Execute SQL
    CalcularSaldos1NivelAnalDesactualizar = True
End Function




Private Function BorrarASiento(EnHistorico As Boolean) As Boolean
    BorrarASiento = False
    
    'Borramos las lineas
    SQL = "Delete from "
    If EnHistorico Then
        SQL = SQL & "hlinapu"
    Else
        SQL = SQL & "linapu"
    End If
    SQL = SQL & " WHERE numasien = " & Numasiento
    SQL = SQL & " AND numdiari = " & numdiari
    SQL = SQL & " AND fechaent='" & Fecha & "'"
    ConnConta.Execute SQL
    
    
    'La cabecera
    SQL = "Delete from "
    If EnHistorico Then
        SQL = SQL & "hcabapu"
    Else
        SQL = SQL & "cabapu"
    End If
    SQL = SQL & " WHERE numdiari =" & numdiari
    SQL = SQL & " AND fechaent='" & Fecha & "'"
    SQL = SQL & " AND numasien=" & Numasiento
    ConnConta.Execute SQL
    
    BorrarASiento = True
End Function



Private Sub Form_Unload(Cancel As Integer)
If NumErrores > 0 Then CerrarFichero
End Sub

Private Sub CerrarFichero()
On Error Resume Next
If NE = 0 Then Exit Sub
Close #NE
If Err.Number <> 0 Then Err.Clear
End Sub


'Esta funcion me servira para actualizar los asientos k
' se generaran desde TESORERIA.
'YA los hemos metido en tmoactualziar
Private Function ObtenerRegistrosParaActualizar() As Boolean
Dim Cad As String
    Label1.Caption = "Prepara proceso."
    Label1.Refresh
    ObtenerRegistrosParaActualizar = False
    'Borramos temporal
    'Conn.Execute "Delete From tmpactualizarError where codusu = " & vUsu.Codigo
    
    Set RS = New ADODB.Recordset
    RS.Open "Select count(*) from tmpActualizar WHERE codusu =" & vUsu.Codigo, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RS.EOF Then
        'NINGUN REGISTTRO A ACTUALIZAR
        Numasiento = 0
    Else
        Numasiento = RS.Fields(0)
    End If
    RS.Close
    If Numasiento = 0 Then
        MsgBox "Ningún asiento para actualizar desde tesoreria.", vbExclamation
        Exit Function
    End If
    
    'Cargamos valores
    If Numasiento < 32000 Then
        CargaProgres CInt(Numasiento)
        INC = 1
    End If
    
    'Ponemos en marcha la peli
    If Numasiento > 20 Then PonerAVI 1
    
    
    
    'Ponemos el form como toca
    Label1.Caption = "Obtener registros actualización."
    lblAsiento.Caption = ""
    Me.Height = 3315
    Me.frame1Asiento.visible = True
    Me.Refresh
    Me.Height = 3315
    Me.Refresh
    
    RS.Open "Select * from tmpactualizar  WHERE codusu =" & vUsu.Codigo, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        IncrementaProgres 1
        'Para poder acceder a ellos desde cualquier sitio
        Numasiento = RS!numasien
        Fecha = Format(RS!FechaEnt, FormatoFecha)
        numdiari = RS!numdiari
        'No esta bloqueado
        'Comprobamos que esta cuadrado
        Cad = RegistroCuadrado
        If Cad <> "" Then
            InsertaError Cad
            'Borramos de tmpactualizar
            Cad = "delete from tmpactualizar where codusu =" & vUsu.Codigo
            Cad = Cad & " AND numdiari =" & RS!numdiari & " AND numasien =" & RS!numasien
            Cad = Cad & " AND fechaent ='" & Format(RS!FechaEnt, FormatoFecha) & "'"
            conn.Execute Cad
        End If
        

        'Siguiente
        RS.MoveNext
    Wend
    RS.Close
    
    
    'ACtualizarRegistros
    ActualizaASientosDesdeTMP

    'Ahora si todo ha ido bien mostraremos datos de las actualizaciones
    'Set Rs = Nothing
    'Set Rs = New ADODB.Recordset
    'SQL = "Select count(*) from tmpactualizar where codusu=" & vUsu.Codigo
    'Rs.Open SQL, connconta, adOpenForwardOnly, adLockOptimistic, adCmdText
    Me.Height = 4965
    frame1Asiento.visible = False
    
    Me.frameResultados.visible = True
    Me.Refresh
    Screen.MousePointer = vbHourglass
    If NumErrores > 0 Then
        Close #NE
        Label7.Caption = "Se han producido errores."
        CargaListAsiento
    Else
       Label7.Caption = "NO se han producido errores."
       Me.Refresh
       ObtenerRegistrosParaActualizar = True
    End If
    
End Function

Private Function BloqAsien() As String
On Error Resume Next
    'Bloqueamos e insertamos
    BloqAsien = ""
    
    If BloquearAsiento(CStr(Numasiento), CStr(numdiari), Fecha) Then
        'Utilizamos una variable existente
        cuenta = "INSERT INTO tmpactualizar (numdiari, fechaent, numasien, codusu) VALUES ("
        cuenta = cuenta & numdiari & ",'"
        cuenta = cuenta & Fecha & "',"
        cuenta = cuenta & Numasiento & ","
        cuenta = cuenta & vUsu.Codigo & ")"
        conn.Execute cuenta
        If Err.Number <> 0 Then
            Err.Clear
            BloqAsien = "Error al insertar temporal"
            desBloquearAsiento CStr(Numasiento), CStr(numdiari), Fecha
        End If
    Else
        BloqAsien = "Error al bloquear el asiento."
    End If
End Function

Private Sub PonerAVI(NumAVI As Integer)
On Error GoTo EPonerAVI
    If NumAVI = 1 Then
        Me.Animation1.Open App.Path & "\actua.avi"
        Me.Animation1.Play
        Me.Animation1.visible = True
    Else
    
    End If
Exit Sub
EPonerAVI:
    MuestraError Err.Number, "Poner Video"
End Sub


Private Function RegistroCuadrado() As String
    Dim Deb As Currency
    Dim hab As Currency
    Dim RSUM As ADODB.Recordset

    'Trabajamos con RS que es global
    RegistroCuadrado = "" 'Todo bien
    
    Set RSUM = New ADODB.Recordset
    SQL = "SELECT Sum(linapu.timporteD) AS SumaDetimporteD, Sum(linapu.timporteH) AS SumaDetimporteH"
    SQL = SQL & " ,linapu.numdiari,linapu.fechaent,linapu.numasien"
    SQL = SQL & " From linapu GROUP BY linapu.numdiari, linapu.fechaent, linapu.numasien "
    SQL = SQL & " HAVING (((linapu.numdiari)=" & numdiari
    SQL = SQL & ") AND ((linapu.fechaent)='" & Fecha
    SQL = SQL & "') AND ((linapu.numasien)=" & Numasiento
    SQL = SQL & "));"
    
    
    
    
    
    
    RSUM.Open SQL, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RSUM.EOF Then
        Deb = DBLet(RSUM.Fields(0), "N")
        'Deb = Round(Deb, 2)
        hab = RSUM.Fields(1)
        'Hab = Round(Hab, 2)
        CCost = ""
    Else
        Deb = 0
        hab = -1
        CCost = "Asiento sin lineas"
    End If
    
    RSUM.Close
    Set RSUM = Nothing
    If Deb <> hab Then
        If CCost = "" Then CCost = "Asiento descuadrado"
        RegistroCuadrado = CCost
    End If

End Function

Private Function InsertaError(ByRef Cadena As String)
Dim vS As String
    'Insertamos en errores
    'Esta lo tratamos con error especifico
    
    On Error Resume Next


        'Insertamos error para ASIENTOS
        vS = numdiari & "|"
        vS = vS & Fecha & "|"
        vS = vS & Numasiento & "|"
        vS = vS & Cadena & "|"
    

    'Modificacion del 10 de marzo
    'Conn.Execute vS
    AñadeError vS
    
    If Err.Number <> 0 Then
        MsgBox "Se ha producido un error." & vbCrLf & Err.Description & vbCrLf & vS
        Err.Clear
    End If
End Function





Private Function ActualizaASientosDesdeTMP()
Dim RT As Recordset


'Para el progress
Numasiento = ProgressBar1.Max
Me.lblAsiento.Caption = "Nº asiento:"
If Numasiento < 3000 Then
    CargaProgres Numasiento * 10
    Else
    CargaProgres 32000
End If
INC = 1


SQL = "Select * from tmpactualizar where codusu=" & vUsu.Codigo
Set RT = New ADODB.Recordset
RT.Open SQL, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
While Not RT.EOF
    Numasiento = RT!numasien
    FechaAsiento = RT!FechaEnt
    numdiari = RT!numdiari
    'Actualiza el asiento
    If ActualizaAsiento = False Then
         desBloquearAsiento CStr(Numasiento), CStr(numdiari), Fecha

    End If

    'Siguiente
    RT.MoveNext
Wend
RT.Close
Set RT = Nothing
End Function







Private Sub BorrarArchivoTemporal()
On Error Resume Next
If Dir(App.Path & "\ErrActua.txt") <> "" Then Kill App.Path & "\ErrActua.txt"
If Err.Number <> 0 Then MuestraError Err.Number, "Borrar fichero temporal"
End Sub


'QUITAR###
Private Function BloquearAsiento(N As String, D As String, F As String) As Boolean

End Function

Private Function desBloquearAsiento(N As String, D As String, F As String) As Boolean

End Function



Private Sub EliminarArchivoErrores()
On Error Resume Next
If Dir(App.Path & "\ErrActua.txt") <> "" Then Kill App.Path & "\ErrActua.txt"
If Err.Number <> 0 Then MuestraError Err.Number, "Borrar fichero temporal"
End Sub

'Devuelve, para un nivel determinado, cuantos digitos tienen las cuentas
' a ese nivel
Public Function DigitosNivel(numnivel As Integer) As Integer
    Select Case numnivel
    Case 1
        DigitosNivel = vEmpresa.numdigi1

    Case 2
        DigitosNivel = vEmpresa.numdigi2

    Case 3
        DigitosNivel = vEmpresa.numdigi3

    Case 4
        DigitosNivel = vEmpresa.numdigi4

    Case 5
        DigitosNivel = vEmpresa.numdigi5

    Case 6
        DigitosNivel = vEmpresa.numdigi6

    Case 7
        DigitosNivel = vEmpresa.numdigi7

    Case 8
        DigitosNivel = vEmpresa.numdigi8

    Case 9
        DigitosNivel = vEmpresa.numdigi9

    Case 10
        DigitosNivel = vEmpresa.numdigi10

    Case Else
        DigitosNivel = -1
    End Select
End Function


