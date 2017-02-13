VERSION 5.00
Begin VB.Form frmTrzManPalet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manejo de palets"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5760
   Icon            =   "frmTrzManPalet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDesAsign 
      Caption         =   "Desasignar la tarjeta del palet sin mas"
      Height          =   495
      Left            =   510
      TabIndex        =   11
      Top             =   2760
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resultado consulta "
      Height          =   1485
      Left            =   240
      TabIndex        =   3
      Top             =   870
      Width           =   5175
      Begin VB.TextBox txtResul 
         Appearance      =   0  'Flat
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   4935
      End
   End
   Begin VB.TextBox txtLinConf 
      Height          =   285
      Left            =   1410
      TabIndex        =   8
      Top             =   930
      Width           =   735
   End
   Begin VB.TextBox txtFecHora 
      Height          =   285
      Left            =   3690
      TabIndex        =   7
      Top             =   930
      Width           =   1455
   End
   Begin VB.CommandButton cmdAbocar 
      Caption         =   "Abocar el palet por la linea, fecha y hora indicada."
      Height          =   495
      Left            =   450
      TabIndex        =   6
      Top             =   1410
      Width           =   4695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3510
      TabIndex        =   5
      Top             =   3900
      Width           =   1695
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "Consultar"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtRFID 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Linea Conf.:"
      Height          =   255
      Left            =   450
      TabIndex        =   10
      Top             =   930
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha y Hora:"
      Height          =   255
      Left            =   2250
      TabIndex        =   9
      Top             =   930
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "COD.TARJETA:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmTrzManPalet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const IdPrograma = 9009


Dim IdPalet As Long
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset

Private Sub cmdAbocar_Click()
    Dim resultado As Boolean
    If IdPalet = 0 Then
        MsgBox "Debe consultar previamente un palet"
        Exit Sub
    End If
    If txtLinConf = "" Or (Not IsNumeric(txtLinConf)) Then
        MsgBox "Debe introducir un valor correcto para la l�nea de confecci�n"
        Exit Sub
    End If
    If (Not IsDate(txtFecHora)) Then
        MsgBox "Introduzca una fecha y hora v�lida"
        Exit Sub
    End If
    resultado = CargaLineaConfeccion(Val(txtLinConf), txtRFID, CDate(txtFecHora))
    If Not resultado Then
        MsgBox "Se ha producido una incidencia en el abocamiento"
    End If
    MsgBox "Palet abocado, tarjeta desasignada"
    IdPalet = 0
    txtResul = ""
    txtRFID.SetFocus
    
End Sub

Private Sub cmdConsultar_Click()
    'consultar el palet con tarjeta
    If txtRFID = "" Then
        MsgBox "Debe introducir algun valor en el campo de tarjeta"
        Exit Sub
    End If
    
    SQL = "select * from trzpalets where CRFID ='" & txtRFID & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        IdPalet = Rs!IdPalet '????? rafa
        'SQL = "select * from Palets where codPalet = " & CStr(IdPalet)
        SQL = "select nomvarie from variedades where codvarie = " & DBSet(Rs!codvarie, "N")
        Set RS1 = New ADODB.Recordset
        RS1.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        SQL = "Palet:" & CStr(IdPalet) & " Fecha:" & Format(Rs!Fecha, "dd/mm/yyyy") & " Partida = " & CStr(Rs!numnotac) & vbCrLf
        SQL = SQL & "Socio:" & CStr(Rs!Codsocio) & " Campo:" & CStr(Rs!codcampo) & " Cajones:" & CStr(Rs!NumCajones) & " Kilos:" & CStr(Rs!Numkilos) & vbCrLf
        SQL = SQL & "Variedad:" & RS1!nomvarie
        'SQL = SQL & "Producto:" & rs!NomProdu & " Variedad:" & rs!nomvarie
        
        Set RS1 = Nothing
        txtResul.Text = SQL
    Else
        IdPalet = 0
        txtResul.Text = "NO HAY NINGUN PALET CON ESTA TARJETA ASOCIADA"
    End If
    
    Set Rs = Nothing
    
End Sub

Private Sub cmdDesAsign_Click()
    If IdPalet = 0 Then
        MsgBox "Debe consultar previamente un palet"
        Exit Sub
    End If
    
    SQL = "update trzpalets set CRFID = null where IdPalet = " & CStr(IdPalet)
    conn.Execute SQL
    
    MsgBox "Tarjeta desasignada"
    
    IdPalet = 0
    txtResul = ""
    txtRFID.SetFocus
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    IdPalet = 0
End Sub

Private Sub txtRFID_Change()
    txtResul = ""
End Sub

Private Function CargaLineaConfeccion(Linea As Long, RFID As String, FechaHora As Date) As Boolean
    Dim SQL As String
    Dim Rs As ADODB.Recordset
    Dim resultaddo As Boolean
    
    '-- Buscamos que el palet exista
    SQL = "select * from trzpalets where CRFID = '" & RFID & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then
        MsgBox "El palet RFID: " & RFID & " leido en la l�nea " & CStr(Linea) & " no existe en el sistema"
        Exit Function
    End If
    SQL = "insert into trzlineas_cargas(linea,idpalet,fechahora,fecha,tipo) values("
    SQL = SQL & DBSet(Linea, "N") & ","
    SQL = SQL & DBSet(Rs!IdPalet, "N") & ","
    SQL = SQL & DBSet(FechaHora, "FH") & ","
    SQL = SQL & DBSet(FechaHora, "F") & ","
    SQL = SQL & DBSet(Rs!Tipo, "N") & ")"
    conn.Execute SQL
    
    SQL = "update trzpalets set CRFID = null where IdPalet=" & DBSet(Rs!IdPalet, "N")
    conn.Execute SQL
    
    CargaLineaConfeccion = True
End Function

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub
