VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVerErrores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Errores en Albaranes"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
   Icon            =   "frmVerErrores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   5460
      TabIndex        =   4
      Top             =   5400
      Width           =   1035
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3915
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   6906
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   520
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   "Errores en Formas de Pago"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   975
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "FECHAS:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6135
   End
End
Attribute VB_Name = "frmVerErrores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: LAURA    +-+-
' +-+- Fecha: 03/05/06 +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-+

Option Explicit


'Proveedor del que mostramos las no conformidades
Public desdefec As String
Public hastafec As String

Private Sub cmdCerrar_Click()
    'salimos y no hacemos nada
    Unload Me
End Sub

Private Sub Form_Activate()
    PonerFocoBtn Me.cmdCerrar
End Sub

Private Sub Form_Load()
    Label1(1).Caption = "Del " & Format(desdefec, "dd/mm/yyyy") & " al " & Format(hastafec, "dd/mm/yyyy")
    CargarListView
End Sub

Private Sub CargarListView()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim ItmX As ListItem

    On Error GoTo ECargar

    'Los encabezados
    ListView1.ColumnHeaders.Clear
    Me.ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Add , , "Albarán", 1000
    ListView1.ColumnHeaders.Add , , "Fecha", 1100
    ListView1.ColumnHeaders.Add , , "Nombre Cliente", 4100

    
    SQL = "SELECT numalbar,fecalbar,nomsocio "
    SQL = SQL & " FROM scaalb, ssocio, sforpa "
    SQL = SQL & " WHERE scaalb.codsocio = ssocio.codsocio AND scaalb.codforpa = sforpa.codforpa"
    SQL = SQL & " AND scaalb.codsocio = 0 AND sforpa.tipforpa = 4"
    SQL = SQL & " ORDER BY numalbar "
    
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        Set ItmX = ListView1.ListItems.Add
        ItmX.Text = Format(RS.Fields(0).Value)
      
        ItmX.SubItems(1) = RS.Fields(1).Value
        ItmX.SubItems(2) = RS.Fields(2).Value
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    

ECargar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar Errores.", Err.Description
End Sub

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
         Unload Me
    End If
End Sub
