VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMensajes2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mensajes"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11055
   Icon            =   "frmMensajes2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameHcoFrasPozos 
      Height          =   5790
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   10950
      Begin VB.CommandButton cmdCerrarFras 
         Caption         =   "Continuar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   9135
         TabIndex        =   3
         Top             =   5130
         Width           =   1575
      End
      Begin MSComctlLib.ListView ListView23 
         Height          =   4155
         Left            =   240
         TabIndex        =   1
         Top             =   750
         Width           =   10485
         _ExtentX        =   18494
         _ExtentY        =   7329
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Variedad"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Clase "
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripcion"
            Object.Width           =   3706
         EndProperty
      End
      Begin VB.Label Label25 
         Caption         =   "Histórico de Consumo del Contador"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   270
         TabIndex        =   2
         Top             =   270
         Width           =   7980
      End
   End
End
Attribute VB_Name = "frmMensajes2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'====================== VBLES PUBLICAS ================================

Public Event DatoSeleccionado(CadenaSeleccion As String)

Public OpcionMensaje As Byte

'1 .- Historico de recibos de un contador (Monasterios)





Public cadWHERE As String 'Cadena para pasarle la WHERE de la SELECT de los cobros pendientes o de Pedido(para comp. stock)
                          'o CodArtic para seleccionar los Nº Series
                          'para cargar el ListView
                          
Public cadWHERE2 As String
Public campo As String
Public cadena As String ' sql para cargar el listview
Public vCampos As String 'Articulo y cantidad Empipados para Nº de Series
                         'Tambien para pasar el nombre de la tabla de lineas (sliped, slirep,...)
                         'Dependiendo desde donde llamemos, de Pedidos o Reparaciones
Public desdeHco As Boolean

'====================== VBLES LOCALES ================================

Dim PulsadoSalir As Boolean 'Solo salir con el boton de Salir no con aspa del form
Dim PrimeraVez As Boolean

'Para los Nº de Serie
Dim TotalArray As Integer
Dim codArtic() As String
Dim cantidad() As Integer


Dim CadContadores As String

Dim nomColumna As String
Dim nomColumna2 As String
Dim columna As Integer
Dim Columna2 As Integer
Dim Orden As Integer
Dim Orden2 As Integer
Dim PrimerCampo As Integer

Dim vAnt As Integer



Private Sub cmdCerrarFras_Click(Index As Integer)
    Unload Me
End Sub


Private Sub Form_Activate()
Dim OK As Boolean

    
    Select Case OpcionMensaje
        Case 1
            CargarFrasConsumoPozos
            
            PonerFocoBtn cmdCerrarFras(0)
            
    End Select
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim Cad As String
On Error Resume Next
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    
    '[Monica]23/10/2017: hco de facturas del contador (Monasterios)
    Me.FrameHcoFrasPozos.visible = False
    
    PulsadoSalir = True
    PrimeraVez = True
    
    
    
    Select Case OpcionMensaje
    
        Case 1 ' histórico de facturas de consumo del contador
        
            Label25.Caption = "Histórico de Consumo del Contador " & cadena
            Me.Refresh
            DoEvents
        
            H = Me.FrameHcoFrasPozos.Height
            W = Me.FrameHcoFrasPozos.Width
            PonerFrameVisible FrameHcoFrasPozos, True, H, W
    
            Me.Left = 200
            Me.Top = 2850
            DoEvents
    End Select
    'Me.cmdCancel(indFrame).Cancel = True
    Me.Height = H + 350
    Me.Width = W + 70
    
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub Form_Unload(Cancel As Integer)
    If PulsadoSalir = False Then Cancel = 1
    If OpcionMensaje = 49 And vCampos <> "1" Then vCampos = "0"

End Sub



Private Function ObtenerTamanyosArray() As Boolean
'Para el frame de los Nº de Serie de los Articulos
'En cada indice pone en CodArtic(i) el codigo del articulo
'y en Cantidad(i) la cantidad solicitada de cada codartic
Dim I As Integer, J As Integer

    ObtenerTamanyosArray = False
    'Primero a los campos de la tabla
    TotalArray = -1
    J = 0
    Do
        I = J + 1
        J = InStr(I, vCampos, "·")
        If J > 0 Then TotalArray = TotalArray + 1
    Loop Until J = 0
    
    If TotalArray < 0 Then Exit Function
    
    'Las redimensionaremos
    ReDim codArtic(TotalArray)
    ReDim cantidad(TotalArray)
    
    ObtenerTamanyosArray = True
End Function


Private Function SeparaCampos() As Boolean
'Para el frame de los Nº de Serie de los Articulos
Dim Grupo As String
Dim I As Integer
Dim J As Integer
Dim c As Integer 'Contador dentro del array

    SeparaCampos = False
    I = 0
    c = 0
    Do
        J = I + 1
        I = InStr(J, vCampos, "·")
        If I > 0 Then
            Grupo = Mid(vCampos, J, I - J)
            'Y en la martriz
            InsertaGrupo Grupo, c
            c = c + 1
        End If
    Loop Until I = 0
    SeparaCampos = True
End Function


Private Sub InsertaGrupo(Grupo As String, Contador As Integer)
Dim J As Integer
Dim Cad As String

    J = 0
    Cad = ""
    
    'Cod Artic
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        Cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
        J = 1
    End If
    codArtic(Contador) = Cad
    
    'Cantidad
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        Cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
    Else
        Cad = Grupo
        Grupo = ""
    End If
    cantidad(Contador) = Cad
End Sub









Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub CargarFrasConsumoPozos()
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim It As ListItem

    SQL = "select * from rrecibpozos where hidrante = " & DBSet(cadena, "T")
    SQL = SQL & " and codtipom = 'RCP' "
    SQL = SQL & " order by fech_act desc "
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ListView23.ColumnHeaders.Clear

    ListView23.ColumnHeaders.Add , , "Fecha Actual", 2000
    ListView23.ColumnHeaders.Add , , "Lectura", 1900, 1
    ListView23.ColumnHeaders.Add , , "Fecha Anterior", 2000, 1
    ListView23.ColumnHeaders.Add , , "Lectura", 1900, 1
    ListView23.ColumnHeaders.Add , , "Consumo", 2200, 1
    
    ListView23.ListItems.Clear
    
    ListView23.SmallIcons = frmPpal.imgListPpal
    
    
    TotalArray = 0
    While Not Rs.EOF
        Set It = ListView23.ListItems.Add
            
            
        'It.Tag = DevNombreSQL(RS!codCampo)
        It.Text = DBLet(Rs!fech_act, "F")
        It.SubItems(1) = DBLet(Rs!lect_act, "N")
        It.SubItems(2) = DBLet(Rs!fech_ant, "F")
        It.SubItems(3) = DBLet(Rs!lect_ant, "N")
        It.SubItems(4) = Format(DBLet(Rs!Consumo, "N"), "###,###,##0")
        
        
        Rs.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    Rs.Close
    
End Sub



