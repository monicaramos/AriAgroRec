VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCarpAridoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Esta es la pregunta"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   Icon            =   "frmCarpAridoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   5730
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6240
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameSelFolder 
      Height          =   6255
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   5655
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   5415
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   9551
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin VB.CommandButton cmdSelFolder 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   4200
         TabIndex        =   2
         Top             =   5760
         Width           =   1095
      End
      Begin VB.CommandButton cmdSelFolder 
         Caption         =   "Seleccionar"
         Height          =   375
         Index           =   0
         Left            =   3000
         TabIndex        =   1
         Top             =   5760
         Width           =   1095
      End
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCarpAridoc.frx":030A
            Key             =   "v_abierto"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCarpAridoc.frx":6B6C
            Key             =   "v_cerrado"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCarpAridoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DatoSeleccionado(CadenaSeleccion As String)
Public Event CadenaSeleccion()
Public ModoTrabajo As Byte  '---------------------
'  -- Modos de Trabajo

Dim CadenaCarpetas As String

Dim Cortar11 As String
Dim pegar11 As String


Public DatosCopiados As String
Public Opcion As Byte
    '1.- Copiar / Mover Archivos
    '2.-   "  / "  CARPETAS
    
    
    '5.- Propiedades de unos archivos
    '6.- Propiedades carpeta
    
    
    '8.- Importes archivos seleccionados
    '9.- Importes carpeta seleccionada
    '10.- Importe subcarpetas
    
    
    '11.- Cambio de propietario para los archivos
    
    '20.- Seleccionar una carpeta para mover archivos
    
    '21.- Direccion e- mail
    
    '22.- Preguna PATH integrador
    
    '23.- Nueva( o modificar) carpeta para las plantillas
    
    '24.- Seleccionar carpeta para agregar mover las plantillas
    
Public origenDestino As String   'Separados con pipes
Private AntiguoCursor As Byte
Private PrimeraVez As Boolean



Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub cmdSelFolder_Click(Index As Integer)
    If Index = 0 Then
        If TreeView1.SelectedItem Is Nothing Then Exit Sub
        
'        If origenDestino = "1" Then
'            'Es para el traspaso a hco. Ademas de la carpeta voy a llevar todas las subcarpetas colgantes
'            origenDestino = CopiaArchivosCarpetaRecursiva(TreeView1.SelectedItem)
'            DatosCopiados = TreeView1.SelectedItem.FullPath & "·" & origenDestino
'        Else
'            '"0"
            DatosCopiados = TreeView1.SelectedItem.Key & "|" & TreeView1.SelectedItem.Text & "|" & TreeView1.SelectedItem.FullPath & "|"
            RaiseEvent DatoSeleccionado(DatosCopiados)
'        End If
            
    End If
    Unload Me
End Sub


Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If Opcion = 20 Then
            TreeView1.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
    PrimeraVez = True
    
    Me.FrameSelFolder.visible = False
    Select Case Opcion
        
    Case 20
        'En origen destino tendremos
        'si donde debo devolver la carpeta es para
        'los resultado o traspaso a hco ....
        '   0.- Resultados
        '   1.- Traspaso a hco
        If origenDestino = "" Then origenDestino = 0
        FrameSelFolder.visible = True
        Me.cmdSelFolder(1).Cancel = True
        CargaArbol
        Caption = "Seleccione una carpeta"
        H = Me.FrameSelFolder.Height
        W = Me.FrameSelFolder.Width
            
    End Select
    
    Me.Height = H + 420
    Me.Width = W + 120
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cerrar
    Screen.MousePointer = AntiguoCursor
End Sub

'Private Sub PonerLabels()
'    Dim C As Long
'
'    'vienen empipados:
'    ' nombre carpeta
'    ' archvios seleccionados, tamañoselecioados
'    ' archivos carpetas ,  tamño total,ocultos
'    '
'    Label6.Caption = RecuperaValor(DatosCopiados, 1)
'    Label8(0).Caption = RecuperaValor(DatosCopiados, 2)
'    Label8(1).Caption = RecuperaValor(DatosCopiados, 3) & " Kb"
'
'    Label8(2).Caption = RecuperaValor(DatosCopiados, 4)
'    'tamaño
'    Label8(3).Caption = RecuperaValor(DatosCopiados, 5) & " Kb"
'    'Coultos
'    C = Val(RecuperaValor(DatosCopiados, 6))
'    If C > 0 Then Label8(2).Caption = Label8(2).Caption & " - Ocultos " & C
'
'
'    Label8(4).Caption = RecuperaValor(DatosCopiados, 7)
'    C = Val(RecuperaValor(DatosCopiados, 8))
'    If C > 0 Then Label8(4).Caption = Label8(4).Caption & " - Ocultos " & C
'
'
'    'Si la opcion es 6
'    C = InStrRev(Label6.Caption, "\")
'    Text1.Text = ""
'    If C > 0 Then
'        Text1.Text = Mid(Label6.Caption, 1, C - 1)
'        Label6.Caption = Mid(Label6.Caption, C + 1)
'    End If
'
'End Sub




Private Sub frmU_DatoSeleccionado(CadenaSeleccion As String)
'Dim C As String
'    Screen.MousePointer = vbHourglass
'    C = "Select grupos.codgrupo,grupos.nomgrupo from usuariosgrupos,grupos where "
'    C = C & "usuariosgrupos.codgrupo =grupos.codgrupo and codusu=" & RecuperaValor(CadenaSeleccion, 1)
'    C = C & " ORDER BY orden"
'
'    Set miRSAux = New ADODB.Recordset
'    miRSAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    C = ""
'    If Not miRSAux.EOF Then
'        If Not IsNull(miRSAux.Fields(1)) Then C = miRSAux.Fields(1)
'    End If
'    miRSAux.Close
'    Set miRSAux = Nothing
'    If C = "" Then
'        MsgBox "Grupo PPal para el usuario: " & CadenaSeleccion & " NO encontrado", vbExclamation
'        Exit Sub
'    End If
'
'    'Llegado aqui, ponemos
'
'    'vC.userprop = Val(RecuperaValor(CadenaSeleccion, 1))
'    'vC.groupprop = Val(C)
'    Text3(0).Text = RecuperaValor(CadenaSeleccion, 3)
'    Text3(0).Tag = RecuperaValor(CadenaSeleccion, 1)
'    Text3(1).Text = C
    
    Screen.MousePointer = vbDefault

End Sub






'Private Sub CargaElArbolDeAmin()
'Dim NodD As Node
'Dim Nod As Node
'Dim i As Integer
'
'    Set TreeView1.ImageList = Admin.TreeView1.ImageList
'
'    'El raiz
'    Set Nod = Admin.TreeView1.Nodes(1)
'    Set NodD = TreeView1.Nodes.Add(, , Nod.Key, Nod.Text, Nod.Image)
'
'    'Insertamos el primero
'    For i = 2 To Admin.TreeView1.Nodes.Count
'        Set Nod = Admin.TreeView1.Nodes(i)
'        Set NodD = TreeView1.Nodes.Add(Nod.Parent.Key, tvwChild, Nod.Key, Nod.Text, Nod.Image)
'
'    Next i
'    TreeView1.Nodes(2).EnsureVisible
'End Sub
'
'



Private Function CopiaArchivosCarpetaRecursiva(No As Node) As String
Dim Nod As Node
Dim J As Integer
Dim i As Integer
Dim C As String

    'Primero copiamos la carpeta
    C = Mid(No.Key, 2) & "|"
        If No.Children > 0 Then
            J = No.Children
            Set Nod = No.Child
            For i = 1 To J
               C = C & CopiaArchivosCarpetaRecursiva(Nod)
               If i <> J Then Set Nod = Nod.Next
            Next i
        End If
    CopiaArchivosCarpetaRecursiva = C
End Function
    

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub


Private Sub TreeView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub CargaArbol()
Dim cad As String
Dim RS As ADODB.Recordset
Dim Nod As Node
Dim C As Integer
Dim i As Integer
Dim Contador2 As Integer


    TreeView1.Nodes.Clear
    TreeView1.ImageList = Me.ImageList3
    
    cad = " from carpetas"
'    If ModoTrabajo = vbHistNue Or ModoTrabajo = vbHistAnt Then cad = cad & "hco"
'    'Es el usuario propietario
'    If vUsu.Codusu > 0 Then
'        cad = cad & " WHERE "
'        cad = cad & "userprop = " & vUsu.Codusu
'
'        'O el grupo tiene permiso
'        cad = cad & " OR (lecturag & " & vUsu.Grupo & ")"
'
'    End If
'
'    If ModoTrabajo = vbHistNue Or ModoTrabajo = vbHistAnt Then
'        If vUsu.Codusu = 0 Then
'            cad = cad & " WHERE "
'        Else
'            cad = cad & " AND "
'        End If
'        cad = cad & "codequipo = " & vUsu.PC
'    End If
    
    
    'Ordenado por padre
    cad = cad & " ORDER BY Padre,nombre"
    
    
    Set RS = New ADODB.Recordset
    RS.Open "select * " & cad, ConnAridoc, adOpenKeyset, adLockPessimistic, adCmdText
    If RS.EOF Then
        MsgBox "ERROR GRAVE cargando árbol de directorios(Situacion: 1)", vbCritical
        End
    End If
    CadenaCarpetas = "|"
    
    If RS!padre <> 0 Then
        MsgBox "Error en primer NODO. Padre != 0", vbExclamation
        End
    End If
    C = 0
    i = 0
    While i = 0
        INSERTAR_NODO RS, 1
        RS.MoveNext
        If RS.EOF Then
            i = 1
        Else
            If RS!padre <> 0 Then i = 1
        End If
        C = C + 1
    Wend
    
    'Cargo el segundo nivel
    Contador2 = TreeView1.Nodes.Count
    C = 0
    For i = 1 To Contador2
        cad = Mid(TreeView1.Nodes(i).Key, 2)
        RS.MoveFirst
        RS.Find " padre = " & cad, , adSearchForward, 1
        While Not RS.EOF
            C = C + 1
            If RS!padre = cad Then
                INSERTAR_NODO RS, 2
            Else
                RS.MoveLast
                
            End If
            RS.MoveNext
        Wend
    Next i
       
    If C > 0 Then
'                If Not PrimeraVez Then Label3.Caption = "     c   a   r   g   a   n   d   o  "
                'Cargo el tercer nivel
                C = Contador2 + 1
                Contador2 = TreeView1.Nodes.Count
                For i = C To Contador2
                    cad = Mid(TreeView1.Nodes(i).Key, 2)
                    RS.MoveFirst
                    RS.Find " padre = " & cad, , adSearchForward, 1
                    While Not RS.EOF
                        C = C + 1
                        If RS!padre = cad Then
                            INSERTAR_NODO RS, 2
                        Else
                            RS.MoveLast
                        End If
                        RS.MoveNext
                    Wend
                Next i
                
   
                C = Contador2 + 1
                Contador2 = TreeView1.Nodes.Count
                If Contador2 >= C Then
                    For i = C To Contador2
                        
                        CargaArbolRecursivo Mid(TreeView1.Nodes(i).Key, 2), RS, 5
                      
                    Next i
                End If
                    
                        
    End If
    
    
        
    RS.Close
'    If Not PrimeraVez Then Label3.Caption = " AriDoc: Gestión documental"
    If TreeView1.Nodes.Count > 2 Then TreeView1.Nodes(3).EnsureVisible
   
End Sub


Private Function INSERTAR_NODO(ByRef RSS As Recordset, SubNivel As Integer) As Integer
Dim XNodo As Node

On Error GoTo EIns_Nodo

    
    

    INSERTAR_NODO = -1
    If RSS!padre = 0 Then
        'NODO RAIZ
        Set XNodo = TreeView1.Nodes.Add(, tvwChild, "C" & RSS!codcarpeta)
    Else
    
        'NODO HIJO
        Set XNodo = TreeView1.Nodes.Add("C" & RSS!padre, tvwChild, "C" & RSS!codcarpeta)
    End If
    
    XNodo.Text = RSS!Nombre
    'En el tag metemos la seguriad
    XNodo.Tag = RSS!escriturau & "|" & RSS!escriturag & "|"
    
    CadenaCarpetas = CadenaCarpetas & Mid(XNodo.Key, 2) & "|"
    
    
    XNodo.Image = "v_cerrado"
    XNodo.ExpandedImage = "v_abierto"
'    If SubNivel > 4 Then
'        If Not XNodo.Expanded Then
'            XNodo.Image = "falta"
'            XNodo.ExpandedImage = "falta"
'        End If
'    Else
    If RSS!hijos > 0 Then INSERTAR_NODO = XNodo.Index
'    End If
Exit Function
EIns_Nodo:
    Cortar11 = "ERROR GRAVE." & vbCrLf & vbCrLf
    Cortar11 = Cortar11 & Err.Description & vbCrLf & vbCrLf
    Cortar11 = Cortar11 & RSS!codcarpeta & " " & DBLet(RSS!Nombre, "T")
   ' MsgBox Cortar11, vbCritical
    Cortar11 = Cortar11 & vbCrLf & vbCrLf
    Cortar11 = Cortar11 & "Verifique ARIDOC. Si persiste avise a soporte técnico"
    Cortar11 = Cortar11 & vbCrLf & vbCrLf & vbCrLf & "¿FINALIZAR?"
    If MsgBox(Cortar11, vbCritical + vbYesNo) = vbYes Then
        Conn.Close
        End
    End If
End Function


Private Sub CargaArbolRecursivo(CarpePadre As String, ByRef rs1 As ADODB.Recordset, ByVal Nivel As Integer)
Dim C As Integer
Dim i As Integer
Dim cadena As String
Dim Fin As Boolean
 
    'Este esta puesto para cuando es el arranque, que si le cuesta leer que no
    'bloquee el equipo
    If (TreeView1.Nodes.Count Mod 30) = 0 Then DoEvents


    cadena = ""
    C = 0
    rs1.MoveFirst
    rs1.Find " padre = " & CarpePadre, , adSearchForward, 1
    Fin = rs1.EOF
    While Not Fin
        If rs1!padre = CarpePadre Then
        
            i = INSERTAR_NODO(rs1, Nivel)
            If i > 0 Then
                cadena = cadena & rs1!codcarpeta & "|"
                C = C + 1
            End If
            rs1.MoveNext
            If rs1.EOF Then Fin = True
        Else
            Fin = True
        End If
'--monica
'        If Timer - T2 > 1 Then
'            If PrimeraVez Then
'                frmInicio.Label1(2).visible = Not frmInicio.Label1(2).visible
'                frmInicio.Label1(2).Refresh
'
'            Else
'                If Label3.Caption = "" Then
'                    Label3.Caption = "     c   a   r   g   a   n   d   o  "
'                Else
'                    Label3.Caption = ""
'                End If
'                Label3.Refresh
'            End If
'            T2 = Timer
'        End If
    Wend

    If C > 0 Then
        For i = 1 To C
            CargaArbolRecursivo (RecuperaValor(cadena, i)), rs1, Nivel + 1
        Next i
    End If

End Sub




