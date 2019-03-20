VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#17.2#0"; "Codejock.ShortcutBar.v17.2.0.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#17.2#0"; "Codejock.Controls.v17.2.0.ocx"
Begin VB.Form frmPaneContacts2 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TreeView tree 
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4095
      _Version        =   1114114
      _ExtentX        =   7223
      _ExtentY        =   7435
      _StockProps     =   77
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Scroll          =   0   'False
      Appearance      =   6
      IconSize        =   16
      ShowBorder      =   0   'False
   End
   Begin XtremeShortcutBar.ShortcutCaption MainCaption 
      Height          =   360
      Left            =   4320
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   360
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
      Expandable      =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption ItemCaption 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   4695
      _Version        =   1114114
      _ExtentX        =   8281
      _ExtentY        =   503
      _StockProps     =   14
      Caption         =   "Empresas sistema"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      VisualTheme     =   3
      Expandable      =   -1  'True
   End
End
Attribute VB_Name = "frmPaneContacts2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim Cad As String


    Set tree.Icons = frmShortBar.wndShortcutBar.Icons
    tree.IconSize = 16
    
    tree.Font.SIZE = 10
    
     
    
    

    BuscaEmpresas
    
    
    
    UpdateLayout
End Sub

Public Sub BuscaEmpresas()
Dim Prohibidas As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Cad As String
Dim SQL As String

Dim N

'Cargamos las prohibidas
Prohibidas = DevuelveProhibidas

    'Cargamos las empresas
    Set Rs = New ADODB.Recordset
    
    '[Monica]11/04/2014: solo debe de salir las ariconta
    Rs.Open "Select * from usuarios.empresasariconta where conta like 'ariconta%' ORDER BY Codempre", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    tree.Nodes.Clear
    
    While Not Rs.EOF
        Cad = "|" & Rs!codempre & "|"
        If InStr(1, Prohibidas, Cad) = 0 Then
            Cad = Rs!nomempre
            Set N = tree.Nodes.Add(, , CStr("N" & Rs!codempre), Rs!nomresum)   'nomempre
        
            
            
            'ItmX.SubItems(1) = Rs!nomresum
            'Set N = tree.Nodes.Add("NP", tvwChild, "NN" & Rs!codempre, Rs!nomresum)
            ' sacamos las fechas de inicio y fin
            'Sql = "select fechaini, fechafin from " & Trim(Rs!CONTA) & ".parametros"
            'Set Rs2 = New ADODB.Recordset
            'Rs2.Open Sql, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            'If Not Rs2.EOF Then
            '    ItmX.SubItems(2) = Rs2!fechaini & " - " & Rs2!fechafin
            'End If
            'Set Rs2 = Nothing
            
                
            Cad = Rs!CONTA & "|" & Rs!nomresum '& "|" & Rs!Usuario & "|" & Rs!Pass & "|"
            
            If Rs!codempre = vEmpresa.codempre Then
                N.Bold = True
                Set tree.SelectedItem = N
            End If
                
       
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    
    
End Sub


Private Function DevuelveProhibidas() As String
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim i As Integer
    On Error GoTo EDevuelveProhibidas
    DevuelveProhibidas = ""
    Set Rs = New ADODB.Recordset
    i = vUsu.Codigo Mod 1000
    Rs.Open "Select * from usuarios.usuarioempresasariconta WHERE codusu =" & i, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Cad = ""
    While Not Rs.EOF
        Cad = Cad & Rs.Fields(1) & "|"
        Rs.MoveNext
    Wend
    If Cad <> "" Then Cad = "|" & Cad
    Rs.Close
    DevuelveProhibidas = Cad
EDevuelveProhibidas:
    Err.Clear
    Set Rs = Nothing
End Function









Public Sub SetFlatStyle(FlatStyle As Boolean)
      
    Me.BackColor = frmShortBar.wndShortcutBar.PaintManager.PaneBackgroundColor
    tree.BackColor = Me.BackColor
    tree.ForeColor = frmShortBar.wndShortcutBar.PaintManager.PaneTextColor
    
    MainCaption.GradientColorDark = frmShortBar.wndShortcutBar.PaintManager.PaneBackgroundColor
    MainCaption.GradientColorLight = frmShortBar.wndShortcutBar.PaintManager.PaneBackgroundColor

End Sub

Private Sub ItemCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ItemCaption.Expanded = Not ItemCaption.Expanded
    UpdateLayout
End Sub

Private Sub Form_Resize()
    ItemCaption.Width = Me.ScaleWidth
    MainCaption.Left = Me.ScaleWidth - MainCaption.Width
    If Me.Height - tree.top > 100 Then Me.tree.Height = Me.Height - tree.top
End Sub


Sub UpdateLayout()

    Dim top As Long
    
    top = ItemCaption.top + ItemCaption.Height
    If ItemCaption.Expanded Then
        tree.visible = True
        tree.top = 80 + top
        top = 80 + top + tree.Height
    Else
        tree.visible = False
    End If

End Sub


Private Sub MainCaption_ExpandButtonClicked()
    Call frmppal.ExpandButtonClicked
End Sub


Public Sub SeleccionarNodoEmpresa(QueEmpresa As Integer)
Dim i As Integer
    For i = 1 To tree.Nodes.Count
        If Val(Mid(tree.Nodes(i).Key, 2)) = QueEmpresa Then
            Set tree.SelectedItem = tree.Nodes(i)
            tree.SelectedItem.Bold = True
        Else
            tree.Nodes(i).Bold = False
        End If
    Next
End Sub


Private Sub tree_NodeClick(ByVal Node As XtremeSuiteControls.TreeViewNode)
Dim Cad  As String
   
    
    If Val(Mid(Node.Key, 2)) = vEmpresa.codempre Then Exit Sub
    
    'cad = "Desea cambiar a la empresa: " & Node.Text & "?"
    'If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
    '    'Volver a poner el nodo seleccionado el que esta
        Screen.MousePointer = vbHourglass
        frmppal.CambiarEmpresa CInt(Mid(Node.Key, 2))
   '     espera 0.5
   '     Screen.MousePointer = vbDefault
   '
   ' End If
    SeleccionarNodoEmpresa vEmpresa.codempre
    
End Sub



