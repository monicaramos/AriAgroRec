VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMantenusu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de usuarios"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9315
   Icon            =   "frmMantenusu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameUsuario 
      Height          =   5415
      Left            =   1845
      TabIndex        =   17
      Top             =   90
      Width           =   5655
      Begin VB.TextBox Text2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   7
         Left            =   3840
         MaxLength       =   17
         PasswordChar    =   "*"
         TabIndex        =   26
         Text            =   "123456789012345"
         Top             =   4320
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   120
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   4320
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   120
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   3600
         Width           =   5295
      End
      Begin VB.TextBox Text2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   120
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   2880
         Width           =   5295
      End
      Begin VB.TextBox Text2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   3600
         PasswordChar    =   "*"
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton cmdFrameUsu 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   28
         Top             =   4920
         Width           =   1215
      End
      Begin VB.CommandButton cmdFrameUsu 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   2880
         TabIndex        =   27
         Top             =   4920
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   3600
         PasswordChar    =   "*"
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   1200
         Width           =   4335
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmMantenusu.frx":27A2
         Left            =   120
         List            =   "frmMantenusu.frx":27AF
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "mail-password"
         Height          =   255
         Index           =   7
         Left            =   3840
         TabIndex        =   44
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "mail-user"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   43
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Servidor SMTP"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   42
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "e-mail"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   41
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "NUEVO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2040
         TabIndex        =   34
         Top             =   225
         Width           =   3375
      End
      Begin VB.Shape Shape1 
         Height          =   975
         Left            =   2280
         Top             =   1680
         Width           =   3135
      End
      Begin VB.Label Label4 
         Caption         =   "Confirma Pass."
         Height          =   255
         Index           =   3
         Left            =   2400
         TabIndex        =   33
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Password"
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   32
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Nivel"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre completo"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Login"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame FrameEditorMenus 
      Height          =   5895
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   9255
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   5055
         Left            =   240
         TabIndex        =   38
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   8916
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
      Begin VB.CommandButton cmdEditorMenus 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   8160
         TabIndex        =   37
         Top             =   5400
         Width           =   975
      End
      Begin VB.CommandButton cmdEditorMenus 
         Caption         =   "Guardar"
         Height          =   375
         Index           =   0
         Left            =   7080
         TabIndex        =   36
         Top             =   5400
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   40
         Top             =   5400
         Width           =   5055
      End
   End
   Begin VB.Frame FrameNormal 
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmdUsu 
         Height          =   400
         Index           =   3
         Left            =   1800
         Picture         =   "frmMantenusu.frx":27D4
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Prohibir acceso"
         Top             =   5400
         Width           =   400
      End
      Begin VB.CommandButton cmdConfigMenu 
         Caption         =   "Configurar menu"
         Height          =   400
         Left            =   4560
         TabIndex        =   39
         Top             =   5400
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Height          =   1815
         Left            =   3480
         TabIndex        =   8
         Top             =   360
         Width           =   5655
         Begin VB.TextBox Text4 
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   480
            Width           =   4335
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmMantenusu.frx":291E
            Left            =   120
            List            =   "frmMantenusu.frx":292E
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1200
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Nombre completo"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label3 
            Caption         =   "Nivel"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   960
            Width           =   2295
         End
      End
      Begin VB.CommandButton cmdUsu 
         Height          =   400
         Index           =   0
         Left            =   120
         Picture         =   "frmMantenusu.frx":2961
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Nuevo usuario"
         Top             =   5400
         Width           =   400
      End
      Begin VB.CommandButton cmdEmp 
         Height          =   400
         Index           =   0
         Left            =   3480
         Picture         =   "frmMantenusu.frx":2A63
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Nueva bloqueo empresa"
         Top             =   5400
         Width           =   400
      End
      Begin VB.CommandButton cmdUsu 
         Height          =   400
         Index           =   1
         Left            =   600
         Picture         =   "frmMantenusu.frx":2B65
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Modificar usuario"
         Top             =   5400
         Width           =   400
      End
      Begin VB.CommandButton cmdUsu 
         Height          =   400
         Index           =   2
         Left            =   1080
         Picture         =   "frmMantenusu.frx":2C67
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Eliminar usuario"
         Top             =   5400
         Width           =   400
      End
      Begin VB.CommandButton cmdEmp 
         Height          =   400
         Index           =   1
         Left            =   3960
         Picture         =   "frmMantenusu.frx":2D69
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Eliminar bloqueo empresa"
         Top             =   5400
         Width           =   400
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2775
         Left            =   3480
         TabIndex        =   7
         Tag             =   $"frmMantenusu.frx":2E6B
         Top             =   2520
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cod."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   5821
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Resum."
            Object.Width           =   2469
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4935
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   8705
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Login"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Usuarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   15
         Top             =   120
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Empresas NO permitidas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   14
         Top             =   2280
         Width           =   3255
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8040
      TabIndex        =   0
      Top             =   5400
      Width           =   1095
   End
End
Attribute VB_Name = "frmMantenusu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PrimeraVez As Boolean
Dim Sql As String
Dim i As Integer

Dim miRsAux As ADODB.Recordset


Private Sub cmdConfigMenu_Click()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    CargarListEditorMenu
    Label7.Caption = ListView1.SelectedItem.SubItems(1)
    Me.FrameEditorMenus.visible = True
    Me.FrameNormal.visible = False
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdEditorMenus_Click(Index As Integer)
    If Index = 0 Then
        
        GuardarMenuUsuario
    
    End If
    Me.FrameEditorMenus.visible = False
    Me.FrameNormal.visible = True
    
End Sub

Private Sub cmdEmp_Click(Index As Integer)
Dim cont As Integer

    If ListView1.SelectedItem Is Nothing Then
        MsgBox "Seleccione un usuario", vbExclamation
        Exit Sub
    End If
    
    If Index = 0 Then


        'nueva Empresa bloqueada para el usuario
        CadenaDesdeOtroForm = ""
        frmMensajes.OpcionMensaje = 14
        frmMensajes.Show vbModal
        If CadenaDesdeOtroForm <> "" Then
            cont = RecuperaValor(CadenaDesdeOtroForm, 1)
            If cont = 0 Then Exit Sub
            For i = 1 To cont
                'No hacemos nada
            Next i
            For i = 0 To cont - 1
                Sql = RecuperaValor(CadenaDesdeOtroForm, i + cont + 2)
                InsertarEmpresa CInt(Sql)
            Next i
        
        Else
            Exit Sub
        End If
        
    Else
        If ListView2.SelectedItem Is Nothing Then Exit Sub
        Sql = "Va a  desbloquear el acceso" & vbCrLf
        Sql = Sql & vbCrLf & "a la empresa:   " & ListView2.SelectedItem.SubItems(1) & vbCrLf
        Sql = Sql & "para el usuario:   " & ListView1.SelectedItem.SubItems(1) & vbCrLf & vbCrLf & "     �Desea continuar?"
        If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then
            Sql = "Delete FROM usuarios.usuarioempresasariagro WHERE codusu =" & ListView1.SelectedItem.Text
            Sql = Sql & " AND codempre = " & ListView2.SelectedItem.Text
            conn.Execute Sql
        Else
            Exit Sub
        End If
    End If
    'Llegados aqui recargamos los datos del usuario
    Screen.MousePointer = vbHourglass
    DatosUsusario
    Screen.MousePointer = vbDefault
End Sub


Private Sub InsertarEmpresa(Empresa As Integer)
    Sql = "INSERT INTO usuarios.usuarioempresasariagro(codusu,codempre) VALUES ("
    Sql = Sql & ListView1.SelectedItem.Text & "," & Empresa & ")"
    On Error Resume Next
    conn.Execute Sql
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    Else
    
    End If
    
End Sub


Private Sub cmdFrameUsu_Click(Index As Integer)



    If Index = 0 Then
        For i = 0 To Text2.Count - 1
            Text2(i).Text = Trim(Text2(i).Text)
            If i < 4 Then
                If Text2(i).Text = "" Then
                    MsgBox Label4(i).Caption & " requerido.", vbExclamation
                    Exit Sub
                End If
            End If
        Next i
        
        If Combo2.ListIndex < 0 Then
            MsgBox "Seleccione un nivel de acceso", vbExclamation
            Exit Sub
        End If
    
        'Password
        If Text2(2).Text <> Text2(3).Text Then
            MsgBox "Password y confirmacion de password no coinciden", vbExclamation
            Exit Sub
        End If
        
        
        'Ahora vamos con los campos de e-mail
        CadenaDesdeOtroForm = ""
        For i = 4 To 7
            If Text2(i).Text <> "" Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "1"
        Next i
        
        If Len(CadenaDesdeOtroForm) > 0 And Len(CadenaDesdeOtroForm) <> 4 Then
            MsgBox "Falta por rellenar correctamente los datos del e-mail.", vbExclamation
            CadenaDesdeOtroForm = ""
            Exit Sub
        End If
        
        
        
        
        
        
        'Compruebo que el login es unico
        i = 0
        If UCase(Label6.Caption) = "NUEVO" Then
            Set miRsAux = New ADODB.Recordset
            Sql = "Select login from usuarios.usuarios where login='" & Text2(0).Text & "'"
            miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Sql = ""
            If Not miRsAux.EOF Then Sql = "Ya existe en la tabla usuarios uno con el login: " & miRsAux.Fields(0)
            miRsAux.Close
            Set miRsAux = Nothing
            If Sql <> "" Then
                MsgBox Sql, vbExclamation
                Exit Sub
            End If
        Else
            'MODIFICAR
            If FrameUsuario.Tag = 0 Then
                'Estoy modificando un dato normal
                i = CInt(ListView1.SelectedItem.Text)
            Else
                'Estoy agregando un usuario que ya existia en contabi�lidad
                'es decir, le estoy asignando su NIVELUSU de contabilidad
                i = CInt(FrameUsuario.Tag)
            End If
        End If
        
        InsertarModificar i
        
        
    End If
    'Cargar usuarios
    If UCase(Label6.Caption) = "NUEVO" Then
        'CargaUsuarios
        CadenaDesdeOtroForm = ""
    Else
        'Pero cargamos el tag como coresponde
        'ListView1.SelectedItem.Tag = Combo2.ItemData(Combo2.ListIndex) & "|" & Text2(1).Text & "|"
        
        If Me.FrameUsuario.Tag <> 0 Then
            CadenaDesdeOtroForm = FrameUsuario.Tag
        Else
            CadenaDesdeOtroForm = ListView1.SelectedItem.Text
        End If
        
  
    End If
    
    CargaUsuarios
    If CadenaDesdeOtroForm <> "" Then
        For i = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(i).Text = CadenaDesdeOtroForm Then
                    Set ListView1.SelectedItem = ListView1.ListItems(i)
                    Exit For
                End If
        Next i
    End If
    DatosUsusario
    CadenaDesdeOtroForm = ""
    'Para ambos casos
    Me.FrameUsuario.visible = False
    Me.FrameNormal.Enabled = True
    
End Sub


Private Sub InsertarModificar(ByVal CodigoUsuario As Integer)
Dim Ant As Integer
Dim Fin As Boolean

On Error GoTo EInsertarModificar

    Set miRsAux = New ADODB.Recordset
    If UCase(Label6.Caption) = "NUEVO" Then
        
        'Nuevo
        Sql = "Select codusu from usuarios.usuarios where codusu > 0"
        miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Fin = False
        Ant = 1
        While Not Fin
            If miRsAux!CodUsu - Ant > 0 Then
                'Hay un salto
                i = Ant
                Fin = True
            Else
                Ant = Ant + 1
            End If
            If Not Fin Then
                miRsAux.MoveNext
                If miRsAux.EOF Then
                    Fin = True
                    i = Ant
                End If
            End If
        Wend
        miRsAux.Close

        
        Sql = "INSERT INTO usuarios.usuarios (codusu, nomusu,  nivelariagro, login, passwordpropio,dirfich) VALUES ("
        Sql = Sql & i
        Sql = Sql & ",'" & Text2(1).Text & "',"
        'Combo
        Sql = Sql & Combo2.ItemData(Combo2.ListIndex) & ",'"
        Sql = Sql & Text2(0).Text & "','"
        Sql = Sql & Text2(3).Text & "',"
        'DIR FICH tiene
        If Text2(4).Text = "" Then
            CadenaDesdeOtroForm = "NULL"
        Else
            CadenaDesdeOtroForm = ""
            For i = 4 To 7
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & Text2(i).Text & "|"
            Next i
            CadenaDesdeOtroForm = "'" & CadenaDesdeOtroForm & "'"
        End If
        Sql = Sql & CadenaDesdeOtroForm & ")"
        
    Else
        Sql = "UPDATE usuarios.usuarios Set nomusu=" & DBSet(Text2(1).Text, "T")
        
        'Si el combo es administrador compruebo que no fuera en un principio SUPERUSUARIO
        If Combo2.ListIndex = 2 Then
            'Si el combo1 es 3 entonces es super
            If Combo1.ListIndex = 3 Then
                i = 0
            Else
                i = 1
            End If
        Else
            i = Combo2.ItemData(Combo2.ListIndex)
        End If
        Sql = Sql & " , nivelariagro =" & i
        'SQL = SQL & "  , login = '" & Text2(2).Text
        Sql = Sql & "  , passwordpropio = '" & Text2(3).Text & "'"
        
        
        'El e-mail
        If Text2(4).Text = "" Then
            CadenaDesdeOtroForm = "NULL"
        Else
            CadenaDesdeOtroForm = ""
            For i = 4 To 7
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & Text2(i).Text & "|"
            Next i
            CadenaDesdeOtroForm = "'" & CadenaDesdeOtroForm & "'"
        End If
        Sql = Sql & " ,dirfich = " & CadenaDesdeOtroForm
        
        
        
        
        'aqui, en lugar del selecteditem tengo k pasarle el codigo de usuario
        'ya que cuando es nuevo usario y cojo los datos desde otra aplicacion entonces
        'no lo tengo selected y enonces peta
        
        Sql = Sql & " WHERE codusu = " & CodigoUsuario
    End If
    conn.Execute Sql
    CadenaDesdeOtroForm = ""
    Exit Sub
EInsertarModificar:
    MuestraError Err.Number, "EInsertarModificar"
End Sub



Private Sub cmdUsu_Click(Index As Integer)
    
    
    Select Case Index
    Case 0, 1
        limpiar Me
        If Index = 0 Then
            'Nuevo usuario
            
            Label6.Caption = "NUEVO"
            i = 0 'Para el foco
        Else
            'Modificar
            If ListView1.SelectedItem Is Nothing Then Exit Sub
            Label6.Caption = "MODIFICAR"
            Set miRsAux = New ADODB.Recordset
            Sql = "Select * from usuarios.usuarios where codusu = " & ListView1.SelectedItem.Text
            miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If miRsAux.EOF Then
                MsgBox "Error inesperado: Leer datos usuarios", vbExclamation
            Else
                'LimpiarCamposUsuario
                PonerDatosUsuario
            End If
            i = 1 'Para el foco
            FrameUsuario.Tag = 0  'Marcamos que es una modificacion desde un usuario existente
        End If
        Text2(0).Enabled = (Index = 0)
        Me.FrameNormal.Enabled = False
        Me.FrameUsuario.visible = True
        Text2(i).SetFocus
    Case 2, 3
        If ListView1.SelectedItem Is Nothing Then Exit Sub
        i = vUsu.Codigo Mod 1000
        If i = CInt(ListView1.SelectedItem.Text) Then
            MsgBox "El usuario es el mismo con el que esta trabajando actualmente", vbInformation
            Exit Sub
        End If
        
        If Index = 2 Then
            
            Sql = "El usuario " & ListView1.SelectedItem.SubItems(1) & " ser� eliminado y no tendra acceso a los programas de Ariadna (AriConta, AriGes....) ." & vbCrLf
            Sql = Sql & vbCrLf & "                              �Desea continuar?"
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            Sql = "DELETE from usuarios.usuarios where codusu = " & ListView1.SelectedItem.Text
            
        Else
            Sql = "Al usuario " & ListView1.SelectedItem.SubItems(1) & " no le estar� permitido el acceso al programa AriAgro." & vbCrLf
            Sql = Sql & vbCrLf & "                              �Desea continuar?"
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            Sql = "UPDATE usuarios.usuarios SET nivelariagro = -1 WHERE codusu = " & ListView1.SelectedItem.Text
        End If
        Screen.MousePointer = vbHourglass
        conn.Execute Sql
        
            '//El codigo siguiente seria mas logico meterlo en el modulo de usuario
            '   pero de momento un saco de cemento
            If Index = 2 Then EliminarAuxiliaresUsuario CInt(ListView1.SelectedItem.Text)
        
            CargaUsuarios
        Screen.MousePointer = vbDefault
    
    End Select

End Sub




Private Sub EliminarAuxiliaresUsuario(CodUsu As Integer)

    On Error GoTo EEliminarAuxiliaresUsuario
    Sql = "DELETE FROM usuarios.usuarioempresasariagro where codusu =" & CodUsu
    conn.Execute Sql
    
    Sql = "DELETE FROM usuarios.appmenususuario where  codusu =" & CodUsu
    conn.Execute Sql
    
    Exit Sub
EEliminarAuxiliaresUsuario:
    MuestraError Err.Number, "Eliminar Auxiliares Usuario"

End Sub

Private Sub PonerDatosUsuario()
            Text2(0).Text = miRsAux!Login
            Text2(1).Text = miRsAux!nomusu
            Text2(2).Text = miRsAux!passwordpropio
            Text2(3).Text = miRsAux!passwordpropio
            i = miRsAux!nivelariagro
            If i = -1 Then i = 3
            If i < 2 Then
                Combo2.ListIndex = 2
            Else
                If i = 2 Then
                    Combo2.ListIndex = 1
                Else
                    Combo2.ListIndex = 0
                End If
            End If
       
        
        'Cargamos los datos del correo e-mail
        Sql = Trim(DBLet(miRsAux!Dirfich, "T"))
        If Sql <> "" Then
            For i = 1 To 4
                Text2(3 + i).Text = RecuperaValor(Sql, i)
            Next i
        End If

End Sub


Private Sub Combo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
'        SendKeys "{tab}"
        CreateObject("WScript.Shell").SendKeys "{tab}"

    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Me.ListView1.SmallIcons = frmPpal.ImageListB
        Me.ListView2.SmallIcons = frmPpal.ImageListB
        CargaUsuarios
    End If
    FrameEditorMenus.visible = False
    LeerEditorMenus
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()



    PrimeraVez = True
    Me.Icon = frmPpal.Icon
    
    Me.FrameUsuario.visible = False
    Me.FrameNormal.Enabled = True
    Me.cmdEmp(1).Picture = frmPpal.ImageListB.ListImages(7).Picture
    Me.cmdEmp(0).Picture = frmPpal.ImageListB.ListImages(6).Picture
    
    Me.cmdUsu(0).Picture = frmPpal.ImageListB.ListImages(1).Picture
    Me.cmdUsu(1).Picture = frmPpal.ImageListB.ListImages(2).Picture
    Me.cmdUsu(2).Picture = frmPpal.ImageListB.ListImages(3).Picture
    Me.cmdUsu(3).Picture = frmPpal.ImageListB.ListImages(4).Picture
End Sub



Private Sub CargaUsuarios()
Dim Itm As ListItem

    ListView1.ListItems.Clear
    Set miRsAux = New ADODB.Recordset
    '                               Aquellos usuarios k tengan nivel usu -1 NO son de conta
    '  QUitamos codusu=0 pq es el usuario ROOT
    Sql = "Select * from usuarios.usuarios where nivelariagro >=0 and codusu > 0 order by codusu"
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set Itm = ListView1.ListItems.Add
        Itm.Text = miRsAux!CodUsu
        Itm.SubItems(1) = miRsAux!Login
        Itm.SmallIcon = 8
        'Nombre y nivel de usuario
        Sql = miRsAux!nivelariagro & "|" & miRsAux!nomusu & "|"
        Itm.Tag = Sql
        'Sig
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    If ListView1.ListItems.Count > 0 Then
        Set ListView1.SelectedItem = ListView1.ListItems(1)
        DatosUsusario
    End If

End Sub



Private Sub DatosUsusario()
Dim ItmX As ListItem
On Error GoTo EDatosUsu

    If ListView1.SelectedItem Is Nothing Then
        Text4.Text = ""
        Combo1.ListIndex = -1
        Exit Sub
    End If


    Text4.Text = RecuperaValor(ListView1.SelectedItem.Tag, 2)
    'NIVEL
    Sql = RecuperaValor(ListView1.SelectedItem.Tag, 1)
    '                           COMBO                      en Bd
    '                       0.- Consulta                     3
    '                       1.- Normal                       2
    '                       2.- Administrador                1
    '                       3.- SuperUsuario (root)          0
    If Not IsNumeric(Sql) Then Sql = 3
    Select Case Val(Sql)
    Case 2
        Combo1.ListIndex = 1
    Case 1
        Combo1.ListIndex = 2
    Case 0
        Combo1.ListIndex = 3
    Case Else
        Combo1.ListIndex = 0
    End Select
    
    ListView2.ListItems.Clear
    Sql = ListView2.Tag & ListView1.SelectedItem.Text
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
        Set ItmX = ListView2.ListItems.Add
        ItmX.Text = miRsAux.Fields(0)
        ItmX.SubItems(1) = miRsAux!nomempre
        ItmX.SubItems(2) = miRsAux!nomresum
        ItmX.SmallIcon = 5
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    Exit Sub
EDatosUsu:
    MuestraError Err.Number, Err.Description
End Sub

Private Sub ListView1_ItemClick(ByVal item As MSComctlLib.ListItem)
    Screen.MousePointer = vbHourglass
    DatosUsusario
    Screen.MousePointer = vbDefault
End Sub



Private Sub Text2_GotFocus(Index As Integer)
    With Text2(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
'        SendKeys "{tab}"
        CreateObject("WScript.Shell").SendKeys "{tab}"

    End If
End Sub

Private Sub Text2_LostFocus(Index As Integer)
Dim AsignarDatos As Boolean

    Text2(Index).Text = Trim(Text2(Index).Text)
    If Text2(Index).Text = "" Then Exit Sub
    
    If Index = 0 Then
        If UCase(Label6.Caption) = "NUEVO" Then
        
            'Si es nuevo entonces, primero compruebo que no existe el login
            'Si existe, y el usuario tiene nivel conta >=0 entonces
            ' existe en la conta. Si existe pero el nivel conta es -1 entonces
            'lo que hacemos es ponerle los datos y que cambie la opcion de nivel usu
            Sql = "Select * from usuarios.usuarios where login='" & Text2(0).Text & "'"
            Set miRsAux = New ADODB.Recordset
            miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not miRsAux.EOF Then
                'Tiene nivel usu
                If miRsAux!nivelariagro > 0 Then
                    MsgBox "El usuario ya existe.", vbExclamation
                    LimpiarCamposUsuario
                    Text2(0).SetFocus
                    
                Else
                    If miRsAux!CodUsu = 0 Then
                        MsgBox "Esta intentando modificar datos del usuario ADMINISTRADOR", vbCritical
                        AsignarDatos = False
                    Else
                        Sql = "El usuario existe para otras aplicaciones de Ariadna Software." & vbCrLf
                        Sql = Sql & "�Desea agregarlo como usuario al AriAgro?"
                        If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then AsignarDatos = True
                    End If
                    If AsignarDatos Then
                        PonerDatosUsuario
                        'Ahora pongo el label y el campo a disbled
                        Text2(1).SetFocus
                        Label6.Caption = "MODIFICAR"
                        Text2(0).Enabled = False
                        FrameUsuario.Tag = miRsAux!CodUsu 'Pongo el frame al codigo ndel usuario
                    Else
                        LimpiarCamposUsuario
                        Text2(0).SetFocus
                    End If
                End If
            End If
            miRsAux.Close
        End If
    End If
    
End Sub

Private Sub LimpiarCamposUsuario()
    For i = 0 To 7
        Text2(i).Text = ""
    Next i
End Sub

Private Sub LeerEditorMenus()
    On Error GoTo ELeerEditorMenus
    cmdConfigMenu.visible = False
    Sql = "Select count(*) from usuarios.appmenus where aplicacion='AriagroRec'"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
            If miRsAux.Fields(0) > 0 Then cmdConfigMenu.visible = True
        End If
    End If
    miRsAux.Close
        

    
ELeerEditorMenus:
    Set miRsAux = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargarListEditorMenu()
Dim Nod As Node
Dim J As Integer

    TreeView1.Nodes.Clear
    Sql = "Select * from usuarios.appmenus where aplicacion='AriagroRec'"
    Sql = Sql & " ORDER BY padre ,orden"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        If DBLet(miRsAux!padre, "N") = 0 Then
            Set Nod = TreeView1.Nodes.Add(, , "C" & miRsAux!Contador)
        Else
            Sql = "C" & miRsAux!padre
            Set Nod = TreeView1.Nodes.Add(Sql, tvwChild, "C" & miRsAux!Contador)
        End If
        Sql = miRsAux!Name & "|"
        If Not IsNull(miRsAux!indice) Then Sql = Sql & miRsAux!indice
        Nod.Tag = Sql
   
        Nod.Text = miRsAux!Caption
        Nod.Checked = True
        Nod.EnsureVisible
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    If TreeView1.Nodes.Count > 1 Then TreeView1.Nodes(1).EnsureVisible
    
    'AHora ire nodo a nodo buscando los k deshabilitamos de la aplicacion
    Sql = "Select * from usuarios.appmenususuario where aplicacion='AriagroRec' AND codusu =" & ListView1.SelectedItem.Text
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        For i = 1 To TreeView1.Nodes.Count
            Sql = miRsAux!Tag
            If TreeView1.Nodes(i).Tag = Sql Then
                TreeView1.Nodes(i).Checked = False
                If TreeView1.Nodes(i).Children > 0 Then Recursivo2 TreeView1.Nodes(i).Child, TreeView1.Nodes(i).Checked
                Exit For
            End If
        Next i
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    
    
    
    
    
    
    Set miRsAux = Nothing
End Sub



Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
If Node.Children > 0 Then Recursivo2 Node.Child, Node.Checked
End Sub


Private Sub CheckarNodo(N As Node, Valor As Boolean)
Dim NO As Node
    Set NO = N.LastSibling
    Do
        N.Checked = Valor
        If N.Children > 0 Then CheckarNodo N, Valor
        If N.Next <> NO.LastSibling Then Set N = N.Next
    Loop Until NO = N
End Sub

Private Sub Recursivo2(ByVal Nod As Node, Valor As Boolean)
Dim nx As Node
Dim Aux

    
    
    Set nx = Nod.FirstSibling
    While nx <> Nod.LastSibling
        If nx.Children > 0 Then Recursivo2 nx.Child, Valor
        nx.Checked = Valor
        'aux = nx.Root
        'aux = nx.Parent
        Set nx = nx.Next
    Wend
    
    If nx = Nod.LastSibling Then
        If nx.Children > 0 Then Recursivo2 nx.Child, Valor
        nx.Checked = Valor
      End If
    Set nx = Nothing
End Sub


Private Sub GuardarMenuUsuario()
    Sql = "DELETE from usuarios.appmenusUsuario where aplicacion='Ariagrorec' AND codusu =" & ListView1.SelectedItem.Text
    conn.Execute Sql
    
    i = 0
    Sql = "INSERT INTO usuarios.appmenususuario (aplicacion, codusu, codigo, tag) VALUES ('Ariagrorec'," & ListView1.SelectedItem.Text & ","
    RecursivoBD TreeView1.Nodes(1)
End Sub

Private Sub InsertaBD(vtag As String)
Dim c As String
    i = i + 1
    'SQL = "INSERT INTO appmenususuario (aplicacion, codusu, codigo, tag)
    c = Sql & i & ",'" & vtag & "')"
    conn.Execute c
End Sub


Private Sub RecursivoBD(ByVal Nod As Node)
Dim nx As Node
Dim Aux

    
    
    Set nx = Nod.FirstSibling
    While nx <> Nod.LastSibling
        If nx.Children > 0 Then
            If nx.Checked Then RecursivoBD nx.Child
        End If
        If Not nx.Checked Then InsertaBD nx.Tag
        'aux = nx.Root
        'aux = nx.Parent
        Set nx = nx.Next
    Wend
    
    If nx = Nod.LastSibling Then
        If nx.Children > 0 Then
            If nx.Checked Then RecursivoBD nx.Child
        End If
        If Not nx.Checked Then InsertaBD nx.Tag
      End If
    Set nx = Nothing
End Sub


