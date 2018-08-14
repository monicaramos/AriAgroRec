VERSION 5.00
Begin VB.Form frmIdentifica 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   390
      Top             =   3960
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   4920
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   4980
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   4920
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   4020
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   300
      TabIndex        =   6
      Top             =   90
      Width           =   7305
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3420
      TabIndex        =   5
      Top             =   1080
      Width           =   1725
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   2
      Left            =   5175
      TabIndex        =   4
      Top             =   4950
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   3
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   4920
      TabIndex        =   2
      Top             =   3660
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   5535
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "frmIdentifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PrimeraVez As Boolean
Dim T1 As Single
Dim vSegundos As Integer


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        
        If EsMonasterios Then
            Me.Top = 200
        End If
        
        
        
        espera 0.5
        Me.Refresh
        DoEvents
        
        'Vemos datos de ConfigAgro.ini
        Set vConfig = New Configuracion
        If vConfig.Leer = 1 Then

             MsgBox "MAL CONFIGURADO", vbCritical
             End
             Exit Sub
        End If
        
        
        Me.Timer1.Enabled = True
        
         'Abrimos conexion para comprobar el usuario
         'Luego, en funcion del nivel de usuario que tenga cerraremos la conexion
         'y la abriremos con usuario-codigo ajustado a su nivel
         If AbrirConexionUsuarios() = False Then
             MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical
             End
         End If
         
'--[Monica] 29/04/2010 : Quitamos la llamada al locker
'         'La llave
'         Load frmLLave
'         If Not frmLLave.ActiveLock1.RegisteredUser Then
'             'No ESTA REGISTRADO
'             frmLLave.Show vbModal
'         Else
'             Unload frmLLave
'         End If
'--

         'Leemos el ultimo usuario conectado
         NumeroEmpresaMemorizar True
         
         T1 = T1 + 2.5 - Timer
         If T1 > 0 Then espera T1

         
         PonerVisible True
         If Text1(0).Text <> "" Then
            Text1(1).SetFocus
        Else
            Text1(0).SetFocus
        End If
        
        
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    PonerVisible False
    T1 = Timer
    Text1(0).Text = ""
    Text1(1).Text = ""
    PrimeraVez = True
    CargaImagen
    Label2.Caption = "Ver. " & App.Major & "." & App.Minor & "." & App.Revision
    
    
    Label3.Caption = ""
    vSegundos = 60
    Label3.Caption = ""
    
    If EsMonasterios Then
         Me.Top = 200
    End If
    
    
    
End Sub


Private Sub CargaImagen()
    On Error Resume Next
    Me.Image1 = LoadPicture(App.Path & "\entrada.dat")
    Me.Height = Me.Image1.Height
    Me.Width = Me.Image1.Width
    
    If Err.Number <> 0 Then
        MsgBox Err.Description & vbCrLf & vbCrLf & "Error cargando", vbCritical
        Set conn = Nothing
        End
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    NumeroEmpresaMemorizar False
End Sub


Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress (KeyAscii)
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).Text = Trim(Text1(Index).Text)
    
    'Comprobamos si los dos estan con datos
    If Text1(0).Text <> "" And Text1(1).Text <> "" Then
        'Probar conexion usuario
        Validar
    End If
        
    
End Sub



Private Sub Validar()
Dim NuevoUsu As Usuario
Dim OK As Byte

    'Validaremos el usuario y despues el password
    Set vUsu = New Usuario
    
    If vUsu.Leer(Text1(0).Text) = 0 Then
        'Con exito
        If vUsu.PasswdPROPIO = Text1(1).Text Then
            OK = 0
        Else
            OK = 1
        End If

    Else
        OK = 2
    End If
    
    If OK <> 0 Then
        MsgBox "Usuario-Clave Incorrecto", vbExclamation

            Text1(1).Text = ""
            Text1(0).SetFocus
    Else
        'OK
        Screen.MousePointer = vbHourglass
        CadenaDesdeOtroForm = "OK"
        Label1(2).Caption = ""  'Si tarda pondremos texto aquin
        PonerVisible False
        Me.Refresh
        DoEvents
        Screen.MousePointer = vbHourglass
        HacerAccionesBD
        Unload Me
    End If

End Sub

Private Sub HacerAccionesBD()
Dim SQL As String


    
''''    'Limpiamos datos blanace
''''    SQL = "DELETE from Usuarios.ztmpbalancesumas where codusu= " & vUsu.Codigo
''''    Conn.Execute SQL
''''    Me.Refresh
''''
''''    SQL = "DELETE from Usuarios.ztmpconextcab where codusu= " & vUsu.Codigo
''''    Conn.Execute SQL
''''    Me.Refresh
''''
''''    SQL = "DELETE from usuarios.ztmpconext where codusu= " & vUsu.Codigo
''''    Conn.Execute SQL
''''    Me.Refresh
''''
''''    SQL = "DELETE from Usuarios.zcuentas where codusu= " & vUsu.Codigo
''''    Conn.Execute SQL
''''
''''    Me.Refresh
''''    SQL = "DELETE from usuarios.ztmplibrodiario where codusu= " & vUsu.Codigo
''''    Conn.Execute SQL
    
    
End Sub


Private Sub PonerVisible(visible As Boolean)
    Label1(2).visible = Not visible  'Cargando
    Text1(0).visible = visible
    Text1(1).visible = visible
    Label1(0).visible = visible
    Label1(1).visible = visible
End Sub




'Lo que haremos aqui es ver, o guardar, el ultimo numero de empresa
'a la que ha entrado, y el usuario
Private Sub NumeroEmpresaMemorizar(Leer As Boolean)
Dim NF As Integer
Dim cad As String
On Error GoTo ENumeroEmpresaMemorizar


        
    cad = App.Path & "\ultusuT.dat"
    If Leer Then
        If Dir(cad) <> "" Then
            NF = FreeFile
            Open cad For Input As #NF
            Line Input #NF, cad
            Close #NF
            cad = Trim(cad)
            
                'El primer pipe es el usuario
                Text1(0).Text = cad
    
        End If
    Else 'Escribir
        NF = FreeFile
        Open cad For Output As #NF
        cad = Text1(0).Text
        Print #NF, cad
        Close #NF
    End If
ENumeroEmpresaMemorizar:
    Err.Clear
End Sub

Private Sub Timer1_Timer()
    'Label3 = "Si no entra en " & vSegundos & " segundos. La aplicación se cerrará."
    If vSegundos < 50 Then
        Label3 = "Si no hace login, la pantalla se cerrará automáticamente en " & " " & vSegundos & " segundos"
        Me.Refresh
        DoEvents
    End If
    vSegundos = vSegundos - 1
    If vSegundos = -1 Then Unload Me
End Sub
