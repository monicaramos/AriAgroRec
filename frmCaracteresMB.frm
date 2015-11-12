VERSION 5.00
Begin VB.Form frmCaracteresMB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revisión Caracteres Multibase"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6090
   Icon            =   "frmCaracteresMB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameMultibase 
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.CheckBox chkMultibase 
         Caption         =   "Telefónos de Socios"
         Height          =   255
         Index           =   4
         Left            =   2700
         TabIndex        =   13
         Top             =   3150
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkMultibase 
         Caption         =   "Pueblos"
         Height          =   255
         Index           =   3
         Left            =   2700
         TabIndex        =   12
         Top             =   2610
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CommandButton cmdMultiBase 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3480
         TabIndex        =   5
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton cmdMultiBase 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   4680
         TabIndex        =   4
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CheckBox chkMultibase 
         Caption         =   "Socios"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   2640
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkMultibase 
         Caption         =   "Cooperativas"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   2
         Top             =   3180
         Value           =   1  'Checked
         Width           =   2145
      End
      Begin VB.CheckBox chkMultibase 
         Caption         =   "Partidas"
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   1
         Top             =   3720
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.Label Label29 
         Caption         =   "Revisión caracteres multibase"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   720
         TabIndex        =   11
         Top             =   120
         Width           =   4935
      End
      Begin VB.Label Label30 
         Caption         =   "Utlidad para revisar los caracteres especiales que puedan quedar al realizar integraciones. "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   720
         Width           =   5775
      End
      Begin VB.Label Label31 
         Caption         =   "No debe trabajar nadie en la aplicación"
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
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   4815
      End
      Begin VB.Label Label32 
         Caption         =   "A este proceso le puede costar mucho tiempo."
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
         Left            =   360
         TabIndex        =   8
         Top             =   1680
         Width           =   4815
      End
      Begin VB.Label Label33 
         Caption         =   "Datos a revisar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   4815
      End
      Begin VB.Label Label34 
         Caption         =   "Label34"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   4800
         Width           =   5535
      End
      Begin VB.Line Line5 
         X1              =   240
         X2              =   5640
         Y1              =   4140
         Y2              =   4140
      End
   End
End
Attribute VB_Name = "frmCaracteresMB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQL As String

Private Sub cmdMultiBase_Click(Index As Integer)
Dim i As Integer
    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
    
    
    'Comprobamos k ha selecionado algun nivel
    NE = 0
    For i = 0 To Me.chkMultibase.Count - 1
        If Me.chkMultibase(i).Value = 1 Then NE = NE + 1
    Next i
    If NE = 0 Then
        MsgBox "Seleccione donde se van a realizar los cambios", vbExclamation
        Exit Sub
    End If
    
    'Comprobacion si hay alguien trabajando
    If UsuariosConectados Then Exit Sub
    
    SQL = "Seguro que desea continuar con el proceso"
    If MsgBox(SQL, vbCritical + vbYesNoCancel) <> vbYes Then Exit Sub
    
'   'BLOQUEAMOS LA BD
'   If Not Bloquear_DesbloquearBD(True) Then
'        MsgBox "No se ha podido bloquea a nivel de BD.", vbExclamation
'        Exit Sub
'    End If
'
    
    Screen.MousePointer = vbHourglass
    NumRegElim = 0
    For i = 0 To Me.chkMultibase.Count - 1
        If Me.chkMultibase(i).Value = 1 Then
            'Hacemos los cambios para ese valor
            HacerCambios i
        End If
    Next i
'    Bloquear_DesbloquearBD False
    Screen.MousePointer = vbDefault
    Label34.Caption = ""
    SQL = "Proceso finalizado" & vbCrLf & "Se han realizado: " & NumRegElim & " cambio(s)."
    MsgBox SQL, vbInformation
End Sub
Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
    Else
        Screen.MousePointer = vbDefault
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim W, H
    PrimeraVez = True
    Me.frameMultibase.visible = False
    'MULTIBASE
    Me.Caption = "Sustitución caracteres multibase"
    W = Me.frameMultibase.Width
    H = Me.frameMultibase.Height + 300
    Me.frameMultibase.visible = True
    Label34.Caption = ""
    cmdMultiBase(1).Cancel = True
    Me.Width = W + 120
    Me.Height = H + 120
End Sub

Private Sub HacerCambios(ByVal Tabla As Integer)
Dim Cambio As String
Dim Inicio As Integer
Dim Fin As Integer
Dim Cad As String

    'RevisaCaracterMultibase
    Select Case Tabla
    Case 0
        'Socios
        SQL = "Select codsocio,nomsocio,dirsocio,pobsocio,prosocio,observaciones"
        SQL = SQL & " FROM rsocios"
        Inicio = 1 'k es dos
        Fin = 5
    Case 1
        'Cooperativas
        SQL = "Select codcoope,nomcoope,domcoope,pobcoope,procoope  from rcoope"
        Inicio = 1
        Fin = 4
    Case 2
        'Partidas
        SQL = "Select codparti,nomparti FROM rpartida "
        Cad = ""
        Inicio = 1
        Fin = 1
    Case 3
        'Pueblos
        SQL = "Select codpobla,despobla FROM rpueblos "
        Cad = ""
        Inicio = 1
        Fin = 1
    Case 4
        'Telefonos de socios
        SQL = "Select idtelefono,codsocio,nombre,direccion,poblacion,provincia,observaciones FROM rsocios_telefonos "
        Cad = ""
        Inicio = 2
        Fin = 6
    End Select
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        
        While Not RS.EOF
            Label34.Caption = RS.Fields(0) & " - " & RS.Fields(1)
            Label34.Refresh
            Cambio = ""
            For i = Inicio To Fin
                'Campo no nulo
                If Not IsNull(RS.Fields(i)) Then
                    SQL = RS.Fields(i)
                    Cad = RevisaCaracterMultibase(SQL)
                    If SQL <> Cad Then
                        'Han habido cambios
                        If Cambio <> "" Then Cambio = Cambio & ","
'                        Sql = NombreSQL(Cad)
                        SQL = DevNombreSQL(Cad)
                        NumRegElim = NumRegElim + 1
                        Cambio = Cambio & RS.Fields(i).Name & " = '" & SQL & "'"
                    End If
                End If
            Next i
            If Cambio <> "" Then
                'OK HAY K CAMBIAR, k updatear
                Select Case Tabla
                Case 0 'rsocios
                    SQL = "UPDATE rsocios SET " & Cambio & " WHERE codsocio =" & RS.Fields(0)
            
                Case 1 'rcoope
                    SQL = "UPDATE rcoope"
                    SQL = SQL & " SET " & Cambio & " WHERE codcoope = " & DBSet(RS.Fields(0).Value, "N")
                
                Case 2 'rpartida
                    SQL = "UPDATE rpartida SET " & Cambio & " WHERE codparti =" & DBSet(RS.Fields(0).Value, "N")
                
                Case 3 'rpueblos
                    SQL = "UPDATE rpueblos SET " & Cambio & " WHERE codpobla =" & DBSet(RS.Fields(0).Value, "N")
                
                Case 4 'rsocios_telefono
                    SQL = "UPDATE rsocios_telefonos SET " & Cambio & " WHERE idtelefono =" & DBSet(RS.Fields(0).Value, "N")
                    SQL = SQL & " and codsocio = " & DBSet(RS.Fields(1).Value, "N")
                End Select
                
                'Ejecutamos
                conn.Execute SQL
            End If
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
            
End Sub

