VERSION 5.00
Begin VB.Form frmErrores2 
   Caption         =   "Se han producido errores"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8895
   Icon            =   "frmErrores.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Terminar"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Continuar"
      Height          =   375
      Left            =   500
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   3855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmErrores.frx":0442
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frmErrores2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Opcion As Byte


Private Sub cmdAceptar_Click()
    Unload Me
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
    FormularioOK = ""
    If Opcion = 1 Then
        Caption = "Recuperar MYSQL DUMP"
    Else
        Caption = "Se han producio errores."
    
    End If
    Command1.visible = Opcion <> 1
    Me.cmdAceptar.visible = Not Command1.visible
    Me.CmdCancelar.visible = Not Command1.visible
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    If Me.Width < 1000 Then Me.Width = 1000
    If Me.Height < 1000 Then Me.Height = 1000
    Me.Command1.Top = Me.Height - Command1.Height - 500
    Me.Command1.Left = Me.Width - Command1.Width - 500
    Me.CmdCancelar.Top = Command1.Top
    CmdCancelar.Left = Command1.Left
    Me.cmdAceptar.Top = Command1.Top
    
    Me.cmdAceptar.Left = CmdCancelar.Left - cmdAceptar.Width - 320
    
    Text1.Width = Me.Width - 300
    Text1.Height = Command1.Top - 300
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Opcion = 0
End Sub
