VERSION 5.00
Begin VB.Form frmZoom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Zoom"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10080
   Icon            =   "frmZoom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "&Actualizar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7530
      TabIndex        =   1
      Top             =   5505
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8790
      TabIndex        =   2
      Top             =   5505
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4845
      Index           =   0
      Left            =   210
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   555
      Width           =   9615
   End
   Begin VB.Label Label1 
      Caption         =   "Observaciones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   195
      Width           =   7455
   End
End
Attribute VB_Name = "frmZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event Actualizar(vCampo As String)


'variables que se pasan con valor al llamar al formulario de zoom desde otro formulario
Public pValor As String
Public pTitulo As String
Public pModo As Byte 'muestra el campo bloqueado y no se puede modificar segun Modo




Private Sub cmdActualizar_Click()
    'devolver el valor del campo al formulario que lo llamo
    RaiseEvent Actualizar(Text1(0).Text)
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    'salimos y no hacemos nada
    Unload Me
End Sub

Private Sub Form_Load()

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    'obtener el campo correspondiente y mostrarlo en el text
    Text1(0).Text = pValor
    Label1.Caption = pTitulo
    BloquearTxt Text1(0), (pModo <> 3 And pModo <> 4)
    Me.cmdActualizar.visible = (pModo = 3 Or pModo = 4)
    
 '   SendKeys "^{END}"
    CreateObject("WScript.Shell").SendKeys "^{END}"
    
End Sub
