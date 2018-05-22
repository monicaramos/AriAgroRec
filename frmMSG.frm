VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMSG 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6660
   Icon            =   "frmMSG.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6075
      Top             =   4770
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameMSGBOX 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6585
      Begin VB.Frame FrameSiNo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   870
         Left            =   3825
         TabIndex        =   1
         Top             =   2565
         Width           =   2625
         Begin VB.CommandButton CmdNo 
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   0
            Left            =   1395
            TabIndex        =   3
            Top             =   270
            Width           =   1065
         End
         Begin VB.CommandButton CmdSi 
            Caption         =   "Si"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   0
            Left            =   180
            TabIndex        =   2
            Top             =   270
            Width           =   1065
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1590
         Left            =   1305
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "frmMSG.frx":000C
         Top             =   135
         Width           =   5100
      End
      Begin VB.Frame FrameAceptar 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   870
         Left            =   5040
         TabIndex        =   8
         Top             =   2565
         Width           =   1365
         Begin VB.CommandButton CmdAceptar 
            Caption         =   "Aceptar"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   0
            Left            =   180
            TabIndex        =   9
            Top             =   270
            Width           =   1065
         End
      End
      Begin VB.Frame FrameSiNoCancelar 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   870
         Left            =   2565
         TabIndex        =   4
         Top             =   2565
         Width           =   3795
         Begin VB.CommandButton CmdCancelar 
            Caption         =   "Cancelar"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   0
            Left            =   2610
            TabIndex        =   7
            Top             =   270
            Width           =   1065
         End
         Begin VB.CommandButton CmdSi 
            Caption         =   "Si"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   1
            Left            =   180
            TabIndex        =   6
            Top             =   270
            Width           =   1065
         End
         Begin VB.CommandButton CmdNo 
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   1
            Left            =   1395
            TabIndex        =   5
            Top             =   270
            Width           =   1065
         End
      End
      Begin MSComctlLib.Toolbar ToolbarMSG 
         Height          =   330
         Left            =   225
         TabIndex        =   11
         Top             =   135
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ayuda"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: LAURA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(Cliente As String, Observaciones As String)

Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean


Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
Dim PrimeraVez As Boolean




Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me

End Sub



Private Sub CmdSi_Click(Index As Integer)
    ValorDevuelto = vbYes
    Unload Me
End Sub

Private Sub CmdNo_Click(Index As Integer)
    ValorDevuelto = vbNo
    Unload Me
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    ValorDevuelto = vbCancel
    Unload Me
End Sub

Private Sub cmdAceptar_Click(Index As Integer)
    ValorDevuelto = vbOK
    Unload Me
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    ValorDevuelto = vbCancel

    PrimeraVez = True

    FrameSiNo.visible = False
    FrameSiNo.Enabled = False

    FrameSiNoCancelar.visible = False
    FrameSiNoCancelar.Enabled = False
    
    FrameAceptar.visible = False
    FrameAceptar.Enabled = False
    
    Me.Caption = "AriagroRec"


    Select Case NumCod
        'exclamation
        Case 48
            FrameAceptar.visible = True
            FrameAceptar.Enabled = True
            
            PonerFocoBtn CmdAceptar(0)
    
            Me.ToolbarMSG.ImageList = frmPpal.ImageList1.ListImages(2).Picture
            Me.ToolbarMSG.Buttons(1).Image = 12
            
        'information
        Case 64
            FrameAceptar.visible = True
            FrameAceptar.Enabled = True
            
            PonerFocoBtn CmdAceptar(0)
            
            Me.ToolbarMSG.ImageList = frmPpal.ImageList1.ListImages(3).Picture
            Me.ToolbarMSG.Buttons(1).Image = 12
    
        ' question
        Case 32 + 4 '(q + s/n)
            Me.ToolbarMSG.ImageList = frmPpal.ImageList1
            Me.ToolbarMSG.Buttons(1).Image = 12
        
            FrameSiNo.visible = True
            FrameSiNo.Enabled = True
        
            PonerFocoBtn CmdSi(0)
        
        Case 32 + 4 + 256 '(q + s/n + defN)
            Me.ToolbarMSG.ImageList = frmPpal.ImageList1
            Me.ToolbarMSG.Buttons(1).Image = 12
            
            FrameSiNo.visible = True
            FrameSiNo.Enabled = True
            
            PonerFocoBtn CmdNo(0)
            
        Case 32 + 3 '(q + s/n/cancel)
            Me.ToolbarMSG.ImageList = frmPpal.ImageList1
            Me.ToolbarMSG.Buttons(1).Image = 12
            
            FrameSiNoCancelar.visible = True
            FrameSiNoCancelar.Enabled = True
            
            PonerFocoBtn CmdSi(1)
            
        Case 32 + 3 + 256 '(q + s/n/cancel + defN)
            Me.ToolbarMSG.ImageList = frmPpal.ImageList1
            Me.ToolbarMSG.Buttons(1).Image = 12
            
            FrameSiNoCancelar.visible = True
            FrameSiNoCancelar.Enabled = True
    
            PonerFocoBtn CmdNo(1)
    
        Case 32 + 3 + 512 '(q + s/n/cancel + defCancel)
            Me.ToolbarMSG.ImageList = frmPpal.ImageList1
            Me.ToolbarMSG.Buttons(1).Image = 12
            
            FrameSiNoCancelar.visible = True
            FrameSiNoCancelar.Enabled = True
            
            PonerFocoBtn CmdCancelar(0)
    End Select
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
'    MsgBox "a"
    If Cancel = 1 Then ValorDevuelto = vbCancel
End Sub
