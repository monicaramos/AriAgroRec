VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCal 
   BorderStyle     =   0  'None
   Caption         =   "Calendario"
   ClientHeight    =   2340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2550
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   2550
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowToday       =   0   'False
      StartOfWeek     =   166723586
      TitleBackColor  =   11829830
      TitleForeColor  =   -2147483639
      CurrentDate     =   38421
   End
End
Attribute VB_Name = "frmCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: CÈSAR +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public Event Selec(vFecha As Date)
Public NovaData As String

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    RaiseEvent Selec(MonthView1.Value)
    Unload Me
End Sub

Private Sub MonthView1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'enter
        RaiseEvent Selec(MonthView1.Value)
        Unload Me
    ElseIf KeyAscii = 27 Then Unload Me 'escape
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    If (NovaData <> "") Then
        MonthView1.Value = Format(NovaData, "dd/MM/yyyy")
    Else
        MonthView1.Value = Format(Now, "dd/MM/yyyy")
    End If
End Sub
