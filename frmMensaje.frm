VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMensaje 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Mensajes"
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameAcercaDE 
      BorderStyle     =   0  'None
      ForeColor       =   &H00C00000&
      Height          =   4545
      Left            =   90
      TabIndex        =   10
      Top             =   240
      Width           =   5385
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   11
         Top             =   3900
         Width           =   1035
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax: 96 3420938"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   3180
         TabIndex        =   18
         Top             =   3540
         Width           =   1440
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tfno: 902 88 88 78"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   780
         TabIndex        =   17
         Top             =   3540
         Width           =   1650
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "AriagroRec"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   915
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   4560
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   81.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1725
         Left            =   3780
         TabIndex        =   15
         Top             =   60
         Width           =   1350
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "46007 - VALENCIA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   3240
         TabIndex        =   14
         Top             =   3120
         Width           =   1620
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C/ Uruguay, 11 - Despacho 101"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   300
         TabIndex        =   13
         Top             =   3120
         Width           =   2730
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label13"
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
         Left            =   0
         TabIndex        =   12
         Top             =   1200
         Width           =   3795
      End
      Begin VB.Image Image2 
         Height          =   540
         Left            =   1740
         Top             =   2460
         Width           =   2880
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00B48246&
         BorderWidth     =   5
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   4425
         Left            =   90
         Top             =   45
         Width           =   5250
      End
      Begin VB.Image Image1 
         Height          =   4395
         Left            =   -30
         Stretch         =   -1  'True
         Top             =   30
         Width           =   5355
      End
   End
   Begin VB.Frame FrameCobrosPtes 
      Height          =   2955
      Left            =   0
      TabIndex        =   0
      Top             =   90
      Width           =   8295
      Begin VB.CommandButton CmdCancelarCobros 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   6870
         TabIndex        =   3
         Top             =   2400
         Width           =   1035
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1815
         Left            =   210
         TabIndex        =   2
         Top             =   390
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   3201
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton CmdAceptarCobros 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5760
         TabIndex        =   1
         Top             =   2400
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Label2"
         Height          =   345
         Index           =   1
         Left            =   450
         TabIndex        =   4
         Top             =   1470
         Width           =   3555
      End
   End
   Begin VB.Frame FrameErrores 
      Height          =   5505
      Left            =   0
      TabIndex        =   5
      Top             =   45
      Width           =   8835
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   6930
         TabIndex        =   6
         Top             =   4830
         Width           =   1035
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4155
         Left            =   210
         TabIndex        =   7
         Top             =   540
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   7329
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label2 
         Caption         =   "Errores de Comprobación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   270
         TabIndex        =   9
         Top             =   210
         Width           =   3585
      End
      Begin VB.Label Label1 
         Caption         =   "Label2"
         Height          =   345
         Index           =   2
         Left            =   450
         TabIndex        =   8
         Top             =   1470
         Width           =   3555
      End
   End
End
Attribute VB_Name = "frmMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public OpcionMensaje As Byte

'variables que se pasan con valor al llamar al formulario de zoom desde otro formulario

Public pTitulo As String




Private Sub cmdAceptar_Click()
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    'salimos y no hacemos nada
    Unload Me
End Sub

Private Sub CmdAceptarCobros_Click()
     Unload Me
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    PonerFocoBtn Me.cmdAceptar
End Sub


Private Sub Form_Load()
    Me.Shape1.Width = Me.Width - 30
    Me.Shape1.Height = Me.Height - 30

    'obtener el campo correspondiente y mostrarlo en el text
    
    Label1(1).Caption = pTitulo

    FrameErrores.visible = False
    frameAcercaDE.visible = False
    FrameCobrosPtes.visible = False

    Select Case OpcionMensaje
        Case Is <= 3 ' Errores al hacer comprobaciones
            PonerFrameCobrosPtesVisible True, 1000, 2000
            CargarListaErrComprobacion
            Me.Caption = "Errores de Comprobacion: "
            PonerFocoBtn Me.CmdSalir
            
        Case 10 'Errores al contabilizar facturas
            PonerFrameCobrosPtesVisible True, 1000, 2000
            CargarListaErrContab
            Me.Caption = "Facturas NO contabilizadas: "
            PonerFocoBtn Me.CmdAceptarCobros
            
        Case 6
            PonerFrameCobrosPtesVisible True, 1000, 2000
            CargaImagen
    '        Me.Caption = "Acerca de ....."
    '        w = Me.frameAcercaDE.Width
    '        h = Me.frameAcercaDE.Height
            Me.frameAcercaDE.visible = True
            Label13.Caption = "Versión:  " & App.Major & "." & App.Minor & "." & App.Revision & " "
    
        End Select
    

End Sub

Private Sub CargarListaErrContab()
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarList

    Sql = " SELECT  * "
    Sql = Sql & " FROM tmperrfac "
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        ListView1.Height = 4500
        ListView1.Width = 7400
        ListView1.Left = 500
        ListView1.Top = 500

        'Los encabezados
        ListView1.ColumnHeaders.Clear

        If RS.Fields(0).Name = "codprove" Then
            'Facturas de Compra
             ListView1.ColumnHeaders.Add , , "Prove.", 700
        Else 'Facturas de Venta
            ListView1.ColumnHeaders.Add , , "Tipo", 600
        End If
        ListView1.ColumnHeaders.Add , , "Factura", 1000, 1
        ListView1.ColumnHeaders.Add , , "Fecha", 1100, 1
        ListView1.ColumnHeaders.Add , , "Error", 4620
    
        While Not RS.EOF
            Set ItmX = ListView1.ListItems.Add
            'El primer campo será codtipom si llamamos desde Ventas
            ' y será codprove si llamamos desde Compras
            ItmX.Text = RS.Fields(0).Value
            ItmX.SubItems(1) = Format(RS!NumFactu, "0000000")
            ItmX.SubItems(2) = RS!FecFactu
            ItmX.SubItems(3) = RS!Error
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing

ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub

Private Sub PonerFrameCobrosPtesVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Cobros Pendientes Visible y Ajustado al Formulario, y visualiza los controles
'necesario para el Informe

    H = 4600
        
    Select Case OpcionMensaje
        Case 0, 1, 2, 3
            H = 6000
            W = 9200
'            Me.Label1(0).Top = 4800
'            Me.Label1(0).Left = 3400
            Me.CmdSalir.Caption = "&Salir"
            PonerFrameVisible Me.FrameErrores, visible, H, W
            Me.frameAcercaDE.visible = False
            Me.FrameCobrosPtes.visible = False
            
        
        Case 10  'Errores al contabilizar facturas
            H = 6000
            W = 8400
            Me.CmdAceptarCobros.Top = 5300
            Me.CmdAceptarCobros.Left = 4900
    
        Case 6 ' Acerca de
            H = 4485
            W = 5415
            Me.Width = W
            Me.Height = H
            Me.Shape1.Width = W
            Me.Shape1.Height = H
            Me.Shape1.Top = 0
            Me.Shape1.Left = 0
            Me.frameAcercaDE.visible = True
            Me.frameAcercaDE.Left = 5
            Me.frameAcercaDE.Top = 5
            Me.frameAcercaDE.Width = W - 5
            Me.frameAcercaDE.Height = H - 5
            Me.FrameCobrosPtes.visible = False
            Me.FrameErrores.visible = False
        
'            PonerFrameVisible Me.frameAcercaDE, visible, h - 20, w - 20

            Exit Sub
    End Select
            
    
End Sub


Private Sub CargarListaErrComprobacion()
'Muestra la lista Detallada de Facturas que dieron error al contabilizar
'en un ListView
Dim RS As ADODB.Recordset
Dim ItmX As ListItem
Dim Sql As String

    On Error GoTo ECargarListErrComprobacion

    Sql = " SELECT  * "
    Sql = Sql & " FROM tmperrcomprob "
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
'        ListView1.Height = 4500
'        ListView1.Width = 7400
'        ListView1.Left = 500
'        ListView1.Top = 500

        'Los encabezados
        ListView2.ColumnHeaders.Clear

        Select Case OpcionMensaje
            Case 1
                ListView2.ColumnHeaders.Add , , "Error en letra de serie", 6000
            Case 2
                ListView2.ColumnHeaders.Add , , "Error en cuentas contables", 6000
            Case 3
                ListView2.ColumnHeaders.Add , , "Error en tipos de iva", 6000
        
        End Select


'        ListView2.ColumnHeaders.Add , , "Error de comprobación", 5000
'
'        If RS.Fields(0).Name = "codprove" Then
'            'Facturas de Compra
'             ListView1.ColumnHeaders.Add , , "Prove.", 700
'        Else 'Facturas de Venta
'            ListView1.ColumnHeaders.Add , , "Tipo", 600
'        End If
'        ListView1.ColumnHeaders.Add , , "Factura", 1000, 1
'        ListView1.ColumnHeaders.Add , , "Fecha", 1100, 1
'        ListView1.ColumnHeaders.Add , , "Error", 4620
    
        While Not RS.EOF
            Set ItmX = ListView2.ListItems.Add
            'El primer campo será codtipom si llamamos desde Ventas
            ' y será codprove si llamamos desde Compras
            ItmX.Text = RS.Fields(0).Value
'            ItmX.SubItems(1) = Format(RS!NumFactu, "0000000")
'            ItmX.SubItems(2) = RS!FecFactu
'            ItmX.SubItems(3) = RS!Error
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing

ECargarListErrComprobacion:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub

Private Sub CargaImagen()
On Error Resume Next
     Image2.Picture = LoadPicture(App.Path & "\logo.jpg")
'    Image2.Picture = LoadPicture(App.path & "\images\minilogo.bmp")
    'Image1.Picture = LoadPicture(App.path & "\images\fondon.gif")
    Err.Clear
End Sub
