VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdf2c 
      Caption         =   "Command2"
      Height          =   315
      Left            =   2490
      TabIndex        =   5
      Top             =   1530
      Width           =   1455
   End
   Begin VB.CommandButton cmdc2f 
      Caption         =   "Command1"
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Top             =   1500
      Width           =   1725
   End
   Begin VB.TextBox Txtf 
      Height          =   315
      Left            =   2400
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox Txturl 
      Height          =   645
      Left            =   240
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   2460
      Width           =   4365
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Left            =   3330
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtc 
      Height          =   315
      Left            =   210
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1920
      Width           =   1935
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1110
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum tipoConversion
    deCaF
    deFaC
End Enum
Private tipo As tipoConversion

' Estas definiciones están tomadas de lo mostrado en el explorador
' al seleccionar cada una de las funciones del servicio Web
Private Const cSOAPCaF = _
    "<?xml version=""1.0"" encoding=""utf-8""?>" & _
    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
        "<soap:Body>" & _
            "<CaF xmlns=""http://elGuille/WebServices"">" & _
                "<valor>1</valor>" & _
            "</CaF>" & _
        "</soap:Body>" & _
    "</soap:Envelope>"
    
Private Const cSOAPFaC = _
    "<?xml version=""1.0"" encoding=""utf-8""?>" & _
    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
        "<soap:Body>" & _
            "<FaC xmlns=""http://elGuille/WebServices"">" & _
                "<valor>1</valor>" & _
            "</FaC>" & _
        "</soap:Body>" & _
    "</soap:Envelope>"

Private Sub cmdC2F_Click()
    tipo = deCaF
    Dim parser As DOMDocument
    Set parser = New DOMDocument
    ' cargar el código SOAP para CaF
    parser.LoadXml cSOAPCaF
    '
    ' Indicar el parámetro a enviar
    parser.selectSingleNode("/soap:Envelope/soap:Body/CaF/valor").Text = txtc.Text
    '
    ' Mostrar el código XML enviado al servicio Web
    Text1.Text = parser.XML
    Txturl.Text = "http://guille.costasol.net/Net/WebServices/conversor.asmx"
    ' Usar el control Inet para realizar la operación HTTP POST
    Inet1.Execute Txturl.Text, "POST", parser.XML, "Content-Type: text/xml; charset=utf-8" & vbCrLf & "SOAPAction: http:""//elGuille/WebServices/CaF"
End Sub

Private Sub cmdF2C_Click()
    tipo = deFaC
    Dim parser As DOMDocument
    Set parser = New DOMDocument
    ' cargar el código SOAP para FaC
    parser.LoadXml cSOAPFaC
    '
    ' Indicar el parámetro a enviar
    parser.selectSingleNode("/soap:Envelope/soap:Body/FaC/valor").Text = Txtf.Text
    '
    ' Mostrar el código XML enviado al servicio Web
    Text1.Text = parser.XML
    ' Usar el control Inet para realizar la operación HTTP POST
    Inet1.Execute Txturl.Text, "POST", parser.XML, "Content-Type: text/xml; charset=utf-8" & vbCrLf & "SOAPAction: http:""//elGuille/WebServices/FaC"
End Sub

Private Sub Form_Load()
    Text1.Text = "Cliente VB6" & vbCrLf & "del Servicio Web Conversor de ºC a ºF"
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
    '
    If (State = icResponseCompleted) Then ' icResponseCompleted = 12
        Dim s As String
        '
        ' Leer los datos devueltos por el servidor
        s = Inet1.GetChunk(4096)
        Text1.Text = s
        '
        ' Poner los datos en el analizador de XML
        Dim parser As DOMDocument
        Set parser = New DOMDocument
        parser.LoadXml s
        '
        On Error Resume Next
        '
        If tipo = deCaF Then
            Txtf.Text = parser.selectSingleNode("/soap:Envelope/soap:Body/CaFResponse/CaFResult").Text
        Else
            txtc.Text = parser.selectSingleNode("/soap:Envelope/soap:Body/FaCResponse/FaCResult").Text
        End If
        '
        If Err.Number > 0 Then
            Text1.SetFocus
        End If
    End If
End Sub

