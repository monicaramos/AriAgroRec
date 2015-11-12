VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFichaTecIMG 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ficha Documentos. IMAGENES"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5760
   Icon            =   "frmFichaTecIMG.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   4590
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "Nombre|T|N|||rsocios|nomsocio|||"
      Top             =   900
      Width           =   705
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   150
      MaxLength       =   40
      TabIndex        =   0
      Tag             =   "Nombre|T|N|||rsocios|nomsocio|||"
      Text            =   "1234567890123456789012345678901234567890"
      Top             =   900
      Width           =   4305
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4260
      TabIndex        =   3
      Top             =   7050
      Width           =   1125
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   3870
      Top             =   510
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1320
      Top             =   5790
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Orden"
      Height          =   255
      Index           =   2
      Left            =   4620
      TabIndex        =   6
      Top             =   570
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "Descripción Fichero"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   4
      Top             =   570
      Width           =   1455
   End
   Begin VB.Label lblCarga2 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   2
      Top             =   150
      Width           =   5385
   End
   Begin VB.Shape Shape1 
      Height          =   5325
      Left            =   180
      Top             =   1560
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   5100
      Left            =   300
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   4965
   End
   Begin VB.Label Label1 
      Caption         =   "Imagen"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   5
      Top             =   1230
      Width           =   1455
   End
End
Attribute VB_Name = "frmFichaTecIMG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const CarpetaIMG = "ImgFicFT"
Public vDatos As String 'codsocio|nomsocio|

Public Opcion As Byte '0=insertar 1=imprimir 2=eliminar

Dim InsertandoImg As Boolean
Dim PrimeraVez As Boolean


Dim It As ListItem
Dim Contador As Integer
Dim fichero As String
Dim TipoDocu As Byte

Private Sub InsertarDesdeFichero()
Dim Cadena As String
Dim Carpeta As String
Dim Aux As String
Dim J As Integer

    fichero = ""
    cd1.FileName = ""
    cd1.InitDir = "c:\"
    cd1.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
    cd1.MaxFileSize = 1024 * 30
    cd1.Filter = "Archivos PDF|*.pdf|Archivos Jpg|*.jpg|Archivos Png|*.png|Archivos TIFF|*.tif"
    cd1.ShowOpen
    cd1.MaxFileSize = 256
    cd1.CancelError = False
    
    If cd1.FileName = "" Then
        Unload Me
        Exit Sub
    End If
    
    If FileLen(cd1.FileName) / 1000 > 1024 Then
        MsgBox "No se permite insertar ficheros de tamaño superior a 1 M", vbExclamation
        Unload Me
        Exit Sub
    End If
    
    
'    '******* Cambiamos cursor
    Screen.MousePointer = vbHourglass
    InsertandoImg = True

    J = InStr(1, cd1.FileName, Chr(0))
    Cadena = cd1.FileName
    TipoDocu = 0
    If InStr(1, cd1.FileName, "pdf") <> 0 Then TipoDocu = 1
    fichero = Cadena
        
            
    CargarIMG (Cadena)
    InsertandoImg = False
    Screen.MousePointer = vbDefault
    
    Text1(0).Text = CCur(DevuelveValor("select max(orden) from rfichdocs where codsocio = " & DBSet(RecuperaValor(vDatos, 1), "N"))) + 1
    Text1(1).Text = Dir(Cadena)
    PonerFoco Text1(1)
End Sub


Private Function CargarIMG(Archivo As String) As Boolean
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    lblCarga2.Caption = "Cargando ..."
    lblCarga2.Refresh
    CargarIMG = False
    If TipoDocu = 1 Then 'InStr(1, Archivo, ".pdf") <> 0 Then
        Me.Image1.Picture = LoadPicture(App.Path & "\pdf.dat")
    Else
        If InStr(1, Archivo, ".tif") <> 0 Then
            Me.Image1.Picture = LoadPicture(App.Path & "\tif.dat")
        Else
            If InStr(1, Archivo, ".png") Then
                Me.Image1.Picture = LoadPicture(App.Path & "\png.dat")
            Else
                Me.Image1.Picture = LoadPicture(Archivo)
            End If
        End If
    End If

    If Err.Number <> 0 Then
        MsgBox Err.Description, vbExclamation
    Else
        CargarIMG = True
    End If
    lblCarga2.Caption = lblCarga2.Tag
    Screen.MousePointer = vbDefault
End Function

Private Function InsertarImagen() As Boolean
Dim RS As ADODB.Recordset
Dim c As String
Dim L As Long
Dim L1 As Long

Dim k As Integer
Dim eliminar As Boolean

    
    On Error GoTo eInsertarImagen
    
    InsertarImagen = False
    
    AbrirConexion
    
    c = "Select max(codigo) from rfichdocs" '  where codsocio = " & RecuperaValor(vDatos, 1)
    Set RS = New ADODB.Recordset
    RS.Open c, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    L = 0
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then L = RS.Fields(0)
    End If
    L = L + 1
    RS.Close
    
    
'    c = "Select max(orden) from rfichdocs where codsocio = " & RecuperaValor(vDatos, 1)
'    Set RS = New ADODB.Recordset
'    RS.Open c, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    L1 = 0
'    If Not RS.EOF Then
'        If Not IsNull(RS.Fields(0)) Then L1 = RS.Fields(0)
'    End If
'    L1 = L1 + 1
'    RS.Close
    
    
    ' es nuevo
    c = "insert into rfichdocs (codigo, codsocio, descripfich, orden, docum) values"
    c = c & " (" & DBSet(L, "N") & "," & RecuperaValor(vDatos, 1) & "," & DBSet(Me.Text1(1).Text, "T") & "," & DBSet(Text1(0).Text, "N") & "," & DBSet(Dir(fichero), "T") & ")"
    conn.Execute c
    
    espera 0.2
    
    'Abro parar guardar el binary
    c = "Select * from rfichdocs where codigo =" & L '& " and codsocio = " & DBSet(RecuperaValor(vDatos, 1), "N")
    Adodc1.ConnectionString = conn
    Adodc1.RecordSource = c
    Adodc1.Refresh
'
    If Adodc1.Recordset.EOF Then
        'MAAAAAAAAAAAAL

    Else
        'Guardar
        InsertandoImg = True
        CargarIMG fichero 'lw1.ListItems(k).SubItems(2)
        GuardarBinary Adodc1.Recordset!campo, fichero
        Adodc1.Recordset.Update
    End If

    InsertarImagen = True
    Exit Function
    
eInsertarImagen:
    MuestraError Err.Number, "Insertar Imágen", Err.Description
End Function


Private Sub cmdGuardar_Click()

    If Text1(1).Text = "" Then
        MsgBox "Debe introducir una descripción de Fichero. Reintroduzca.", vbExclamation
        PonerFoco Text1(1)
        Exit Sub
    End If
    
    If Text1(0).Text = "" Then
        MsgBox "Debe introducir el orden de la imágen en la lista del socio. Reintroduzca.", vbExclamation
        PonerFoco Text1(0)
        Exit Sub
    End If

    If InsertarImagen Then
        MsgBox "Proceso realizado correctamente.", vbExclamation
    
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Screen.MousePointer = vbHourglass
'        ProcesarCarpetaImagenes
        
        If Opcion = 0 Then InsertarDesdeFichero
        
        lblCarga2.Caption = lblCarga2.Tag
        
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    
    PrimeraVez = True
      ' ICONITOS DE LA BARRA
    Me.lblCarga2.Tag = RecuperaValor(Me.vDatos, 2)
    lblCarga2.Caption = "Leyendo datos BD"
    
'    If Opcion = 0 Then InsertarDesdeFichero
    
End Sub




Private Sub ProcesarCarpetaImagenes()
Dim c As String
    On Error GoTo EProcesarCarpetaImagenes
    c = App.Path & "\" & CarpetaIMG
    If Dir(c, vbDirectory) = "" Then
        MkDir c
    Else
        If Dir(c & "\*.*", vbArchive) <> "" Then Kill c & "\*.*"
    End If
    
    Exit Sub
EProcesarCarpetaImagenes:
    MuestraError Err.Number, "ProcesarCarpetaImagenes"
End Sub

Private Sub Imprimir()
        With frmImprimir
            .FormulaSeleccion = "{rsocios.codsocio}=" & RecuperaValor(vDatos, 1)
            .OtrosParametros = "|pNomEmpre=""" & vEmpresa.nomempre & """|"
            .Titulo = "Imágenes adjuntas"
            .NumeroParametros = 1
            .SoloImprimir = False
            .EnvioEMail = False
            .NombreRPT = "rImgDocs.rpt"
            
            .Opcion = 2015
            .Show vbModal
        End With
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
        Unload Me
    End If
End Sub

