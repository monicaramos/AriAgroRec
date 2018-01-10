VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBuscaGrid 
   Caption         =   "Formulario de búsqueda"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7605
   Icon            =   "frmBuscaGrid.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmBuscaGrid.frx":058A
      Height          =   3590
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   6324
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   1
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowFocus      =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   5160
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
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
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   435
      Left            =   4440
      TabIndex        =   2
      Top             =   5100
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   435
      Left            =   6000
      TabIndex        =   3
      Top             =   5100
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   7275
   End
   Begin VB.Label Label3 
      Caption         =   "Búsqueda"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "TITULO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   7215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cargando datos ...."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1560
      TabIndex        =   5
      Top             =   2640
      Width           =   3675
   End
End
Attribute VB_Name = "frmBuscaGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
' +-+-    Autor: DAVID      +-+-
' +-+- Alguns canvis: CÈSAR +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-

Option Explicit
Public Event Selecionado(CadenaDevuelta As String)

'Variables publicas para montar datos
Public vtabla As String
Public vCampos As String 'columnas en la tabla.Empipados
Public vSelElem As Integer
Public vTitulo As String
Public vSQL As String
'Dentro de campos vendra cada grupo separado por ·
'Y cada grupo sera Desc|Tabla|Tipo|Porcentaje de ancho
Public vDevuelve As String 'Empipados los campos que devuelve



'Variables privadas
Dim PrimeraVez As Boolean
Dim SQL As String
'Las redimensionaremos
Dim TotalArray As Integer
Dim Cabeceras() As String
Dim CabTablas() As String
Dim CabAncho() As Single
Dim TipoCampo() As String
Dim FormatoCampo() As String 'Formato del campo
Private Busca As Boolean
Private DbClick As Boolean



Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Dim columna As String
Dim J As Integer

    If Busca Then
        Busca = False
        Exit Sub
    End If
    DbClick = True
    If adodc1.Recordset.BOF Then
        If adodc1.Recordset.RecordCount > 0 Then adodc1.Recordset.MoveFirst
    End If

    If adodc1.Recordset.RecordCount > 0 Then
        columna = CabTablas(vSelElem)
        J = InStr(1, columna, ".")
        If J > 0 Then columna = Mid(columna, J + 1)
        '---- Añade: Laura 28/04/2005
        J = InStr(1, columna, " as ") 'si columna tiene if o case renombramos ( as nomcolum )
        If J > 0 Then
            columna = Mid(columna, J + 4)
            columna = Trim(columna)
        End If
        '---- Modifica: LAura 27/04/2005 ------------------------
        '---- se añade el formato del campo
        'Antes:
        'Text1.Text = Adodc1.Recordset.Fields(columna)
        If FormatoCampo(vSelElem) <> "" Then
            Text1.Text = Format(adodc1.Recordset.Fields(columna), FormatoCampo(vSelElem))
        Else
            Text1.Text = DBLet(adodc1.Recordset.Fields(columna), TipoCampo(vSelElem))
        End If
        '--------------------------------------------------------
    End If
End Sub

Private Sub cmdRegresar_Click()
Dim vDes As String
Dim I, J As Integer
Dim k As Byte
Dim V As String
Dim NomColum As String

If adodc1.Recordset Is Nothing Then Exit Sub
If adodc1.Recordset.EOF Then Exit Sub
I = 0
vDes = ""
Do
    J = I + 1
    I = InStr(J, vDevuelve, "|")
    If I > 0 Then
        V = Mid(vDevuelve, J, I - J)
        If V <> "" Then
            If IsNumeric(V) Then
                NomColum = CabTablas(Val(V))
                k = InStr(1, NomColum, ".")
                If k > 0 Then NomColum = Mid(NomColum, k + 1)
                If Val(V) <= TotalArray Then vDes = vDes & adodc1.Recordset(NomColum) & "|"
            End If
        End If
    End If
Loop Until I = 0
RaiseEvent Selecionado(vDes)
Unload Me
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub DataGrid1_DblClick()
If adodc1.Recordset Is Nothing Then Exit Sub
If adodc1.Recordset.EOF Then Exit Sub
cmdRegresar_Click
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
Dim cad As String

If adodc1.Recordset Is Nothing Then Exit Sub
If adodc1.Recordset.EOF Then Exit Sub
If vSelElem = ColIndex Then Exit Sub
'cad = "¿Desea reordenar por el concepto " & DataGrid1.Columns(ColIndex).Caption & "?"
'If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
If ColIndex <= TotalArray Then
    Me.Refresh
    DoEvents
    Screen.MousePointer = vbHourglass
    vSelElem = ColIndex
    CargaGrid
    Screen.MousePointer = vbDefault
    Else
    MsgBox "Error cargando tabla. Imposible ordenacion", vbCritical
End If
End Sub

Private Sub Form_Activate()
Dim OK As Boolean
If PrimeraVez Then
    PrimeraVez = False
    Screen.MousePointer = vbHourglass
    OK = ObtenerTamanyosArray
    If OK Then OK = SeparaCampos
    If Not OK Then
        'Error en SQL
        'Salimos
        Unload Me
        Exit Sub
    End If
    CargaGrid
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
PrimeraVez = True
Label1.Caption = vTitulo
DbClick = True
'Adodc1.Password = vUsu.Passwd
End Sub

Private Function SeparaCampos() As Boolean
Dim cad As String
Dim Grupo As String
Dim I As Integer
Dim J As Integer
Dim c As Integer 'Contrador dentro del array

SeparaCampos = False
I = 0
c = 0
Do
    J = I + 1
    I = InStr(J, vCampos, "·")
    If I > 0 Then
        Grupo = Mid(vCampos, J, I - J)
        'Y en la martriz
        InsertaGrupo Grupo, c
        c = c + 1
    End If
Loop Until I = 0
SeparaCampos = True
End Function

Private Sub InsertaGrupo(Grupo As String, Contador As Integer)
Dim I As Integer
Dim J As Integer
Dim cad As String
J = 0


    cad = ""
    
    'Cabeceras
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
        J = 1
    End If
    Cabeceras(Contador) = cad
    
    'TAblas BD
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
    Else
        cad = ""
        Grupo = ""
    End If
    
    CabTablas(Contador) = cad
    
    'El tipo
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
        Else
            cad = ""
            Grupo = ""
    End If
    TipoCampo(Contador) = cad
    
    'El formato
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
    Else
        cad = ""
        Grupo = ""
    End If
    FormatoCampo(Contador) = cad
    
    'Por ultimo
    'ANCHO
    If Grupo = "" Then Grupo = 0
    CabAncho(Contador) = Grupo
End Sub

Private Function ObtenerTamanyosArray() As Boolean
Dim I As Integer
Dim J As Integer
Dim Grupo As String

ObtenerTamanyosArray = False
'Primero a los campos de la tabla
TotalArray = -1
J = 0
Do
    I = J + 1
    J = InStr(I, vCampos, "·")
    If J > 0 Then TotalArray = TotalArray + 1
Loop Until J = 0
If TotalArray < 0 Then Exit Function
'Las redimensionaremos
ReDim Cabeceras(TotalArray)
ReDim CabTablas(TotalArray)
ReDim CabAncho(TotalArray)
ReDim TipoCampo(TotalArray)
ReDim FormatoCampo(TotalArray)
ObtenerTamanyosArray = True
End Function


Private Sub CargaGrid()
Dim cad As String, Orden As String
Dim I As Integer
Dim anc As Single


    'On Error GoTo ECargaGRid '##QUITAR
    'Generamos SQL
    cad = ""
    For I = 0 To TotalArray
        If cad <> "" Then cad = cad & ","
        cad = cad & CabTablas(I)
    Next I
    cad = "SELECT " & cad & " FROM " & vtabla
    If vSQL <> "" Then cad = cad & " WHERE " & vSQL
    '---- Modifica: Laura 28/04/2005  ----------------------
    'antes:
    'cad = cad & " ORDER BY " & CabTablas(vSelElem)
    Orden = CabTablas(vSelElem)
    I = InStr(1, Orden, " as ")
    If I > 0 Then Orden = Mid(Orden, I + 4)
    cad = cad & " ORDER BY " & Orden
    '--------------------------------------------------------
    
    DataGrid1.AllowRowSizing = False
    adodc1.ConnectionString = conn
    adodc1.RecordSource = cad
    adodc1.Refresh
    
    DataGrid1.visible = True
    'Cargamos el grid
    'anc = DataGrid1.Width - 640
    anc = DataGrid1.Width - 582

    For I = 0 To TotalArray
        DataGrid1.Columns(I).AllowSizing = False
        DataGrid1.Columns(I).Caption = Cabeceras(I)
        If FormatoCampo(I) <> "" Then
            DataGrid1.Columns(I).NumberFormat = FormatoCampo(I)
        End If
        If CabAncho(I) = 0 Then
            DataGrid1.Columns(I).visible = False
            Else
            DataGrid1.Columns(I).Width = anc * (CabAncho(I) / 100)
        End If
    Next I


    'Habilitamos el text1 para que escriban
    Text1.Enabled = True
    If Not adodc1.Recordset.EOF Then
        'Le ponemos el 1er registro
        cad = CabTablas(vSelElem)
        'si hay punto en nombre columa lo quitamos: tabla.colum -> colum
        I = InStr(1, cad, ".")
        If I > 0 Then cad = Mid(cad, I + 1)
        
        '---- Añade: LAura 28/04/2005
        'Si hay if/case en nombre columna cogemos el renombrado: if(colum=x,,) as colum
        I = InStr(1, cad, " as ")
        If I > 0 Then
            cad = Mid(cad, I + 4)
            cad = Trim(cad)
        End If
        
        '---- Modifica: Laura 27/04/2005 --------------
        '---- se añade el formato del campo
        'antes:
        'Text1.Text = Adodc1.Recordset(cad)
        If FormatoCampo(vSelElem) <> "" Then
            Text1.Text = Format(adodc1.Recordset(cad), FormatoCampo(vSelElem))
        Else
            'Text1.Text = DBSet(Adodc1.Recordset(cad), TipoCampo(vSelElem))
            Text1.Text = DBLet(adodc1.Recordset(cad), TipoCampo(vSelElem))
        End If
        '-----------------------------------------------
        PonerFoco Text1
    Else
        PonerFocoBtn cmdSalir
    End If
    
    DataGrid1.Columns(vSelElem).Caption = Cabeceras(vSelElem) & " (*)"
Exit Sub
ECargaGrid:
    MuestraError Err.Number, "Carga grid." & vbCrLf & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DataGrid1.Enabled = False
End Sub

Private Sub Text1_Change()
Dim SQLDBGRID As String
Dim I As Byte

    If DbClick Then
        DbClick = False
        Exit Sub
    End If
    Busca = True
    SQLDBGRID = CabTablas(vSelElem)
    
    '---- Añade: Laura 16/03/2006
    I = InStr(1, SQLDBGRID, " as ") 'si columna tiene if o case renombramos ( as nomcolum )
    If I > 0 Then
        SQLDBGRID = Mid(SQLDBGRID, I + 4)
        SQLDBGRID = Trim(SQLDBGRID)
    End If
    
    '----- si hay punto en nombre columa lo quitamos: tabla.colum -> colum
    I = InStr(1, SQLDBGRID, ".")
    If I > 0 Then SQLDBGRID = Mid(SQLDBGRID, I + 1)
    '-----------------------------------------------------------------------
    
    Select Case TipoCampo(vSelElem)
        Case "N"
            If Not IsNumeric(Text1.Text) Then
                If adodc1.Recordset.RecordCount > 0 Then adodc1.Recordset.MoveFirst
                Exit Sub
            End If
            '---- Modifica: Laura 27/04/2005  -------------------
            '---- se añade el formato
            'antes:
            'SQLDBGRID = SQLDBGRID & " >= " & Trim(Text1)
             If Len(Trim(Text1)) > Len(FormatoCampo(vSelElem)) Then
                SQLDBGRID = SQLDBGRID & " >= " & Val(Mid(Trim(Text1), 1, Len(FormatoCampo(vSelElem))))
            Else
                SQLDBGRID = SQLDBGRID & " >= " & Val(Trim(Text1))
            End If
            '-----------------------------------------------------
        Case "T"
            SQLDBGRID = SQLDBGRID & " >= '" & Trim(Text1) & "'"
    End Select
    Screen.MousePointer = vbHourglass
    
    
        
    adodc1.Recordset.Find SQLDBGRID, , adSearchForward, 1
    Screen.MousePointer = vbDefault
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdRegresar_Click
End If
End Sub

' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
Private Sub DataGrid1_GotFocus()
  WheelHook DataGrid1
End Sub
Private Sub DataGrid1_LostFocus()
  WheelUnHook
End Sub

