VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmEntKilosEst 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entrada de Kilos Estimados"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13005
   ClipControls    =   0   'False
   Icon            =   "frmEntKilosEst.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   13005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame5 
      Caption         =   "TOTALES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Left            =   3540
      TabIndex        =   22
      Top             =   5790
      Width           =   5355
      Begin VB.TextBox txtAux2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1380
         TabIndex        =   24
         Top             =   210
         Width           =   1350
      End
      Begin VB.TextBox txtAux2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   3810
         TabIndex        =   23
         Top             =   210
         Width           =   1350
      End
      Begin VB.Label Label4 
         Caption         =   "Hanegadas: "
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   450
         TabIndex        =   26
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "Hectáreas: "
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   2940
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   660
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Tag             =   "Recolectado|N|N|0|1|rentradas|recolect||N|"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Frame Frame2 
      Height          =   620
      Left            =   120
      TabIndex        =   12
      Top             =   675
      Width           =   12765
      Begin VB.CheckBox Check1 
         Caption         =   "Ordenado por Campo"
         Height          =   315
         Left            =   5760
         TabIndex        =   21
         Top             =   180
         Width           =   1965
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   7980
         TabIndex        =   16
         Top             =   120
         Width           =   4575
         Begin VB.OptionButton Option1 
            Caption         =   "Cooperativa"
            Height          =   225
            Index           =   0
            Left            =   210
            TabIndex        =   20
            Top             =   120
            Width           =   1275
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Sigpac"
            Height          =   225
            Index           =   1
            Left            =   1560
            TabIndex        =   19
            Top             =   120
            Width           =   795
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Catastro"
            Height          =   225
            Index           =   2
            Left            =   2460
            TabIndex        =   18
            Top             =   120
            Width           =   945
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Cultivado"
            Height          =   225
            Index           =   3
            Left            =   3480
            TabIndex        =   17
            Top             =   120
            Width           =   1065
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   435
         Left            =   7770
         TabIndex        =   15
         Top             =   120
         Width           =   4875
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   1110
         MaxLength       =   16
         TabIndex        =   0
         Tag             =   "Cod.Socio|N|N|0|999999|rcampos|codsocio|000000|N|"
         Text            =   "Text1"
         Top             =   200
         Width           =   840
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   0
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text2"
         Top             =   200
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Socio"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   195
         Width           =   495
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   810
         ToolTipText     =   "Buscar socio"
         Top             =   210
         Width           =   255
      End
   End
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   320
      Left            =   2130
      Locked          =   -1  'True
      MaxLength       =   16
      TabIndex        =   5
      Tag             =   "Canaforo|N|N|||rcampos|canaforo|###,###,##0|N|"
      Text            =   "canaforo"
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   5805
      Width           =   3255
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   240
         TabIndex        =   11
         Top             =   180
         Width           =   2355
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   10620
      TabIndex        =   1
      Top             =   5880
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   11805
      TabIndex        =   2
      Top             =   5895
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   11790
      TabIndex        =   3
      Top             =   5880
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   13005
      _ExtentX        =   22939
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Incrementar/Decrementar Aforo"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   6600
         TabIndex        =   9
         Top             =   90
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmEntKilosEst.frx":000C
      Height          =   4395
      Left            =   120
      TabIndex        =   6
      Top             =   1350
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   7752
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   8880
      Top             =   5160
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
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
   Begin VB.Label Label10 
      Caption         =   "Cargando datos ........."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "frmEntKilosEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmSoc As frmManSocios 'Socios
Attribute frmSoc.VB_VarHelpID = -1

Private Modo As Byte

Dim kCampo As Integer

Dim CadenaConsulta As String

Dim gridCargado As Boolean 'Si el DataGrid ya tiene todos los Datos cargados.
                           'Para el RowColChange, si el grid no esta totalmente cargado el CargaTxtAux da error.

Dim PrimeraVez As Boolean

Private HaDevueltoDatos As Boolean


Private Sub cmdAceptar_Click()
Dim Cad As String

    On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
        Case 1 'Buscar registros a inventariar
            If Text1(0).Text <> "" Then
                CargaGrid True
                'Poner Modo Modificar columna Existencia Real
                If Not Data1.Recordset.EOF Then
                    PonerModo 4
                    CargaTxtAux True, True
                Else 'No existen registros en la tabla sinven para ese criterio de búsqueda
                    Cad = "No hay campos de alta para este socio."
                    MsgBox Cad, vbInformation
                    PonerFoco Text1(0)
                End If
            Else
                Cad = "Criterio de Búsqueda incompleto." & vbCrLf
                Cad = Cad & "Debe introducir el socio "
                MsgBox Cad, vbExclamation
                PonerFoco Text1(0)
            End If
            
        Case 4 'Modificar el campo de aforo
            CargaTxtAux False, False
            PonerModo 2
    End Select
    Screen.MousePointer = vbDefault
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
    On Error GoTo ECancelar
    
     
            
    Select Case Modo
        Case 1 'Buscar
            LimpiarCampos
            PonerModo 2
            CargaTxtAux False, False
            CargaGrid False
        Case 4  ' 4: Modificar
            PonerModo 2
            CargaTxtAux False, False
            CargaGrid True
    End Select
    
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Data1.Recordset.EOF And gridCargado And Modo = 4 Then
       CargaTxtAux True, True
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim i As Integer
    
'    'Icono del formulario
'    Me.Icon = frmPpal.Icon

    'ICONOS de La toolbar
    With Toolbar1
        .ImageList = frmPpal.imgListComun
        .DisabledImageList = frmPpal.imgListComun_BN
        
        'ASignamos botones
        .Buttons(1).Image = 1 'Buscar
        .Buttons(4).Image = 4 'Modificar
        .Buttons(5).Image = 31 'Incrementar / decrementar porcentaje
        .Buttons(6).Image = 11 'Salir
    End With
    
    For i = 0 To imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    CargaCombo
    LimpiarCampos   'Limpia los campos TextBox
    PrimeraVez = True

    Option1(0).Value = True
    Option1(1).Value = False
    Option1(2).Value = False

    PonerModo 0
    CargaGrid (Modo = 2)
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaGrid(enlaza As Boolean)
Dim i As Byte
Dim Sql As String
On Error GoTo ECarga

    gridCargado = False
    
    Sql = MontaSQLCarga(enlaza)
    CargaGridGnral DataGrid1, Me.Data1, Sql, PrimeraVez
    
    PrimeraVez = False
        
        
    '[Monica]26/08/2011: Modificacion si es Picassent quiere que saque el nro de orden
    'Cod. Campo
    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
        DataGrid1.Columns(0).Caption = "Huerto"
        DataGrid1.Columns(0).Width = 1000
        DataGrid1.Columns(0).NumberFormat = "0000"
        DataGrid1.Columns(0).Alignment = dbgCenter
    Else
        DataGrid1.Columns(0).Caption = "Campo"
        DataGrid1.Columns(0).Width = 1000
        DataGrid1.Columns(0).NumberFormat = "00000000"
        DataGrid1.Columns(0).Alignment = dbgCenter
    End If
    
    
    DataGrid1.Columns(0).Width = 1000
    DataGrid1.Columns(0).NumberFormat = "00000000"
    DataGrid1.Columns(0).Alignment = dbgCenter
    
    
    
    DataGrid1.Columns(1).Caption = "Partida"
    DataGrid1.Columns(1).Width = 2000
    DataGrid1.Columns(1).Alignment = dbgCenter
       
    'Variedad
    DataGrid1.Columns(2).Caption = "Variedad"
    DataGrid1.Columns(2).Width = 1800
    DataGrid1.Columns(2).Alignment = dbgCenter
    
    'Hanegadas
    DataGrid1.Columns(3).Caption = "Hanegadas"
    DataGrid1.Columns(3).Width = 1100
    DataGrid1.Columns(3).NumberFormat = "##,###,##0.00"
    DataGrid1.Columns(3).Alignment = dbgRight
    
    'Hectareas
    DataGrid1.Columns(4).Caption = "Hectareas"
    DataGrid1.Columns(4).Width = 1100
    DataGrid1.Columns(4).NumberFormat = "##,###,##0.0000"
    DataGrid1.Columns(4).Alignment = dbgRight
    
    'Arboles
    DataGrid1.Columns(5).Caption = "Arboles"
    DataGrid1.Columns(5).Width = 1100
    DataGrid1.Columns(5).NumberFormat = "##,###,##0"
    DataGrid1.Columns(5).Alignment = dbgCenter
    
    'Recolectado
    DataGrid1.Columns(6).visible = False
    
    'descripcion Recolectado
    DataGrid1.Columns(7).Caption = "Recolect"
    DataGrid1.Columns(7).Width = 1100
    DataGrid1.Columns(7).Alignment = dbgCenter
    
    'Aforo
    DataGrid1.Columns(8).Caption = "Aforo"
    DataGrid1.Columns(8).Width = 1500
    DataGrid1.Columns(8).NumberFormat = "##,###,##0"
    DataGrid1.Columns(8).Alignment = dbgCenter
    
    'Aforo Real
    DataGrid1.Columns(9).Caption = "Aforo Real"
    DataGrid1.Columns(9).Width = 1400
    DataGrid1.Columns(9).NumberFormat = "##,###,##0"
    DataGrid1.Columns(9).Alignment = dbgCenter
    
    '[Monica]26/08/2011: Modificacion si es Picassent el campo lo pongo al final
    'Cod. Campo
    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
        DataGrid1.Columns(10).Caption = ""
        DataGrid1.Columns(10).Width = 0
        DataGrid1.Columns(10).Alignment = dbgCenter
    End If
    
    
    For i = 0 To DataGrid1.Columns.Count - 1
        DataGrid1.Columns(i).AllowSizing = False
        DataGrid1.Columns(i).Locked = True
    Next i
    
    DataGrid1.ScrollBars = dbgAutomatic
    gridCargado = True
    
    CalcularTotales Sql
    
    
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag, Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posición adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        txtAux.Top = 290
        txtAux.visible = visible
        Combo1(1).Top = 290
        Combo1(1).visible = visible
    Else
        DeseleccionaGrid Me.DataGrid1
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
'                Combo1(1).ListIndex = ValorCombo(Combo1(1))
                PosicionarCombo Combo1(1), Data1.Recordset!Recolect
                Combo1(1).Locked = False
                txtAux.Text = DBLet(Data1.Recordset!Canaforo, "N")
                txtAux.Text = Format(txtAux.Text, "###,###,##0")
                txtAux.Locked = False
        End If

        If DataGrid1.Row < 0 Then
            alto = DataGrid1.Top + 220
        Else
            alto = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) + 20
        End If
        
        'Fijamos altura y posición Top
        '-------------------------------
        Combo1(1).Top = alto
'        Combo1(1).Height = DataGrid1.RowHeight
        txtAux.Top = alto
        txtAux.Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        Combo1(1).Left = DataGrid1.Columns(7).Left + 130 'codalmac
        Combo1(1).Width = DataGrid1.Columns(7).Width - 10
        txtAux.Left = DataGrid1.Columns(8).Left + 130 'codalmac
        txtAux.Width = DataGrid1.Columns(8).Width - 10
        
        'Los ponemos Visibles o No
        '--------------------------
        Combo1(1).visible = visible
        txtAux.visible = visible
    End If
    PonerFocoCmb Combo1(1)
    If visible Then
        Combo1(1).TabIndex = 2
        txtAux.TabIndex = 3
        txtAux.SelStart = 0
        txtAux.SelLength = Len(txtAux.Text)
    Else
        Combo1(1).TabIndex = 5
        txtAux.TabIndex = 6
    End If
    
'    PonerFoco txtAux
    
'    If visible Then
'        txtAux.TabIndex = 2
'        txtAux.SelStart = 0
'        txtAux.SelLength = Len(txtAux.Text)
'    Else
'        txtAux.TabIndex = 5
'    End If
End Sub


Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
'Socios
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(0)
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Then Exit Sub
 
    Screen.MousePointer = vbHourglass
    imgBuscar(0).Tag = Index
    
    Select Case Index
        Case 0 'Codigo Socio
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
            frmSoc.Show vbModal
            Set frmSoc = Nothing
    End Select
    PonerFoco Text1(Index)
    Screen.MousePointer = vbDefault
End Sub


Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim campo As String
Dim Tabla As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    
    If Text1(Index).Text = "" Then
        Text2(Index).Text = ""
    Else
        Select Case Index
            Case 0 'Codigo Socio
                campo = "nomsocio"
                Tabla = "rsocios"
        End Select
        Text2(Index).Text = PonerNombreDeCod(Text1(Index), Tabla, campo)
        If Text1(Index).Text <> "" And Text2(Index).Text = "" Then PonerFoco Text1(Index)
     End If
End Sub


Private Sub txtAux_GotFocus()
    ConseguirFocoLin txtAux
End Sub

Private Sub txtAux_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo EKeyD
    If KeyCode = 38 Or KeyCode = 40 Then
        ModificarExistencia
    End If

    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
                If DataGrid1.Bookmark > 0 Then
                    DataGrid1.Bookmark = DataGrid1.Bookmark - 1
                End If
        
'                If DataGrid1.Row > 0 Then
'                    DataGrid1.Row = DataGrid1.Row - 1
'                    CargaTxtAux True, True
''                elseif
'                End If
        Case 40 'Desplazamiento Flecha Hacia Abajo
'                PasarSigReg
                 txtAux_KeyPress (13)
               
    End Select
    
    
EKeyD:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub txtAux_KeyPress(KeyAscii As Integer)

'    KEYpress KeyAscii
    
   If KeyAscii = 13 Then 'ENTER
        ModificarExistencia
        PasarSigReg
   ElseIf KeyAscii = 27 Then
        cmdCancelar_Click 'ESC
   End If
End Sub


Private Sub txtAux_LostFocus()
    If Screen.ActiveControl.Name = "cmdCancelar" Then Exit Sub
    With txtAux
        .Text = Trim(.Text)
        If .Text = "" Then Exit Sub
        'Formato tipo 1: Decimal(12,2)
        .Text = Format(.Text, "###,###,##0")
    End With
    
    
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Busqueda
            BotonBuscar
        Case 4 'Modificar
            BotonModificar
        Case 5 'Incrementar/decrementar porcentaje
            BotonIncrementarDecrementar
        Case 6 'Salir
            Unload Me
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte
Dim b As Boolean
       
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    'b = (Kmodo = 2)
   
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
'    BloquearText1 Me, Modo
    b = (Modo <> 1)
    BloquearTxt Text1(0), b
    
    b = (Modo = 0) Or (Modo = 2)
    PonerBotonCabecera b
   
    Select Case Kmodo
'    Case 0    'Modo Inicial
'        PonerBotonCabecera True
'        lblIndicador.Caption = ""
        
    Case 1 'Modo Buscar
'        PonerBotonCabecera False
        PonerFoco Text1(0)
'        lblIndicador.Caption = "BÚSQUEDA"
'    Case 2    'Visualización de Datos
'        PonerBotonCabecera True
'    Case 3 'Insertar Datos en el Datagrid
'        PonerBotonCabecera False 'Poner Aceptar y Cancelar Visible
'        lblIndicador.Caption = "MODIFICAR"
    End Select
           
    b = Modo <> 0 And Modo <> 2 And Modo <> 4
   
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = b
    Next i

    b = (Modo = 1)
    Toolbar1.Buttons(1).Enabled = Not b
    Toolbar1.Buttons(4).Enabled = Not b And (Not (Modo = 0 Or Modo = 4))

    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


Private Function MontaSQLCarga(enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim Sql As String


    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then  ' si es Picassent quiere el nro de orden
        Sql = "SELECT rcampos.nrocampo, rpartida.nomparti, "
    Else
        Sql = "SELECT rcampos.codcampo, rpartida.nomparti, "
    End If
    Sql = Sql & " variedades.nomvarie, "
    
    If Option1(0).Value Then Sql = Sql & " round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",2) hdas, rcampos.supcoope has, "
    If Option1(1).Value Then Sql = Sql & " round(rcampos.supsigpa / " & DBSet(vParamAplic.Faneca, "N") & ",2) hdas, rcampos.supsigpa has, "
    If Option1(2).Value Then Sql = Sql & " round(rcampos.supcatas / " & DBSet(vParamAplic.Faneca, "N") & ",2) hdas, rcampos.supcatas has, "
    If Option1(3).Value Then Sql = Sql & " round(rcampos.supculti / " & DBSet(vParamAplic.Faneca, "N") & ",2) hdas, rcampos.supculti has, "
        
    Sql = Sql & " rcampos.nroarbol,recolect, CASE rcampos.recolect when 0 then ""Cooper"" when 1 then ""Socio"" end as desrecolect, "
    Sql = Sql & " round(rcampos.canaforo * (1 + " & DBSet(vParamAplic.PorcIncreAforo, "N") & "/ 100 ), 0) canaforo, rcampos.canaforo canafororeal "
    
    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
        Sql = Sql & ", codcampo "
    End If
    Sql = Sql & " FROM (rcampos INNER JOIN rpartida on rcampos.codparti=rpartida.codparti) INNER JOIN variedades ON rcampos.codvarie=variedades.codvarie "
    
    If enlaza Then
        Sql = Sql & " WHERE rcampos.codsocio = " & Text1(0).Text & " and rcampos.fecbajas is null "
    Else
        Sql = Sql & " WHERE rcampos.codsocio = -1"
    End If

'    SQL = SQL & " ORDER BY rcampos.codsocio, rcampos.codcampo"
    If Check1.Value = 1 Then
        Sql = Sql & " ORDER BY 1 "
    Else
        Sql = Sql & " ORDER BY 3 "
    End If
    MontaSQLCarga = Sql
End Function


Private Sub BotonBuscar()
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Ponemos el grid lineasfacturas enlazando a ningun sitio
        CargaGrid False
        CargaTxtAux False, False
        Text1(0).BackColor = vbYellow
    Else
        'Ya estamos en Modo de Busqueda
'        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonModificar()
    If Data1.Recordset.EOF Then Exit Sub
    PonerModo 4
    CargaTxtAux True, True
End Sub


Private Sub BotonIncrementarDecrementar()
    AbrirListadoTomaDatos (4)
End Sub


Private Function DatosOk() As Boolean
'Solo se actualiza el campo de Existencia Real
    txtAux.Text = Trim(txtAux.Text)

    If txtAux.Text <> "" And EsNumerico(txtAux.Text) Then
        If PonerFormatoEntero(txtAux) Then
            If AforoSuperiorARdtoMaximo(txtAux) Then
                If MsgBox("El aforo introducido es superior al rendimiento máximo por Hanegada." & vbCrLf & " ¿ Desea continuar ? ", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                    DatosOk = False
                    Exit Function
                End If
            End If
            DatosOk = True
        Else
            DatosOk = False
        End If
        'DatosOk = True
    Else
        DatosOk = False
    End If
End Function


Private Sub PonerBotonCabecera(b As Boolean)
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    If b Then Me.lblIndicador.Caption = ""
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Function ActualizarExistencia(Recolec As Byte, Canti As String) As Boolean
'Actualiza la cantidad de stock Inventariada (Existencia Real en Almacen)
Dim Sql As String
Dim ADonde As String

    On Error GoTo EActualizar

    conn.BeginTrans
    'Actualizar la Tabla: rcampos con la cantidad introducida
    '-------------------------------------------------------
    ADonde = "Modificando Kilos Estimados(Tabla: rcampos)."
    Sql = "UPDATE rcampos Set recolect = " & DBSet(Combo1(1).ListIndex, "N") & ",canaforo = " & DBSet(Canti, "N")
    Sql = Sql & " WHERE codcampo =" & Data1.Recordset!codcampo
    
'    SQL = SQL & " AND codsocio =" & Val(Text1(0).Text)
    conn.Execute Sql
    
    ActualizarExistencia = True
        
EActualizar:
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
         Sql = "Actualizando Recolectado y Kilos Estimados." & vbCrLf & "--------------------------------------------" & vbCrLf
         Sql = Sql & ADonde
         MuestraError Err.Number, Sql, Err.Description
         conn.RollbackTrans
         ActualizarExistencia = False
    Else
        ActualizarExistencia = True
        conn.CommitTrans
    End If
End Function


Private Sub PasarSigReg()
'Nos situamos en el siguiente registro
    If DataGrid1.Bookmark < Data1.Recordset.RecordCount Then
'        DataGrid1.Row = DataGrid1.Row + 1
        DataGrid1.Bookmark = DataGrid1.Bookmark + 1
    ElseIf DataGrid1.Bookmark = Data1.Recordset.RecordCount Then
        PonerFocoBtn Me.cmdAceptar
    End If
End Sub


Private Function ModificarExistencia() As Boolean
Dim NumReg As Long
Dim Indicador As String

    If DatosOk Then
        If ActualizarExistencia(Combo1(1).ListIndex, txtAux.Text) Then
            TerminaBloquear
            NumReg = Data1.Recordset.AbsolutePosition
            CargaGrid True
            If SituarDataPosicion(Data1, NumReg, Indicador) Then

            End If
            ModificarExistencia = True
        Else
            ModificarExistencia = False
        End If
    End If
End Function


Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 1 To Combo1.Count - 1
        Combo1(i).Clear
    Next i
    
    'tipo de recoleccion
    Combo1(1).AddItem "Cooper"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Socio"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
  
End Sub

Private Function AforoSuperiorARdtoMaximo(aforo As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset

    AforoSuperiorARdtoMaximo = False
    
    Sql = "select "
    
    If Option1(0).Value Then Sql = Sql & " round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",2) as hanegadas, round(round(rcampos.supcoope / " & DBSet(vParamAplic.Faneca, "N") & ",2) * variedades.rdtomaximo,0) as aforomaximo  "
    If Option1(1).Value Then Sql = Sql & " round(rcampos.supsigpa / " & DBSet(vParamAplic.Faneca, "N") & ",2) as hanegadas, round(round(rcampos.supsigpa / " & DBSet(vParamAplic.Faneca, "N") & ",2) * variedades.rdtomaximo,0) as aforomaximo "
    If Option1(2).Value Then Sql = Sql & " round(rcampos.supcatas / " & DBSet(vParamAplic.Faneca, "N") & ",2) as hanegadas, round(round(rcampos.supcatas / " & DBSet(vParamAplic.Faneca, "N") & ",2) * variedades.rdtomaximo,0) as aforomaximo "
    If Option1(3).Value Then Sql = Sql & " round(rcampos.supculti / " & DBSet(vParamAplic.Faneca, "N") & ",2) as hanegadas, round(round(rcampos.supculti / " & DBSet(vParamAplic.Faneca, "N") & ",2) * variedades.rdtomaximo,0) as aforomaximo "

    Sql = Sql & " from rcampos, variedades where rcampos.codcampo = " & Data1.Recordset!codcampo
    Sql = Sql & " and rcampos.codvarie = variedades.codvarie "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        AforoSuperiorARdtoMaximo = (DBLet(Rs.Fields!aforomaximo, "N") < CLng(aforo))
    End If
    
End Function


Private Sub CalcularTotales(cadena As String)
Dim Hdas  As Currency
Dim Has As Currency

Dim Rs As ADODB.Recordset
Dim Sql As String

    On Error Resume Next
    
    Sql = Sql & "select sum(hdas), sum(has) from (" & cadena & ") aaaaa"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Hdas = 0
    Has = 0
    txtAux2(1).Text = ""
    txtAux2(2).Text = ""
    
    If TotalRegistrosConsulta(cadena) = 0 Then Exit Sub
    
    If Not Rs.EOF Then
        If Rs.Fields(0).Value <> 0 Then Hdas = DBLet(Rs.Fields(0).Value, "N") 'Solo es para saber que hay registros que mostrar
        If Rs.Fields(1).Value <> 0 Then Has = DBLet(Rs.Fields(1).Value, "N") 'Solo es para saber que hay registros que mostrar
    
        txtAux2(1).Text = Format(Hdas, "###,###,##0.00")
        txtAux2(2).Text = Format(Has, "###,###,##0.0000")
    End If
    Rs.Close
    Set Rs = Nothing

    
    DoEvents
    
End Sub


