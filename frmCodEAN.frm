VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCodEAN 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Códigos EAN asociados"
   ClientHeight    =   9930
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   11325
   Icon            =   "frmCodEAN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   135
      TabIndex        =   12
      Top             =   90
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   13
         Top             =   180
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Nuevo"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar"
               Object.Tag             =   "2"
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Buscar"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver Todos"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imprimir"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Salir"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
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
      Height          =   290
      Index           =   3
      Left            =   6885
      MaxLength       =   13
      TabIndex        =   11
      Tag             =   "Ref.Cliente|N|S|||codigoean|refclien|#######0||"
      Top             =   4275
      Width           =   1080
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
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
      Height          =   290
      Index           =   2
      Left            =   5715
      MaxLength       =   13
      TabIndex        =   10
      Tag             =   "Código EAN|T|N|||codigoean|codigoean|||"
      Top             =   4275
      Width           =   1080
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
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
      Height          =   290
      Index           =   0
      Left            =   165
      MaxLength       =   16
      TabIndex        =   9
      Tag             =   "Código Forfaits|T|N|||codigoean|codforfait||S|"
      Top             =   4275
      Width           =   900
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
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
      Height          =   290
      Index           =   1
      Left            =   3105
      MaxLength       =   4
      TabIndex        =   8
      Tag             =   "Variedad|N|N|0|9999|codigoean|codvarie|0000|S|"
      Top             =   4275
      Width           =   540
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1170
      TabIndex        =   7
      Top             =   4275
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   3735
      TabIndex        =   6
      Top             =   4275
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
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
      Left            =   8865
      TabIndex        =   4
      Top             =   9390
      Visible         =   0   'False
      Width           =   1065
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
      Left            =   10035
      TabIndex        =   5
      Top             =   9405
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
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
      Left            =   10035
      TabIndex        =   2
      Top             =   9405
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   9135
      Width           =   3105
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   1
         Top             =   240
         Width           =   2790
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   330
      Left            =   3120
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmCodEAN.frx":000C
      Height          =   8145
      Left            =   135
      TabIndex        =   3
      Top             =   900
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   14367
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
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
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmCodEAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: LAURA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

' **************** PER A QUE FUNCIONE EN UN ATRE MANTENIMENT ********************
' 0. Posar-li l'atribut Datasource a "adodc1" del Datagrid1. Canviar el Caption
'    del formulari
' 1. Canviar els TAGs i els Maxlength de TextAux(0) i TextAux(1)
' 2. En PonerModo(vModo) repasar els indexs del botons, per si es canvien
' 3. En la funció BotonAnyadir() canviar la taula i el camp per a SugerirCodigoSiguienteStr
' 4. En la funció BotonBuscar() canviar el nom de la clau primaria
' 5. En la funció BotonEliminar() canviar la pregunta, les descripcions de la
'    variable SQL i el contingut del DELETE
' 6. En la funció posamaxlength() repasar el maxlength de TextAux(0)
' 7. En Form_Load() repasar la barra d'iconos (per si es vol canviar algún) i
'    canviar la consulta per a vore tots els registres
' 8. En Toolbar1_ButtonClick repasar els indexs de cada botó per a que corresponguen
' 9. En la funció CargaGrid canviar l'ORDER BY (normalment per la clau primaria);
'    canviar ademés els noms dels camps, el format i si fa falta la cantitat;
'    repasar els index dels botons modificar i eliminar.
'    NOTA: si en Form_load ya li he posat clausula WHERE, canviar el `WHERE` de
'    `SQL = CadenaConsulta & " WHERE " & vSQL` per un `AND`
' 10. En txtAux_LostFocus canviar el mensage i el format del camp
' 11. En la funció DatosOk() canviar els arguments de DevuelveDesdeBD i el mensage
'    en cas d'error
' 12. En la funció SepuedeBorrar() canviar les comprovacions per a vore si es pot
'    borrar el registre
' ********************************************************************************

Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'altre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
'Public NuevoCodigo As String

'codigo que tiene el campo en el momento que se llama desde otro formulario
'nos situamos en ese valor
Public CodigoActual As String

Public Tipo As Byte
' Tipo: 0- codigoactual = cliente
'       1- codigoactual = forfait
'       2- codigoactual = variedad

Public Cliente As Long

Private CadenaConsulta As String
Private CadB As String

Dim Modo As Byte
'----------- MODOS ----------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'-----------------------------------------------
Dim PrimeraVez As Boolean


Private Sub PonerModo(vModo)
Dim b As Boolean

    Modo = vModo
'    PonerIndicador lblIndicador, Modo
    b = (Modo = 2)
    
    If b Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    txtAux(0).visible = Not b
    txtAux(1).visible = Not b
    txtAux(2).visible = Not b
    txtAux(3).visible = Not b
    txtAux2(0).visible = Not b
    txtAux2(1).visible = Not b
        
    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    DataGrid1.Enabled = b
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = b
    
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor.
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu 'Activar/Desact botones de menu segun permisos del usuario
    
    'Si estamos modo Modificar bloquear clave primaria
    BloquearTxt txtAux(0), (Modo = 4)
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los txtAux
    PonerLongCamposGnral Me, Modo, 3
End Sub


Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botones del la toobar y del menu, segun el modo en que estemos
Dim b As Boolean

    b = (Modo = 2)
'    mnOpciones.Enabled = b
    'Buscar
    Toolbar1.Buttons(5).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(1).Enabled = False
    Me.mnNuevo.Enabled = False
    
'    b = (b And adodc1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(2).Enabled = False
    Me.mnModificar.Enabled = False
    'Eliminar
    Toolbar1.Buttons(3).Enabled = False
    Me.mnEliminar.Enabled = False
    
    'Imprimir
    'Toolbar1.Buttons(11).Enabled = b
    Toolbar1.Buttons(8).Enabled = False
End Sub


Private Sub BotonVerTodos()
    Select Case Tipo
        Case 0 ' cliente
            CargaGrid "codigoean.codclien=" & CodigoActual
        Case 1 ' forfait
            CargaGrid "codigoean.codforfait=" & DBSet(CodigoActual, "T")
        Case 2 ' variedad
            CargaGrid "codigoean.codvarie=" & CodigoActual
    End Select
            
    CadB = ""
    PonerModo 2
End Sub


Private Sub BotonBuscar()
'    lblIndicador.Caption = "BUSQUEDA"
    ' ***************** canviar per la clau primaria ********
    Select Case Tipo
        Case 0 ' cliente
            CargaGrid "codigoean.codforfait = '-1'"
        Case 1 ' forfait
            CargaGrid "codigoean.codclien = -1"
        Case 2 ' variedad
            CargaGrid "codigoean.codclien = -1"
    End Select
        
    '*******************************************************************************
    'Buscar
    txtAux(0).Text = ""
    txtAux(1).Text = ""
    txtAux(2).Text = ""
    txtAux2(0).Text = ""
    txtAux2(1).Text = ""

    LLamaLineas DataGrid1.Top + 240, 1
    PonerFoco txtAux(0)
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    'Fijamos el ancho
    txtAux(0).Top = alto
    txtAux(1).Top = alto
    txtAux(2).Top = alto
    txtAux(3).Top = alto
    txtAux2(0).Top = alto
    txtAux2(1).Top = alto
End Sub

Private Sub cmdAceptar_Click()
    Select Case Modo
        Case 1  'BUSQUEDA
            CadB = ObtenerBusqueda(Me)
            If CadB <> "" Then
                PonerModo 2
                Select Case Tipo
                    Case 0 ' cliente
                        CargaGrid CadB & " AND codigoean.codclien=" & DBSet(CodigoActual, "N")
                    Case 1 ' forfait
                        CargaGrid CadB & " AND codigoean.codforfait=" & DBSet(CodigoActual, "T")
                    Case 2 ' variedad
                        CargaGrid CadB & " AND codigoean.codvarie=" & DBSet(CodigoActual, "N")
                End Select
                        
'                lblIndicador.Caption = "RESULTADO BUSQUEDA"
                PonerFocoGrid Me.DataGrid1
            End If
    End Select
End Sub


Private Sub cmdCancelar_Click()
'On Error Resume Next

    Select Case Modo
        Case 1 'BUSQUEDA
            Select Case Tipo
                Case 0 ' cliente
                    If CadB <> "" Then
                        CargaGrid CadB & " AND codclien=" & DBSet(CodigoActual, "N")
                    Else
                        CargaGrid "codclien =" & DBSet(CodigoActual, "N")
                    End If
                Case 1 ' forfait
                    If CadB <> "" Then
                        CargaGrid CadB & " AND codforfait=" & DBSet(CodigoActual, "T")
                    Else
                        CargaGrid "codforfait =" & DBSet(CodigoActual, "T")
                    End If
                Case 2 ' variedad
                    If CadB <> "" Then
                        CargaGrid CadB & " AND codvarie=" & DBSet(CodigoActual, "N")
                    Else
                        CargaGrid "codvarie =" & DBSet(CodigoActual, "N")
                    End If
            End Select
    End Select
    PonerModo 2
    PonerFocoGrid Me.DataGrid1
'    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub cmdRegresar_Click()
Dim Cad As String
Dim i As Integer
Dim J As Integer
Dim Aux As String

    If adodc1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    Cad = ""
    i = 0
    Do
        J = i + 1
        i = InStr(J, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, i - J)
            J = Val(Aux)
            Cad = Cad & adodc1.Recordset.Fields(J) & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
     PonerContRegIndicador
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault

     If PrimeraVez Then
        PrimeraVez = False
        PonerModo 2
'        If Me.CodigoActual <> "" Then
'            Select Case Tipo
'                Case 0 ' me dan el cliente
'                    txtAux(0).Tag = "Código Forfaits|T|N|||codigoean|codforfait||S|"
'                    txtAux(0).MaxLength = 16
'                    txtAux(1).Tag = "Variedad|N|N|0|9999|codigoean|codvarie|0000|S|"
'                    txtAux(1).MaxLength = 4
'
'                Case 1 ' me dan el forfait
'                    txtAux(0).Tag = "Código Cliente|N|N|0|999999|codigoean|codclien|000000|S|"
'                    txtAux(0).MaxLength = 6
'                    txtAux(1).Tag = "Variedad|N|N|0|9999|codigoean|codvarie|0000|S|"
'                    txtAux(1).MaxLength = 4
'
'                Case 2 ' me dan la variedad
'                    txtAux(0).Tag = "Código Cliente|N|N|0|999999|codigoean|codclien|000000|S|"
'                    txtAux(0).MaxLength = 6
'                    txtAux(1).Tag = "Código Forfaits|T|N|||codigoean|codforfait||S|"
'                    txtAux(1).MaxLength = 16
'
'            End Select
'            SituarData Me.adodc1, "codclien = " & DBSet(CodigoActual, "N"), "", True
'        End If
    End If
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    'Icono del formulario
    Me.Icon = frmPpal.Icon

'    With Me.Toolbar1
'        .HotImageList = frmPpal.imgListComun_OM
'        .DisabledImageList = frmPpal.imgListComun_BN
'        .ImageList = frmPpal.imgListComun
'        'el 1 es separadors
'        .Buttons(2).Image = 1   'Buscar
'        .Buttons(3).Image = 2   'Todos
'        'el 4 i el 5 son separadors
'        .Buttons(6).Image = 3   'Insertar
'        .Buttons(7).Image = 4   'Modificar
'        .Buttons(8).Image = 5   'Borrar
'        'el 9 i el 10 son separadors
''        .Buttons(11).Image = 10  'Imprimir
'        .Buttons(11).Image = 11  'Salir
'    End With

    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
        'el 1 es separadors
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2   'Todos
        'el 9 i el 10 son separadors
        .Buttons(8).Image = 10  'imprimir
    End With

    ' La Ayuda
'    With Me.ToolbarAyuda
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 12
'    End With



    '## A mano
'    chkVistaPrevia.Value = CheckValueLeer(Name)
      
'    PonerOpcionesMenu  'En funcion del usuario
    
    
    If Me.CodigoActual <> "" Then
        Select Case Tipo
            Case 0 ' me dan el cliente
                txtAux(0).Tag = "Código Forfaits|T|N|||codigoean|codforfait||S|"
                txtAux(0).MaxLength = 16
                txtAux(1).Tag = "Variedad|N|N|0|9999|codigoean|codvarie|0000|S|"
                txtAux(1).MaxLength = 4
                cmdAceptar.Left = 6365
                cmdCancelar.Left = 7565
                cmdRegresar.Left = 7565
                
            Case 1 ' me dan el forfait
                txtAux(0).Tag = "Código Cliente|N|N|0|999999|codigoean|codclien|000000|S|"
                txtAux(0).MaxLength = 6
                txtAux(1).Tag = "Variedad|N|N|0|9999|codigoean|codvarie|0000|S|"
                txtAux(1).MaxLength = 4
                cmdAceptar.Left = 5365
                cmdCancelar.Left = 6565
                cmdRegresar.Left = 6565
            
            Case 2 ' me dan la variedad
                txtAux(0).Tag = "Código Cliente|N|N|0|999999|codigoean|codclien|000000|S|"
                txtAux(0).MaxLength = 6
                txtAux(1).Tag = "Código Forfaits|T|N|||codigoean|codforfait||S|"
                txtAux(1).MaxLength = 16
                cmdAceptar.Left = 6365
                cmdCancelar.Left = 7565
                cmdRegresar.Left = 7565
            
        End Select
'        SituarData Me.adodc1, "codclien = " & DBSet(CodigoActual, "N"), "", True
    End If
    
    
    '****************** canviar la consulta *********************************+
    Select Case Tipo
        Case 0 ' cliente
            CadenaConsulta = "Select codigoean.codforfait, forfaits.nomconfe, codigoean.codvarie, variedades.nomvarie, codigoean.codigoean, codigoean.refclien "
            CadenaConsulta = CadenaConsulta & " from codigoean, forfaits, variedades where codigoean.codforfait = forfaits.codforfait "
            CadenaConsulta = CadenaConsulta & " and  codigoean.codvarie = variedades.codvarie "
            
            CadB = ""
            CargaGrid "codigoean.codclien =" & DBSet(CodigoActual, "N")
        Case 1 ' forfait
            CadenaConsulta = "Select codigoean.codclien, clientes.nomclien, codigoean.codvarie, variedades.nomvarie, codigoean.codigoean, codigoean.refclien  "
            CadenaConsulta = CadenaConsulta & " from codigoean, clientes, variedades where codigoean.codclien = clientes.codclien "
            CadenaConsulta = CadenaConsulta & " and  codigoean.codvarie = variedades.codvarie "
        
            CadB = ""
            CargaGrid "codigoean.codforfait =" & DBSet(CodigoActual, "T")
        Case 2 ' variedad
            CadenaConsulta = "Select codigoean.codclien, clientes.nomclien, codigoean.codforfait, forfaits.nomconfe, codigoean.codigoean, codigoean.refclien  "
            CadenaConsulta = CadenaConsulta & " from codigoean, clientes, forfaits where codigoean.codclien = clientes.codclien "
            CadenaConsulta = CadenaConsulta & " and  codigoean.codforfait = forfaits.codforfait "
        
            CadB = ""
            CargaGrid "codigoean.codvarie =" & DBSet(CodigoActual, "N")
    End Select
    '************************************************************************
    
'    lblIndicador.Caption = ""
'    PonerModo 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
    If Modo = 4 Then TerminaBloquear
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub


Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 5
                BotonBuscar
        Case 6
                BotonVerTodos
'        Case 6
'                BotonAnyadir
'        Case 7
'                mnModificar_Click
'        Case 8
'                BotonEliminar
'        Case 11 'Imprimir
'            AbrirListado (2)  'OpcionListado=2
        Case 11 'Salir
                mnSalir_Click
    End Select
End Sub


Private Sub CargaGrid(Optional vSQL As String)
    Dim Sql As String, tots As String
    
    adodc1.ConnectionString = conn
    If vSQL <> "" Then
        Sql = CadenaConsulta & " and  " & vSQL
    Else
        Sql = CadenaConsulta
    End If
    '********************* canviar el ORDER BY *********************++
    Select Case Tipo
        Case 0 ' cliente
            Sql = Sql & " ORDER BY codigoean.codforfait, codigoean.codvarie"
        Case 1 ' forfait
            Sql = Sql & " ORDER BY codigoean.codclien, codigoean.codvarie"
        Case 2 ' variedad
            Sql = Sql & " ORDER BY codigoean.codclien, codigoean.codforfait"
    End Select
    '**************************************************************++
    
    adodc1.RecordSource = Sql
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockOptimistic
    adodc1.Refresh
    Set DataGrid1.DataSource = adodc1
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 290
    
    
    Select Case Tipo
        Case 0 ' cliente
            tots = "S|txtAux(0)|T|Forfait|1750|;S|txtAux2(0)|T|Descripción|3100|;"
            tots = tots & "S|txtAux(1)|T|Var.|600|;S|txtAux2(1)|T|Descripción|3100|;"
            tots = tots & "S|txtAux(2)|T|Código EAN|1600|;S|txtAux(3)|T|Ref.Cli.|900|;"
            Me.DataGrid1.Width = 11700
            Me.Width = 12300
            
            
        Case 1 ' forfait
            tots = "S|txtAux(0)|T|Cliente|700|;S|txtAux2(0)|T|Nombre|3400|;"
            tots = tots & "S|txtAux(1)|T|Var.|600|;S|txtAux2(1)|T|Descripción|3100|;"
            tots = tots & "S|txtAux(2)|T|Código EAN|1600|;S|txtAux(3)|T|Ref.Cli.|900|;"
            
            Me.DataGrid1.Width = 11000
            Me.Width = 11600
        
        Case 2 ' variedad
            tots = "S|txtAux(0)|T|Cliente|700|;S|txtAux2(0)|T|Nombre|3400|;"
            tots = tots & "S|txtAux(1)|T|Forfait|1750|;S|txtAux2(1)|T|Descripción|3100|;"
            tots = tots & "S|txtAux(2)|T|Código EAN|1600|;S|txtAux(3)|T|Ref.Cli.|900|;"
            
            Me.DataGrid1.Width = 12200
            Me.Width = 12800

    End Select
    Me.Height = 10350 '8145 '6705
    arregla tots, DataGrid1, Me, 350
   
'   'Habilitamos modificar y eliminar
'   Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
'   Toolbar1.Buttons(8).Enabled = Not adodc1.Recordset.EOF
'   mnModificar.Enabled = Not adodc1.Recordset.EOF
'   mnEliminar.Enabled = Not adodc1.Recordset.EOF
'
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerContRegIndicador()
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    If (Modo = 2 Or Modo = 0) Then
        cadReg = PonerContRegistros(Me.adodc1)
        If CadB = "" Then
            lblIndicador.Caption = cadReg
        Else
            lblIndicador.Caption = "BUSQUEDA: " & cadReg
        End If
    End If
End Sub

'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGrid1_GotFocus()
'  WheelHook DataGrid1
'End Sub
'Private Sub DataGrid1_Lostfocus()
'  WheelUnHook
'End Sub
