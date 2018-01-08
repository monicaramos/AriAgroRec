VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTercAlbPdtes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7095
   Icon            =   "frmTercAlbPdtes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameCobros 
      Height          =   5535
      Left            =   90
      TabIndex        =   11
      Top             =   120
      Width           =   6825
      Begin VB.CheckBox Check2 
         Caption         =   "Agrupado por Variedad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   8
         Top             =   4530
         Width           =   2640
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Incluir Entradas Socios"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   7
         Top             =   4155
         Width           =   2640
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         ItemData        =   "frmTercAlbPdtes.frx":000C
         Left            =   1845
         List            =   "frmTercAlbPdtes.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "Tipo Parcela|N|N|0|1|rcampos|tipoparc||N|"
         Top             =   4410
         Width           =   1845
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Height          =   360
         Index           =   5
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text5"
         Top             =   3585
         Width           =   3810
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Height          =   360
         Index           =   4
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Text5"
         Top             =   3210
         Width           =   3810
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   5
         Top             =   3600
         Width           =   870
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   4
         Top             =   3210
         Width           =   870
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1605
         Width           =   1350
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Código Postal|T|S|||clientes|codposta|||"
         Top             =   1245
         Width           =   1350
      End
      Begin VB.CommandButton cmdCancel 
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
         Left            =   5490
         TabIndex        =   10
         Top             =   4905
         Width           =   1065
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
         Left            =   4305
         TabIndex        =   9
         Top             =   4905
         Width           =   1065
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   2
         Top             =   2190
         Width           =   870
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   3
         Top             =   2565
         Width           =   870
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Height          =   360
         Index           =   0
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text5"
         Top             =   2190
         Width           =   3810
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Height          =   360
         Index           =   1
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   2565
         Width           =   3810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Informe"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   240
         Index           =   36
         Left            =   510
         TabIndex        =   26
         Top             =   4140
         Width           =   1560
      End
      Begin VB.Label Label1 
         Caption         =   "Albaranes de Terceros"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   495
         TabIndex        =   25
         Top             =   390
         Width           =   4680
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1530
         MouseIcon       =   "frmTercAlbPdtes.frx":0010
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3585
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1545
         MouseIcon       =   "frmTercAlbPdtes.frx":0162
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3210
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   240
         Index           =   2
         Left            =   510
         TabIndex        =   24
         Top             =   2970
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   855
         TabIndex        =   23
         Top             =   3585
         Width           =   600
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   855
         TabIndex        =   22
         Top             =   3210
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Albarán"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   16
         Left            =   510
         TabIndex        =   19
         Top             =   945
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   870
         TabIndex        =   18
         Top             =   1245
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   870
         TabIndex        =   17
         Top             =   1605
         Width           =   600
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1530
         Picture         =   "frmTercAlbPdtes.frx":02B4
         ToolTipText     =   "Buscar fecha"
         Top             =   1245
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1530
         Picture         =   "frmTercAlbPdtes.frx":033F
         ToolTipText     =   "Buscar fecha"
         Top             =   1605
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   870
         TabIndex        =   16
         Top             =   2190
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   870
         TabIndex        =   15
         Top             =   2565
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   240
         Index           =   11
         Left            =   510
         TabIndex        =   14
         Top             =   1950
         Width           =   540
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1560
         MouseIcon       =   "frmTercAlbPdtes.frx":03CA
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   2190
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1545
         MouseIcon       =   "frmTercAlbPdtes.frx":051C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   2565
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmTercAlbPdtes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor MANOLO +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmVar As frmComVar 'Variedad
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'Socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report

Dim PrimeraVez As Boolean

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me

End Sub


Private Sub cmdAceptar_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim I As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim Tipo As Byte
Dim cTabla As String


InicializarVbles
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    CadParam = CadParam & "|pUsu=" & vUsu.Codigo & "|"
    numParam = numParam + 1
    
    
     '======== FORMULA  ====================================
    'Seleccionar registros de la empresa conectada
'    Codigo = "{" & tabla & ".codempre}=" & vEmpresa.codEmpre
'    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
'    If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
    
    
    'D/H Socio Tercero
    cDesde = Trim(txtCodigo(0).Text)
    cHasta = Trim(txtCodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rhisfruta.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    'D/H Variedad
    cDesde = Trim(txtCodigo(4).Text)
    cHasta = Trim(txtCodigo(5).Text)
    nDesde = txtNombre(4).Text
    nHasta = txtNombre(5).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rhisfruta.codvarie}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad= """) Then Exit Sub
    End If
    
    'D/H Fecha albaran
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rhisfruta.fecalbar}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, "", "", "pDHFecha= """) Then Exit Sub
    End If
    
    ' se queda todo como estaba
    If vParamAplic.Cooperativa <> 9 Then ' si es ddistinto de natural
        Select Case Combo1(0).ListIndex
            Case 0
                ' Tipo = 0: albaranes pendientes de facturar
                cadFormula = "{tmpinformes.codusu} = {@pUsu}"
    
                If ProcesarCambios(cadSelect) Then
                    cadTitulo = "Albaranes Pendientes de Facturar "
                    If Check1.Value = 0 Then cadTitulo = cadTitulo & "Terceros"
                    cadNombreRPT = "rAlbPdtesTerc.rpt"
    
                    LlamarImprimir
                End If
    
            Case 1
                ' Tipo = 1: albaranes facturados
                If Not AnyadirAFormula(cadFormula, "{rsocios.tipoprod} = 1") Then Exit Sub
                If Not AnyadirAFormula(cadSelect, "{rsocios.tipoprod} = 1") Then Exit Sub
    
                cTabla = tabla & " INNER JOIN rsocios On rhisfruta.codsocio = rsocios.codsocio"
                If HayRegistros(cTabla, cadSelect) Then
                    cadTitulo = "Albaranes Facturados "
                    If Check1.Value = 0 Then cadTitulo = cadTitulo & "Terceros"
                    cadNombreRPT = "rAlbFactuTerc.rpt"
                    LlamarImprimir
                End If
    
            Case 2
                ' todos (facturados y no facturados)
                If Not AnyadirAFormula(cadFormula, "{rsocios.tipoprod} = 1") Then Exit Sub
                If Not AnyadirAFormula(cadSelect, "{rsocios.tipoprod} = 1") Then Exit Sub
    
                cTabla = tabla & " INNER JOIN rsocios On rhisfruta.codsocio = rsocios.codsocio"
    
                If HayRegistros(cTabla, cadSelect) Then
                      cadTitulo = "Albaranes Terceros"
                      If Check1.Value = 1 Then cadTitulo = cadTitulo & " y no Terceros"
    
                      cadNombreRPT = "rAlbaranesTerc.rpt"
                      LlamarImprimir
                End If
    
        End Select
    Else
        ' caso de natural
        cadFormula = "{tmpinformes.codusu} = {@pUsu}"
        
        If Check2.Value = 1 Then
            CadParam = CadParam & "pGroup={tmpinformes.importe2}|"
        Else
            CadParam = CadParam & "pGroup={tmpinformes.importe1}|"
        End If
        numParam = numParam + 1
        
    
        If ProcesarCambiosNew(cadSelect) Then
            Select Case Combo1(0).ListIndex
                Case 0
                    cadTitulo = "Albaranes Pendientes de Facturar "
                    If Check1.Value = 0 Then cadTitulo = cadTitulo & "Terceros"
                
                Case 1
                    cadTitulo = "Albaranes Facturados "
                    If Check1.Value = 0 Then cadTitulo = cadTitulo & "Terceros"
                
                
                Case 2
                    cadTitulo = "Albaranes Terceros"
                    If Check1.Value = 1 Then cadTitulo = cadTitulo & " y no Terceros"
                
            End Select
            
            cadNombreRPT = "rAlbaranesTercSoc.rpt"
        
            LlamarImprimir
        End If
    End If


End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(2)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
     For H = 0 To 1
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Next H
     For H = 4 To 5
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
     Next H

    '###Descomentar
'    CommitConexion
         
    FrameCobrosVisible True, H, W
    indFrame = 5
    tabla = "rhisfruta"
    
    CargaCombo
    
    Check1.visible = (vParamAplic.Cooperativa = 9)
    Check1.Enabled = (vParamAplic.Cooperativa = 9)
    
    Check2.visible = (vParamAplic.Cooperativa = 9)
    Check2.Enabled = (vParamAplic.Cooperativa = 9)
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdcancel.Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    txtCodigo(CByte(imgFec(2).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de socios
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Variedades
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgFec_Click(Index As Integer)
'FEchas
    Dim esq, dalt As Long
    Dim obj As Object
    
    Set frmC = New frmCal

    esq = imgFec(Index).Left
    dalt = imgFec(Index).Top

    Set obj = imgFec(Index).Container

    While imgFec(Index).Parent.Name <> obj.Name
        esq = esq + obj.Left
        dalt = dalt + obj.Top
        Set obj = obj.Container
    Wend
       
    ' es desplega dalt i cap a la esquerra
    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + 420 + 30

    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
    imgFec(2).Tag = Index 'independentment de les dates que tinga, sempre pose l'index en la 27
    If txtCodigo(Index).Text <> "" Then frmC.NovaData = txtCodigo(Index).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco txtCodigo(CByte(imgFec(2).Tag))
    ' ***************************
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 'Socios Terceros
            AbrirFrmTerceros (Index)
        
        Case 4, 5 'VARIEDADES
            AbrirFrmVariedades (Index)
        
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub Optcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub OptNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        PonerFocoBtn Me.cmdAceptar
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'14/02/2007 antes
'    KEYpress KeyAscii
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'socio desde
            Case 1: KEYBusqueda KeyAscii, 1 'socio hasta
            Case 4: KEYBusqueda KeyAscii, 4 'variedad desde
            Case 5: KEYBusqueda KeyAscii, 5 'variedad hasta
            Case 2: KEYFecha KeyAscii, 2 'fecha desde
            Case 3: KEYFecha KeyAscii, 3 'fecha hasta
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFec_Click (Indice)
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    
    Select Case Index
        Case 0, 1 'socio tercero
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
        Case 2, 3 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 4, 5 'VARIEDAD
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "00")
    End Select
End Sub

Private Sub FrameCobrosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCobros.visible = visible
    If visible = True Then
        Me.FrameCobros.Top = -90
        Me.FrameCobros.Left = 0
        Me.FrameCobros.Height = 5790
        Me.FrameCobros.Width = 6930
        W = Me.FrameCobros.Width
        H = Me.FrameCobros.Height
    End If
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    CadParam = ""
    numParam = 0
End Sub

Private Function PonerDesdeHasta(codD As String, codH As String, nomD As String, nomH As String, param As String) As Boolean
'IN: codD,codH --> codigo Desde/Hasta
'    nomD,nomH --> Descripcion Desde/Hasta
'Añade a cadFormula y cadSelect la cadena de seleccion:
'       "(codigo>=codD AND codigo<=codH)"
' y añade a cadParam la cadena para mostrar en la cabecera informe:
'       "codigo: Desde codD-nomd Hasta: codH-nomH"
Dim devuelve As String
Dim devuelve2 As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(codD, codH, Codigo, TipCod)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    If TipCod <> "F" Then 'Fecha
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        devuelve2 = CadenaDesdeHastaBD(codD, codH, Codigo, TipCod)
        If devuelve2 = "Error" Then Exit Function
        If Not AnyadirAFormula(cadSelect, devuelve2) Then Exit Function
    End If
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            CadParam = CadParam & AnyadirParametroDH(param, codD, codH, nomD, nomH)
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function

Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = CadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .ConSubInforme = True
        .EnvioEMail = False
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmTerceros(Indice As Integer)
    indCodigo = Indice
    Set frmSoc = New frmManSocios
    frmSoc.DatosADevolverBusqueda = "0|1|"
    frmSoc.Show vbModal
    Set frmSoc = Nothing
End Sub

Private Sub AbrirFrmVariedades(Indice As Integer)
    indCodigo = Indice
    Set frmVar = New frmComVar
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.CodigoActual = txtCodigo(indCodigo)
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub

Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
        '.SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
        .OtrosParametros = CadParam
        .NumeroParametros = numParam
        '##descomen
'        .MostrarTree = MostrarTree
'        .Informe = MIPATH & Nombre
'        .InfConta = InfConta
        '##
        
'        If NombreSubRptConta <> "" Then
'            .SubInformeConta = NombreSubRptConta
'        Else
'            .SubInformeConta = ""
'        End If
        '##descomen
'        .ConSubInforme = ConSubInforme
        '##
        .Opcion = ""
'        .ExportarPDF = (chkEMAIL.Value = 1)
        .Show vbModal
    End With
    
'    If Me.chkEMAIL.Value = 1 Then
'    '####Descomentar
'        If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
'    End If
    Unload Me
End Sub

Private Sub AbrirEMail()
    If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
End Sub


Private Function HayRegistros(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
Dim Rs As ADODB.Recordset

    SQL = "Select * FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Rs.EOF Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
        HayRegistros = False
    Else
        HayRegistros = True
    End If

End Function

Private Function ProcesarCambios(cadWHERE As String) As Boolean
Dim SQL As String
Dim Sql1 As String
Dim I As Integer
Dim HayReg As Integer
Dim B As Boolean

On Error GoTo eProcesarCambios

    HayReg = 0
    
    conn.Execute "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
        
    If cadWHERE <> "" Then
        cadWHERE = QuitarCaracterACadena(cadWHERE, "{")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "}")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "_1")
    End If
        
    SQL = "insert into tmpinformes (codusu, codigo1) "
    
    SQL = SQL & " select " & DBSet(vUsu.Codigo, "N") & ", rhisfruta.numalbar "
    SQL = SQL & " from rhisfruta inner join rsocios on rhisfruta.codsocio = rsocios.codsocio "
    SQL = SQL & " where not numalbar in (select numalbar from rlifter) "
    SQL = SQL & "  and rsocios.tipoprod = 1 "
    If cadWHERE <> "" Then SQL = SQL & " and " & cadWHERE
    
    conn.Execute SQL
        
    ProcesarCambios = HayRegistros("tmpinformes", "codusu = " & vUsu.Codigo)

eProcesarCambios:
    If Err.Number <> 0 Then
        ProcesarCambios = False
    End If
End Function



Private Function ProcesarCambiosNew(cadWHERE As String) As Boolean
Dim SQL As String
Dim Sql1 As String
Dim I As Integer
Dim HayReg As Integer
Dim B As Boolean

On Error GoTo eProcesarCambios

    HayReg = 0
    
    conn.Execute "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
        
    If cadWHERE <> "" Then
        cadWHERE = QuitarCaracterACadena(cadWHERE, "{")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "}")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "_1")
    End If
        
    SQL = "insert into tmpinformes (codusu, codigo1, fecha1, importe1, importe2, importe3, importe4, precio1, importe5, nombre1) "
    
    Select Case Combo1(0).ListIndex
        Case 0 ' no facturados
            SQL = SQL & " select " & DBSet(vUsu.Codigo, "N") & ", rhisfruta.numalbar, rhisfruta.fecalbar, "
            SQL = SQL & " rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.numcajon, rhisfruta.kilosnet, "
            SQL = SQL & " rhisfruta.prestimado,round(if(rhisfruta.prestimado is null, 0,rhisfruta.prestimado) * kilosnet,2) importe , null "
            SQL = SQL & " from rhisfruta inner join rsocios on rhisfruta.codsocio = rsocios.codsocio "
            SQL = SQL & " where not (numalbar,fecalbar,codvarie) in (select numalbar,fechaalb,codvarie from rlifter) "
            SQL = SQL & "  and rsocios.tipoprod = 1 "
            If cadWHERE <> "" Then SQL = SQL & " and " & cadWHERE
            
        
            If Check1.Value = 1 Then
                SQL = SQL & " union "
                SQL = SQL & " select " & DBSet(vUsu.Codigo, "N") & ", rhisfruta.numalbar, rhisfruta.fecalbar, "
                SQL = SQL & " rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.numcajon, rhisfruta.kilosnet, "
                SQL = SQL & " rhisfruta.prestimado,round(if(rhisfruta.prestimado is null, 0,rhisfruta.prestimado) * kilosnet,2) importe, null "
                SQL = SQL & " from rhisfruta inner join rsocios on rhisfruta.codsocio = rsocios.codsocio "
                SQL = SQL & " where not (numalbar,fecalbar,codvarie) in (select numalbar,fecalbar,codvarie from rfactsoc_albaran where codtipom in (select codtipom from usuarios.stipom where tipodocu in (1,2))) "
                SQL = SQL & " and rsocios.tipoprod <> 1 "
                If cadWHERE <> "" Then SQL = SQL & " and " & cadWHERE
            End If
        
        
        Case 1 ' facturados
            SQL = SQL & " select " & DBSet(vUsu.Codigo, "N") & ", rhisfruta.numalbar, rhisfruta.fecalbar, "
            SQL = SQL & " rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.numcajon, rhisfruta.kilosnet, "
            SQL = SQL & " round(rlifter.importel / rhisfruta.kilosnet,4) ,rlifter.importel importe, rlifter.numfactu "
            SQL = SQL & " from (rhisfruta inner join rsocios on rhisfruta.codsocio = rsocios.codsocio) "
            SQL = SQL & " inner join rlifter on rlifter.numalbar = rhisfruta.numalbar and rlifter.fechaalb = rhisfruta.fecalbar and rlifter.codvarie = rhisfruta.codvarie "
            SQL = SQL & " where rsocios.tipoprod = 1 "
            If cadWHERE <> "" Then SQL = SQL & " and " & cadWHERE
            
            If Check1.Value = 1 Then
                SQL = SQL & " union "
                SQL = SQL & " select " & DBSet(vUsu.Codigo, "N") & ", rhisfruta.numalbar, rhisfruta.fecalbar, "
                SQL = SQL & " rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.numcajon, rhisfruta.kilosnet, "
                SQL = SQL & " round((sum(rfactsoc_albaran.importe) - sum(rfactsoc_albaran.imporgasto))/ rhisfruta.kilosnet,4), sum(rfactsoc_albaran.importe) - sum(rfactsoc_albaran.imporgasto)  importe, rfactsoc_albaran.numfactu    "
                SQL = SQL & " from ((rhisfruta inner join rsocios on rhisfruta.codsocio = rsocios.codsocio) "
                SQL = SQL & " inner join rfactsoc_albaran on rfactsoc_albaran.numalbar = rhisfruta.numalbar and rfactsoc_albaran.fecalbar = rhisfruta.fecalbar and rfactsoc_albaran.codvarie = rhisfruta.codvarie) "
                SQL = SQL & " inner join rfactsoc on rfactsoc_albaran.codtipom = rfactsoc.codtipom and rfactsoc_albaran.numfactu = rfactsoc.numfactu and rfactsoc_albaran.fecfactu = rfactsoc.fecfactu "
                '[Monica]06/11/2013. Cambiado por lo de abajo
                'sql = sql & " where rfactsoc_albaran.codtipom in (select codtipom from usuarios.stipom where tipodocu = 2) and "
                SQL = SQL & " where (rfactsoc_albaran.codtipom in (select codtipom from usuarios.stipom where tipodocu = 2) or  "
                SQL = SQL & " (rfactsoc_albaran.codtipom in (select codtipom from usuarios.stipom where tipodocu = 1) and not rfactsoc_albaran.numalbar in (select numalbar from rfactsoc_albaran where codtipom in (select codtipom from usuarios.stipom where tipodocu = 2)))) and "
                SQL = SQL & " not (rfactsoc_albaran.codtipom, rfactsoc_albaran.numfactu, rfactsoc_albaran.fecfactu) in (select rectif_codtipom,rectif_numfactu,rectif_fecfactu from rfactsoc where not  rectif_codtipom is null and not rectif_numfactu is null and not rectif_fecfactu is null) and "
                SQL = SQL & " rsocios.tipoprod <> 1 "
                If cadWHERE <> "" Then SQL = SQL & " and " & cadWHERE
                
                SQL = SQL & " group by 1,2,3,4,5,6,7,10 "
                
            End If
            
        Case 2 ' todos
            ' no facturados
            SQL = SQL & " select " & DBSet(vUsu.Codigo, "N") & ", rhisfruta.numalbar, rhisfruta.fecalbar, "
            SQL = SQL & " rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.numcajon, rhisfruta.kilosnet, "
            SQL = SQL & " rhisfruta.prestimado,round(if(rhisfruta.prestimado is null, 0,rhisfruta.prestimado) * kilosnet,2) importe , null "
            SQL = SQL & " from rhisfruta inner join rsocios on rhisfruta.codsocio = rsocios.codsocio "
            SQL = SQL & " where not (numalbar,fecalbar,codvarie) in (select numalbar,fechaalb,codvarie from rlifter) "
            SQL = SQL & "  and rsocios.tipoprod = 1 "
            If cadWHERE <> "" Then SQL = SQL & " and " & cadWHERE
            
        
            If Check1.Value = 1 Then
                SQL = SQL & " union "
                SQL = SQL & " select " & DBSet(vUsu.Codigo, "N") & ", rhisfruta.numalbar, rhisfruta.fecalbar, "
                SQL = SQL & " rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.numcajon, rhisfruta.kilosnet, "
                SQL = SQL & " rhisfruta.prestimado,round(if(rhisfruta.prestimado is null, 0,rhisfruta.prestimado) * kilosnet,2) importe, null "
                SQL = SQL & " from rhisfruta inner join rsocios on rhisfruta.codsocio = rsocios.codsocio "
                SQL = SQL & " where not (numalbar,fecalbar,codvarie) in (select numalbar,fecalbar,codvarie from rfactsoc_albaran where codtipom in (select codtipom from usuarios.stipom where tipodocu in (1,2))) "
                SQL = SQL & " and rsocios.tipoprod <> 1 "
                If cadWHERE <> "" Then SQL = SQL & " and " & cadWHERE
            End If
        
            'facturados
            SQL = SQL & " union "
            SQL = SQL & " select " & DBSet(vUsu.Codigo, "N") & ", rhisfruta.numalbar, rhisfruta.fecalbar, "
            SQL = SQL & " rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.numcajon, rhisfruta.kilosnet, "
            SQL = SQL & " round(rlifter.importel / rhisfruta.kilosnet,4) ,rlifter.importel importe, rlifter.numfactu "
            SQL = SQL & " from (rhisfruta inner join rsocios on rhisfruta.codsocio = rsocios.codsocio) "
            SQL = SQL & " inner join rlifter on rlifter.numalbar = rhisfruta.numalbar and rlifter.fechaalb = rhisfruta.fecalbar and rlifter.codvarie = rhisfruta.codvarie "
            SQL = SQL & " where rsocios.tipoprod = 1 "
            If cadWHERE <> "" Then SQL = SQL & " and " & cadWHERE
            
            If Check1.Value = 1 Then
                SQL = SQL & " union "
                SQL = SQL & " select " & DBSet(vUsu.Codigo, "N") & ", rhisfruta.numalbar, rhisfruta.fecalbar, "
                SQL = SQL & " rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.numcajon, rhisfruta.kilosnet, "
                SQL = SQL & " round((sum(rfactsoc_albaran.importe) - sum(rfactsoc_albaran.imporgasto))/ rhisfruta.kilosnet,4), sum(rfactsoc_albaran.importe) - sum(rfactsoc_albaran.imporgasto)  importe, rfactsoc_albaran.numfactu    "
                SQL = SQL & " from ((rhisfruta inner join rsocios on rhisfruta.codsocio = rsocios.codsocio) "
                SQL = SQL & " inner join rfactsoc_albaran on rfactsoc_albaran.numalbar = rhisfruta.numalbar  and rfactsoc_albaran.fecalbar = rhisfruta.fecalbar and rfactsoc_albaran.codvarie = rhisfruta.codvarie) "
                SQL = SQL & " inner join rfactsoc on rfactsoc_albaran.codtipom = rfactsoc.codtipom and rfactsoc_albaran.numfactu = rfactsoc.numfactu and rfactsoc_albaran.fecfactu = rfactsoc.fecfactu "
                '[Monica]06/11/2013. Cambiado por lo de abajo
                'sql = sql & " where rfactsoc_albaran.codtipom in (select codtipom from usuarios.stipom where tipodocu = 2) and "
                SQL = SQL & " where (rfactsoc_albaran.codtipom in (select codtipom from usuarios.stipom where tipodocu = 2) or  "
                SQL = SQL & " (rfactsoc_albaran.codtipom in (select codtipom from usuarios.stipom where tipodocu = 1) and not rfactsoc_albaran.numalbar in (select numalbar from rfactsoc_albaran where codtipom in (select codtipom from usuarios.stipom where tipodocu = 2)))) and "
                'hasta aqui
                SQL = SQL & " not (rfactsoc_albaran.codtipom, rfactsoc_albaran.numfactu, rfactsoc_albaran.fecfactu) in (select rectif_codtipom,rectif_numfactu,rectif_fecfactu from rfactsoc where not  rectif_codtipom is null and not rectif_numfactu is null and not rectif_fecfactu is null) and "
                SQL = SQL & " rsocios.tipoprod <> 1 "
                If cadWHERE <> "" Then SQL = SQL & " and " & cadWHERE
                
                SQL = SQL & " group by 1,2,3,4,5,6,7,10 "
                
            End If
            
    End Select
    
    conn.Execute SQL
        
    ProcesarCambiosNew = HayRegistros("tmpinformes", "codusu = " & vUsu.Codigo)

eProcesarCambios:
    If Err.Number <> 0 Then
        ProcesarCambiosNew = False
    End If
End Function

Private Sub InsertaLineaEnTemporal(ByRef ItmX As ListItem)
Dim SQL As String
Dim Codmacta As String
Dim Rs As ADODB.Recordset
Dim Sql1 As String

        Sql1 = "insert into tmpinformes(codusu, codigo1) values ("
        Sql1 = Sql1 & DBSet(vUsu.Codigo, "N") & "," & DBSet(ItmX.Text, "N") & ")"

        conn.Execute Sql1
    
End Sub


' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim I As Integer
Dim Rs As ADODB.Recordset
Dim SQL As String


    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For I = 0 To Combo1.Count - 1
        Combo1(I).Clear
    Next I
    
    'tipo de hectareas
    Combo1(0).AddItem "Pendientes Facturar"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Facturados"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Todos"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    
End Sub
