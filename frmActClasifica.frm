VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmActClasifica 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualizar Entradas Clasificadas"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6750
   Icon            =   "frmActClasifica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameEntradasCampo 
      Height          =   6300
      Left            =   45
      TabIndex        =   10
      Top             =   45
      Width           =   6615
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
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1035
         Width           =   1005
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
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   0
         Top             =   630
         Width           =   1005
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "Text5"
         Top             =   1035
         Width           =   3645
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
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "Text5"
         Top             =   630
         Width           =   3645
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text5"
         Top             =   2130
         Width           =   3645
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text5"
         Top             =   1725
         Width           =   3645
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
         Left            =   1755
         MaxLength       =   3
         TabIndex        =   3
         Top             =   2130
         Width           =   1005
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
         Left            =   1755
         MaxLength       =   3
         TabIndex        =   2
         Top             =   1725
         Width           =   1005
      End
      Begin VB.CommandButton Command4 
         Height          =   440
         Left            =   7860
         Picture         =   "frmActClasifica.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command3 
         Height          =   440
         Left            =   7860
         Picture         =   "frmActClasifica.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
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
         Index           =   3
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   3165
         Width           =   3645
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
         Index           =   2
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   2760
         Width           =   3645
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
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   5
         Top             =   3165
         Width           =   1005
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
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   4
         Top             =   2760
         Width           =   1005
      End
      Begin VB.CommandButton CmdAceptar 
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
         Index           =   0
         Left            =   4140
         TabIndex        =   8
         Top             =   5625
         Width           =   1035
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
         Index           =   0
         Left            =   5325
         TabIndex        =   9
         Top             =   5625
         Width           =   1035
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
         Index           =   7
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   7
         Top             =   4140
         Width           =   1320
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
         Index           =   6
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   6
         Top             =   3735
         Width           =   1320
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   405
         TabIndex        =   26
         Top             =   4680
         Width           =   6030
         _ExtentX        =   10636
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
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
         Index           =   2
         Left            =   495
         TabIndex        =   32
         Top             =   450
         Width           =   540
      End
      Begin VB.Label Label2 
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
         Index           =   1
         Left            =   690
         TabIndex        =   31
         Top             =   690
         Width           =   690
      End
      Begin VB.Label Label2 
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
         Index           =   0
         Left            =   690
         TabIndex        =   30
         Top             =   1125
         Width           =   645
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1440
         MouseIcon       =   "frmActClasifica.frx":0620
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1440
         MouseIcon       =   "frmActClasifica.frx":0772
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   675
         Width           =   240
      End
      Begin VB.Label lblProgres 
         Caption         =   "Label1"
         Height          =   240
         Left            =   405
         TabIndex        =   27
         Top             =   5085
         Width           =   6000
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1440
         MouseIcon       =   "frmActClasifica.frx":08C4
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2160
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1440
         MouseIcon       =   "frmActClasifica.frx":0A16
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   1725
         Width           =   240
      End
      Begin VB.Label Label2 
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
         Index           =   28
         Left            =   690
         TabIndex        =   25
         Top             =   2160
         Width           =   645
      End
      Begin VB.Label Label2 
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
         Index           =   18
         Left            =   690
         TabIndex        =   24
         Top             =   1725
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clase"
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
         Left            =   495
         TabIndex        =   23
         Top             =   1485
         Width           =   525
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1440
         Picture         =   "frmActClasifica.frx":0B68
         ToolTipText     =   "Buscar fecha"
         Top             =   3735
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1440
         Picture         =   "frmActClasifica.frx":0BF3
         ToolTipText     =   "Buscar fecha"
         Top             =   4140
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1440
         MouseIcon       =   "frmActClasifica.frx":0C7E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3195
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1440
         MouseIcon       =   "frmActClasifica.frx":0DD0
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label2 
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
         Index           =   26
         Left            =   690
         TabIndex        =   20
         Top             =   3195
         Width           =   645
      End
      Begin VB.Label Label2 
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
         Index           =   25
         Left            =   690
         TabIndex        =   19
         Top             =   2805
         Width           =   690
      End
      Begin VB.Label Label2 
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
         Index           =   24
         Left            =   495
         TabIndex        =   18
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label2 
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
         Index           =   21
         Left            =   690
         TabIndex        =   17
         Top             =   4140
         Width           =   645
      End
      Begin VB.Label Label2 
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
         Index           =   20
         Left            =   690
         TabIndex        =   16
         Top             =   3795
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Index           =   19
         Left            =   495
         TabIndex        =   15
         Top             =   3510
         Width           =   600
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6075
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmActClasifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: LAURA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar n� oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(Cliente As String, Observaciones As String)

Private Conexion As Byte
'1.- Conexi�n a BD Ariges  2.- Conexi�n a BD Conta

Private HaDevueltoDatos As Boolean


Private WithEvents frmSec As frmManSeccion 'Secciones
Attribute frmSec.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'Socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmVar As frmComVar 'Variedades de comercial
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCla As frmComercial 'Ayuda de Clases de comercial
Attribute frmCla.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'Mensajes
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmMens1 As frmMensajes 'Mensajes de entradas duplicadas
Attribute frmMens1.VB_VarHelpID = -1


'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe


Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'n� de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'C�digo para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String
Dim CodTipoMov As String
Dim Continuar As Boolean

Dim PrimeraVez As Boolean
Dim Contabilizada As Byte

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me

End Sub

Private Sub cmdAceptar_Click(Index As Integer)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim SQL As String
Dim HayReg As Boolean
Dim cTabla As String

    InicializarVbles
    
    'A�adir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    Select Case Index
        Case 0 ' informe de entradas de bascula
            '======== FORMULA  ====================================
            'D/H SOCIO
            cDesde = Trim(txtCodigo(4).Text)
            cHasta = Trim(txtCodigo(5).Text)
            nDesde = txtNombre(4).Text
            nHasta = txtNombre(5).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rclasifica.codsocio}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
            End If
            
            'D/H CLASE
            cDesde = Trim(txtCodigo(0).Text)
            cHasta = Trim(txtCodigo(1).Text)
            nDesde = txtNombre(0).Text
            nHasta = txtNombre(1).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{variedades.codclase}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
            End If
            
            'D/H VARIEDAD
            cDesde = Trim(txtCodigo(2).Text)
            cHasta = Trim(txtCodigo(3).Text)
            nDesde = txtNombre(2).Text
            nHasta = txtNombre(3).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rclasifica.codvarie}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad=""") Then Exit Sub
            End If

            'D/H fecha
            cDesde = Trim(txtCodigo(6).Text)
            cHasta = Trim(txtCodigo(7).Text)
            nDesde = ""
            nHasta = ""
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{rclasifica.fechaent}"
                TipCod = "F"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
            End If
            
            tabla = "(rclasifica INNER JOIN variedades ON rclasifica.codvarie = variedades.codvarie) "

            'Comprobar si hay registros a Mostrar antes de abrir el Informe
            If HayRegParaInforme(tabla, cadSelect) Then
                
                SQL = "delete from tmpclasifica where codusu = " & vUsu.Codigo
                conn.Execute SQL
            
                ' comprobamos que no existan las notas a actualizar en el hco de entradas
                If vParamAplic.SeRespetaNota Then
                    HayReg = HayRegEnHcoEntradas(tabla, cadSelect)
                    If HayReg Then
                        Set frmMens1 = New frmMensajes
                        frmMens1.OpcionMensaje = 19
                        frmMens1.Label1(3).Caption = "Entradas duplicadas en el Hist�rico"
                        frmMens1.Label1(2).visible = False
                        frmMens1.CmdAceptarPal.visible = False
                        frmMens1.CmdAceptarPal.Enabled = False
                        frmMens1.CmdCanPal.Caption = "&Salir"
                        frmMens1.Show vbModal
                        Set frmMens1 = Nothing
                    
                        MsgBox "No se ha podido realizar el proceso. Revise.", vbExclamation
                        cmdCancel_Click (0)
                        
                        Exit Sub
                    End If
                End If
            
                HayReg = HayRegSinClasificacion(tabla, cadSelect)
                If HayReg Then
'[Monica]:04/06/2010 antes no dejabamos seguir si habian registros sin clasificacion
'                    ahora preguntamos si quieren seguir actualizando solo los clasificados
'                    MsgBox "Hay registros sin clasificaci�n. Revise.", vbExclamation
                    If MsgBox("Hay registros sin clasificaci�n." & vbCrLf & " � Desea continuar con la actualizaci�n de registros clasificados ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                        Exit Sub
                    Else
                        HayReg = False
                    End If
                End If
'                HayReg = HayRegSinClasificacion(Tabla, cadSelect)
                
                '[Monica] 06/05/2010: si hay registros sin gastos correctos (acarreo, recoleccion)
                '                     a�adida la condicion de que no lo compruebe si es alzira
                If vParamAplic.Cooperativa <> 4 Then
                    HayReg = HayRegSinGastosCorrectos(tabla, cadSelect)
                End If
                    
                If HayReg Then
                    Dim cad As String
                    
                    Set frmMens = New frmMensajes
'                    frmMens.cadWHERE = cadSelect
                    frmMens.OpcionMensaje = 19
                    frmMens.Show vbModal
                    Set frmMens = Nothing
                
                    If Continuar Then
                        If ActualizarTabla(tabla, cadSelect) Then
                            MsgBox "Proceso realizado correctamente.", vbExclamation
                            cmdCancel_Click (0)
                        End If
                    End If
                Else
                    If ActualizarTabla(tabla, cadSelect) Then
                        MsgBox "Proceso realizado correctamente.", vbExclamation
                        cmdCancel_Click (0)
                    End If
                End If
            End If
    End Select
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(4)
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

    For H = 0 To imgBuscar.Count - 1
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    
    Me.Pb1.visible = False
    Me.lblProgres.visible = False
    
    tabla = "rclasifica"
    CodTipoMov = "ALF"
    
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtCodigo(CByte(imgFec(0).Tag) + 6).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmCla_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") ' codigo de clase
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) ' descripcion
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim SQL As String
    
    Continuar = False
    If CadenaSeleccion <> "0" Then
'        sql = "not rclasifica.numnotac in (" & CadenaSeleccion & ")"
'        If Not AnyadirAFormula(cadSelect, sql) Then Exit Sub
        Continuar = True
    End If
    
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 ' Clase
            AbrirFrmClase (Index)
        
        Case 2, 3 'VARIEDADES
            AbrirFrmVariedad (Index)
        
        Case 4, 5 'SOCIOS
            AbrirFrmSocio (Index)
    
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub imgFec_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
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

    menu = Me.Height - Me.ScaleHeight 'ac� tinc el heigth del men� i de la toolbar

    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40

    imgFec(0).Tag = Index '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtCodigo(Index + 6).Text <> "" Then frmC.NovaData = txtCodigo(Index + 6).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtCodigo(CByte(imgFec(0).Tag) + 6) '<===
    ' ********************************************

End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'15/02/2007
'    KEYpress KeyAscii
'ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'clase desde
            Case 1: KEYBusqueda KeyAscii, 1 'clase hasta
            Case 2: KEYBusqueda KeyAscii, 2 'variedad desde
            Case 3: KEYBusqueda KeyAscii, 3 'variedad hasta
            Case 4: KEYBusqueda KeyAscii, 4 'socio desde
            Case 5: KEYBusqueda KeyAscii, 5 'socio hasta
            Case 6: KEYFecha KeyAscii, 0 'fecha desde
            Case 7: KEYFecha KeyAscii, 1 'fecha hasta
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFec_Click (Indice)
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1 'CLASES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "clases", "nomclase", "codclase", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
    
        Case 4, 5 'SOCIOS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
    
        Case 6, 7 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 2, 3 'VARIEDADES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
    End Select
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
'A�ade a cadFormula y cadSelect la cadena de seleccion:
'       "(codigo>=codD AND codigo<=codH)"
' y a�ade a cadParam la cadena para mostrar en la cabecera informe:
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
        .EnvioEMail = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmClase(Indice As Integer)
    indCodigo = Indice
    Set frmCla = New frmComercial
    
    AyudaClasesCom frmCla, txtCodigo(Indice).Text
    
    Set frmCla = Nothing
End Sub

Private Sub AbrirFrmVariedad(Indice As Integer)
    indCodigo = Indice
    Set frmVar = New frmComVar
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub

Private Sub AbrirFrmSocio(Indice As Integer)
    indCodigo = Indice
    Set frmSoc = New frmManSocios
    frmSoc.DatosADevolverBusqueda = "0|1|"
    frmSoc.Show vbModal
    Set frmSoc = Nothing
End Sub

Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
'        .SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
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
        .Opcion = 0
'        .ExportarPDF = (chkEMAIL.Value = 1)
        .Show vbModal
    End With
    
'    If Me.chkEMAIL.Value = 1 Then
'    '####Descomentar
'        If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
'    End If
    Screen.MousePointer = vbDefault
    
    Unload Me
End Sub

Private Sub AbrirEMail()
    If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
End Sub


Private Function DatosOK() As Boolean
Dim B As Boolean
Dim SQL As String
Dim Sql2 As String
Dim vClien As cSocio
    
    B = True
    If txtCodigo(9).Text = "" Or txtCodigo(10).Text = "" Or txtCodigo(11).Text = "" Then
        MsgBox " ", vbExclamation
        B = False
    End If
    DatosOK = B

End Function


Private Function ActualizarTabla(cTabla As String, cWhere As String) As Boolean
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim cadMen As String
Dim I As Long
Dim B As Boolean
Dim vTipoMov As CTiposMov
Dim numalbar As Long
Dim devuelve As String
Dim Existe As Boolean
Dim NumRegis As Long

Dim cTabla2 As String
Dim cWhere2 As String
Dim RS1 As ADODB.Recordset

    On Error GoTo eActualizarTabla
    
    ActualizarTabla = False

'rhisfruta
'numalbar,fecalbar,codvarie,codsocio,codcampo,tipoentr,recolect,kilosbru,numcajon,kilosnet,
'imptrans , impacarr, imprecol, imppenal, impreso
'
'rhisfruta_entradas
'numalbar,numnotac,fechaent,horaentr,kilosbru,numcajon,kilosnet,observac,imptrans,impacarr,
'imprecol , imppenal
'
'rhisfruta_clasif
'numalbar , CodVarie, codcalid, KilosNet
'
'rhisfruta_incidencia
'numalbar , numnotac, codincid


    ' [Monica] 04/06/2010 cargamos la temporal con las entradas selecccionadas que tengan clasificacion
    BorrarTMPNotas
    B = CrearTMPNotas()
    If Not B Then
         Exit Function
    End If
    
    
    SQL = "insert into tmpNotas (numnotac, kilosnet) select rclasifica.numnotac, "
    SQL = SQL & "  sum(rclasifica_clasif.kilosnet) kilos from (" & QuitarCaracterACadena(QuitarCaracterACadena(cTabla, "{"), "}")
    SQL = SQL & ") inner join rclasifica_clasif on rclasifica.numnotac = rclasifica_clasif.numnotac "
    If cWhere <> "" Then
        SQL = SQL & " where " & QuitarCaracterACadena(QuitarCaracterACadena(cWhere, "}"), "{")
    End If
    SQL = SQL & " group by 1  having sum(rclasifica_clasif.kilosnet) <> 0 "
    
    conn.Execute SQL
    ' 04/06/2010 tendremos que enlazar en todas partes con tmpclasifica


    cTabla2 = "((" & cTabla & ") INNER JOIN rsocios on rclasifica.codsocio = rsocios.codsocio) "
    cTabla2 = cTabla2 & " INNER JOIN tmpNotas ON rclasifica.numnotac = tmpNotas.numnotac"
    
    Sql2 = "select rclasifica.codsocio , rsocios.tipoprod from " & cTabla2
    cTabla2 = QuitarCaracterACadena(cTabla2, "{")
    cTabla2 = QuitarCaracterACadena(cTabla2, "}")
    
    If cWhere <> "" Then
        cWhere2 = QuitarCaracterACadena(cWhere, "{")
        cWhere2 = QuitarCaracterACadena(cWhere, "}")
        cWhere2 = QuitarCaracterACadena(cWhere, "_1")
        Sql2 = Sql2 & " WHERE " & cWhere2
    End If
    Sql2 = Sql2 & " GROUP BY 1,2 "
    
    
    Pb1.visible = True
    lblProgres.visible = True
    
    Sql1 = "select count(*) from (" & Sql2 & ") as total"
    
    NumRegis = TotalRegistros(Sql1)
    If NumRegis = 0 Then
        ActualizarTabla = False
        Pb1.visible = False
        lblProgres.visible = False
        MsgBox "No se han podido actualizar registros", vbExclamation
        Exit Function
    End If
    
    
    Me.Pb1.Max = NumRegis
    Me.Refresh
    Me.Pb1.Value = 0
    
    Set RS1 = New ADODB.Recordset
    RS1.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    conn.BeginTrans
    
    I = 0
    B = True
    
    While Not RS1.EOF And B
            
        I = I + 1
        
        Me.Pb1.Value = Me.Pb1.Value + 1
        lblProgres.Caption = "Linea: " & I & ". Socio: " & Format(DBLet(RS1!Codsocio, "N"), "00000000")
        Me.Refresh
        
        cTabla = QuitarCaracterACadena(cTabla, "{")
        cTabla = QuitarCaracterACadena(cTabla, "}")
        
        '[Monica]24/09/2013: en el caso de ser Picassent no tengo en cuenta si es tercero o no para agrupar
        If ((DBLet(RS1.Fields(1).Value, "N") <> 1) And vParamAplic.SeAgrupanNotas And vParamAplic.Cooperativa <> 2 And vParamAplic.Cooperativa <> 16) Or _
           (vParamAplic.SeAgrupanNotas And (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16)) Then   ' caso de no ser tercero
            ' si no es tercero y se agrupan notas
            '[Monica]30/01/2014: en el caso de Alzira se rompe tambien por capataz
            If vParamAplic.Cooperativa = 4 Then
                SQL = "Select rclasifica.fechaent,rclasifica.codcampo,rclasifica.tipoentr,rclasifica.recolect,rclasifica.codvarie,rclasifica.codsocio,rclasifica.transportadopor, rclasifica.codcapat FROM " & QuitarCaracterACadena(cTabla, "_1")
            Else
                '[Monica]04/10/2016: Coopic rompe tb por nro de documento
                If vParamAplic.Cooperativa = 16 Then
                    SQL = "Select rclasifica.fechaent,rclasifica.codcampo,rclasifica.tipoentr,rclasifica.recolect,rclasifica.codvarie,rclasifica.codsocio,rclasifica.transportadopor,rclasifica.contrato FROM " & QuitarCaracterACadena(cTabla, "_1")
                Else
                    SQL = "Select rclasifica.fechaent,rclasifica.codcampo,rclasifica.tipoentr,rclasifica.recolect,rclasifica.codvarie,rclasifica.codsocio,rclasifica.transportadopor FROM " & QuitarCaracterACadena(cTabla, "_1")
                End If
            End If
            SQL = SQL & ", tmpNotas "
            If cWhere <> "" Then
                cWhere = QuitarCaracterACadena(cWhere, "{")
                cWhere = QuitarCaracterACadena(cWhere, "}")
                cWhere = QuitarCaracterACadena(cWhere, "_1")
                SQL = SQL & " WHERE " & cWhere & " and rclasifica.codsocio = " & DBSet(RS1!Codsocio, "N")
            Else
                SQL = SQL & " where rclasifica.codsocio = " & DBSet(RS1!Codsocio, "N")
            End If
            SQL = SQL & " and rclasifica.numnotac = tmpNotas.numnotac "
            '[Monica]30/01/2014: en el caso de Alzira se rompe tambien por capataz
            '        04/10/2016: coopic agrupado por contrato
            If vParamAplic.Cooperativa = 4 Or vParamAplic.Cooperativa = 16 Then
                SQL = SQL & " GROUP BY 1,2,3,4,5,6,7,8 "
            Else
                SQL = SQL & " GROUP BY 1,2,3,4,5,6,7 "
            End If
                
            Set Rs = New ADODB.Recordset
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not Rs.EOF And B
                Set vTipoMov = New CTiposMov
                If vTipoMov.Leer(CodTipoMov) Then
                    numalbar = vTipoMov.ConseguirContador(CodTipoMov)
        
                    Do
                        devuelve = DevuelveDesdeBDNew(cAgro, "rhisfruta", "numalbar", "numalbar", CStr(numalbar), "N")
                        If devuelve <> "" Then
                            'Ya existe el contador incrementarlo
                            Existe = True
                            vTipoMov.IncrementarContador (CodTipoMov)
                            numalbar = vTipoMov.ConseguirContador(CodTipoMov)
                        Else
                            Existe = False
                        End If
                    Loop Until Not Existe
        
                    
                    B = InsertarCabecera(Rs, numalbar, cadMen, False)
                    cadMen = "Insertando Cabecera: " & cadMen
                    
                    If B Then
                        B = InsertarEntradas(Rs, numalbar, cadMen, False)
                        cadMen = "Insertando Entradas: " & cadMen
                    End If
                    
                    
                    If B Then
                        B = InsertarClasificacion(Rs, numalbar, cadMen, False)
                        cadMen = "Insertando Clasificacion: " & cadMen
                    End If
                    
                    If B Then
                        B = InsertarIncidencias(Rs, numalbar, cadMen, False)
                        cadMen = "Insertando Incidencias: " & cadMen
                    End If
                    
                    If B Then
                        B = RecalculaPrecioEstimadoCabecera(numalbar, cadMen)
                        cadMen = "Recalculando Precio Estimado Cabecera: " & cadMen
                    End If
                    
                    '[Monica]14/10/2010 a�ado la condicion de que no sea Picassent
                    '[Monica]27/04/2010 Calculo de costes de transporte, si es por tarifas y la entrada no es de venta campo
                    If vParamAplic.TipoPortesTRA And DBLet(Rs!TipoEntr, "N") <> 1 And vParamAplic.Cooperativa <> 2 Then 'And vParamAplic.Cooperativa <> 16 Then
                        If B Then
                            B = CalculoGastosTransporte(numalbar, cadMen, False)
                            cadMen = "Calculando Gastos de Transporte: " & cadMen
                        End If
                    Else
                        '[Monica]25/02/2011 a�ado la condicion de que sea Picassent
                        If vParamAplic.TipoPortesTRA And DBLet(Rs!TipoEntr, "N") <> 1 Then
                            B = CalculoGastosTransporte(numalbar, cadMen, True)
                            cadMen = "Calculando Gastos de Transporte: " & cadMen
                        End If
                    End If
                    
                    '[Monica]04/05/2010 Reparto de albaranes
                    If B And Not vParamAplic.CooproenEntradas Then
                        B = RepartoAlbaranes(numalbar, cadMen)
                        cadMen = "Reparto Coopropietarios: " & cadMen
                    End If
                    
                    
                    If B Then
                        B = EliminarRegistro(Rs, cadMen, False)
                        cadMen = "Eliminando Registro: " & cadMen
                    End If
                    
                    If B Then
                        cadMen = "Error al actualizar el contador del Pedido."
                        vTipoMov.IncrementarContador (CodTipoMov)
                    End If
                Else
                    B = False
                End If
                
                Set vTipoMov = Nothing
                
                Rs.MoveNext
            Wend
        Else
            ' caso de ser un socio tercero
            ' o no se agrupan notas
            '[Monica]30/01/2014: en el caso de alzira se rompe por capataz
            If vParamAplic.Cooperativa = 4 Then
                SQL = "Select rclasifica.fechaent,rclasifica.codcampo,rclasifica.tipoentr,rclasifica.recolect,rclasifica.codvarie,rclasifica.codsocio, rclasifica.numnotac,rclasifica.transportadopor, rclasifica.codcapat FROM " & QuitarCaracterACadena(cTabla, "_1")
            Else
                If vParamAplic.Cooperativa = 16 Then
                    SQL = "Select rclasifica.fechaent,rclasifica.codcampo,rclasifica.tipoentr,rclasifica.recolect,rclasifica.codvarie,rclasifica.codsocio, rclasifica.numnotac,rclasifica.transportadopor, rclasifica.contrato FROM " & QuitarCaracterACadena(cTabla, "_1")
                Else
                    SQL = "Select rclasifica.fechaent,rclasifica.codcampo,rclasifica.tipoentr,rclasifica.recolect,rclasifica.codvarie,rclasifica.codsocio, rclasifica.numnotac,rclasifica.transportadopor FROM " & QuitarCaracterACadena(cTabla, "_1")
                End If
            End If
            SQL = SQL & ", tmpNotas "
            If cWhere <> "" Then
                cWhere = QuitarCaracterACadena(cWhere, "{")
                cWhere = QuitarCaracterACadena(cWhere, "}")
                cWhere = QuitarCaracterACadena(cWhere, "_1")
                SQL = SQL & " WHERE " & cWhere & " and rclasifica.codsocio = " & DBSet(RS1!Codsocio, "N")
            Else
                SQL = SQL & " where rclasifica.codsocio = " & DBSet(RS1!Codsocio, "N")
            End If
            SQL = SQL & " and rclasifica.numnotac = tmpNotas.numnotac "
            '[Monica]30/01/2014: en el caso de alzira se rompe por capataz
            If vParamAplic.Cooperativa = 4 Or vParamAplic.Cooperativa = 16 Then
                SQL = SQL & " GROUP BY 1,2,3,4,5,6,7,8,9 "
            Else
                SQL = SQL & " GROUP BY 1,2,3,4,5,6,7,8 "
            End If
                
                
            Set Rs = New ADODB.Recordset
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
            
            While Not Rs.EOF And B
                Set vTipoMov = New CTiposMov
                If vTipoMov.Leer(CodTipoMov) Then
                    ' si no se respeta el nro de nota se coge el nro de albaran del contador
                    If Not vParamAplic.SeRespetaNota Then
                    
                        numalbar = vTipoMov.ConseguirContador(CodTipoMov)
            
                        Do
                            devuelve = DevuelveDesdeBDNew(cAgro, "rhisfruta", "numalbar", "numalbar", CStr(numalbar), "N")
                            If devuelve <> "" Then
                                'Ya existe el contador incrementarlo
                                Existe = True
                                vTipoMov.IncrementarContador (CodTipoMov)
                                numalbar = vTipoMov.ConseguirContador(CodTipoMov)
                            Else
                                Existe = False
                            End If
                        Loop Until Not Existe
                    Else
                    
                    ' se respeta iguala el nro de albaran con el nro de nota
                        numalbar = DBLet(Rs.Fields!NumNotac, "N")
                    End If
        
                    B = InsertarCabecera(Rs, numalbar, cadMen, True)
                    cadMen = "Insertando Cabecera: " & cadMen
                    
                    If B Then
                        B = InsertarEntradas(Rs, numalbar, cadMen, True)
                        cadMen = "Insertando Entradas: " & cadMen
                    End If
                    
                    
                    If B Then
                        B = InsertarClasificacion(Rs, numalbar, cadMen, True)
                        cadMen = "Insertando Clasificacion: " & cadMen
                    End If
                    
                    If B Then
                        B = InsertarIncidencias(Rs, numalbar, cadMen, True)
                        cadMen = "Insertando Incidencias: " & cadMen
                    End If
                    
                    If B Then
                        B = RecalculaPrecioEstimadoCabecera(numalbar, cadMen)
                        cadMen = "Recalculando Precio Estimado Cabecera: " & cadMen
                    End If
                    
                    '[Monica]14/10/2010 a�ado la condicion de que no sea Picassent
                    '[Monica]27/04/2010 Calculo de costes de transporte, si es por tarifas y la entrada no es de venta campo
                    If vParamAplic.TipoPortesTRA And DBLet(Rs!TipoEntr, "N") <> 1 And vParamAplic.Cooperativa <> 2 And vParamAplic.Cooperativa <> 16 Then
                        If B Then
                            B = CalculoGastosTransporte(numalbar, cadMen, False)
                            cadMen = "Calculando Gastos de Transporte: " & cadMen
                        End If
                    Else
                        '[Monica]25/02/2011 a�ado la condicion de que sea Picassent
                        If vParamAplic.TipoPortesTRA And DBLet(Rs!TipoEntr, "N") <> 1 Then
                            B = CalculoGastosTransporte(numalbar, cadMen, True)
                            cadMen = "Calculando Gastos de Transporte: " & cadMen
                        End If
                    End If

                    '[Monica]04/05/2010 Reparto de albaranes
                    If B And Not vParamAplic.CooproenEntradas Then
                        B = RepartoAlbaranes(numalbar, cadMen)
                        cadMen = "Reparto Coopropietarios: " & cadMen
                    End If
                    
                    '[Monica]10/03/2017: para el caso de natural hay que cargar la entrada en el historico de ariagro2
                    If B And vParamAplic.Cooperativa = 9 And vEmpresa.BDAriagro = "ariagro1" Then
                        'B = InsertarHcoHortonature(Rs!NumNotac, numalbar, cadMen)
                        'cadMen = "Inserci�n en Hco de Entradas de Hortonature: " & cadMen
                        '[Monica]03/04/2017: no insertamos en el hco sino en la clasificacion sin la datos de la clasificacion para que los introduzcan
                        B = InsertarClasificaHortonature(Rs!NumNotac, numalbar, cadMen)
                        cadMen = "Inserci�n en Entradas Clasificadas de Hortonature: " & cadMen
                    End If
                    
                    
                    If B Then
                        B = EliminarRegistro(Rs, cadMen, True)
                        cadMen = "Eliminando Registro: " & cadMen
                    End If
                    
                    If B Then
                        cadMen = "Error al actualizar el contador del Pedido."
                        vTipoMov.IncrementarContador (CodTipoMov)
                    End If
                Else
                    B = False
                End If
                
                Set vTipoMov = Nothing
                
                Rs.MoveNext
            Wend
        
        
        End If
        
        Set Rs = Nothing
        
        RS1.MoveNext
    Wend
    
' 11-06-2009
' antes de tener en cuenta si era o no tercero el socio
'
'    cTabla = QuitarCaracterACadena(cTabla, "{")
'    cTabla = QuitarCaracterACadena(cTabla, "}")
'    SQL = "Select rclasifica.fechaent,rclasifica.codcampo,rclasifica.tipoentr,rclasifica.recolect,rclasifica.codvarie,rclasifica.codsocio FROM " & QuitarCaracterACadena(cTabla, "_1")
'    If cWhere <> "" Then
'        cWhere = QuitarCaracterACadena(cWhere, "{")
'        cWhere = QuitarCaracterACadena(cWhere, "}")
'        cWhere = QuitarCaracterACadena(cWhere, "_1")
'        SQL = SQL & " WHERE " & cWhere
'    End If
'    SQL = SQL & " GROUP BY 1,2,3,4,5,6 "
    
'    Pb1.visible = True
'    lblProgres.visible = True
'
'    Sql1 = "select count(*) from (" & SQL & ") as total"
'
'    NumRegis = TotalRegistros(Sql1)
'    If NumRegis = 0 Then
'        ActualizarTabla = False
'        Pb1.visible = False
'        lblProgres.visible = False
'        MsgBox "No se han podido actualizar registros", vbExclamation
'        Exit Function
'    End If
'
'    Me.Pb1.Max = NumRegis
'    Me.Refresh
'    Me.Pb1.Value = 0
'
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    conn.BeginTrans
'
'    I = 0
'    b = True
'    While Not RS.EOF And b
'        I = I + 1
'
'        Me.Pb1.Value = Me.Pb1.Value + 1
'        lblProgres.Caption = "Linea: " & I & ". Campo: " & Format(DBLet(RS!codCampo, "N"), "00000000")
'        Me.Refresh
'
'        Set vTipoMov = New CTiposMov
'        If vTipoMov.Leer(CodTipoMov) Then
'            numalbar = vTipoMov.ConseguirContador(CodTipoMov)
'
'            Do
'                devuelve = DevuelveDesdeBDNew(cAgro, "rhisfruta", "numalbar", "numalbar", CStr(numalbar), "N")
'                If devuelve <> "" Then
'                    'Ya existe el contador incrementarlo
'                    Existe = True
'                    vTipoMov.IncrementarContador (CodTipoMov)
'                    numalbar = vTipoMov.ConseguirContador(CodTipoMov)
'                Else
'                    Existe = False
'                End If
'            Loop Until Not Existe
'
'
'            b = InsertarCabecera(RS, numalbar, cadMen)
'            cadMen = "Insertando Cabecera: " & cadMen
'
'            If b Then
'                b = InsertarEntradas(RS, numalbar, cadMen)
'                cadMen = "Insertando Entradas: " & cadMen
'            End If
'
'
'            If b Then
'                b = InsertarClasificacion(RS, numalbar, cadMen)
'                cadMen = "Insertando Clasificacion: " & cadMen
'            End If
'
'            If b Then
'                b = InsertarIncidencias(RS, numalbar, cadMen)
'                cadMen = "Insertando Incidencias: " & cadMen
'            End If
'
'            If b Then
'                b = RecalculaPrecioEstimadoCabecera(numalbar, cadMen)
'                cadMen = "Recalculando Precio Estimado Cabecera: " & cadMen
'            End If
'
'
'            If b Then
'                b = EliminarRegistro(RS, cadMen)
'                cadMen = "Eliminando Registro: " & cadMen
'            End If
'
'            If b Then
'                cadMen = "Error al actualizar el contador del Pedido."
'                vTipoMov.IncrementarContador (CodTipoMov)
'            End If
'        Else
'            b = False
'        End If
'
'        Set vTipoMov = Nothing
'
'        RS.MoveNext
'    Wend
    
eActualizarTabla:
    If Err.Number <> 0 Or Not B Then
        B = False
        MuestraError Err.Number, "Actualizando Entrada", Err.Description & cadMen
    End If
    If B Then
        conn.CommitTrans
        ActualizarTabla = True
    Else
        conn.RollbackTrans
        ActualizarTabla = False
    End If
End Function


Private Function InsertarCabecera(ByRef Rs As ADODB.Recordset, Albaran As Long, cadErr As String, Estercero As Boolean) As Boolean
'Insertando en tabla conta.cabfact
Dim SQL As String
Dim Sql1 As String
Dim RS1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim cad As String
Dim NumCajones As Currency
Dim Transporte As Currency
Dim vPrecio As String
Dim Precio As Currency
Dim AlbaranE As String

    On Error GoTo EInsertar
    

    cadErr = ""

'rhisfruta
'numalbar,fecalbar,codvarie,codsocio,codcampo,tipoentr,recolect,kilosbru,numcajon,kilosnet,
'imptrans , impacarr, imprecol, imppenal, impreso
    
    SQL = "insert into rhisfruta (numalbar,fecalbar,codvarie,codsocio,codcampo,tipoentr,recolect,transportadopor,kilosbru,"
    SQL = SQL & "numcajon,kilosnet,imptrans,impacarr,imprecol,imppenal,impreso,kilostra,contrato ) values "

    Sql1 = "select sum(kilosbru) as kilosbru ,sum(numcajon) as numcajon,sum(rclasifica.kilosnet) as kilosnet,sum(imptrans) as imptrans, sum(impacarr) as impacarr,"
    Sql1 = Sql1 & " sum(imprecol) as imprecol,sum(imppenal) as imppenal,sum(rclasifica.kilostra) as kilostra from rclasifica, tmpNotas "
    Sql1 = Sql1 & " where rclasifica.fechaent = " & DBSet(Rs!FechaEnt, "F") & " and "
    Sql1 = Sql1 & " rclasifica.codcampo =" & DBSet(Rs!codcampo, "N") & " and "
    Sql1 = Sql1 & " rclasifica.tipoentr =" & DBSet(Rs!TipoEntr, "N") & " and "
    Sql1 = Sql1 & " rclasifica.recolect =" & DBSet(Rs!Recolect, "N") & " and "
    Sql1 = Sql1 & " rclasifica.codvarie =" & DBSet(Rs!codvarie, "N") & " and "
    Sql1 = Sql1 & " rclasifica.codsocio =" & DBSet(Rs!Codsocio, "N") & " and "
    
    '[Monica]30/01/2014: en el caso de alzira se rompe por capataz
    If vParamAplic.Cooperativa = 4 Then
        Sql1 = Sql1 & " rclasifica.codcapat =" & DBSet(Rs!codcapat, "N") & " and "
    End If
    '[Monica]04/10/2016: para el caso de Coopic miramos el contrato
    If vParamAplic.Cooperativa = 16 Then
        Sql1 = Sql1 & " rclasifica.contrato = " & DBSet(Rs!contrato, "T") & " and "
    End If
    
    Sql1 = Sql1 & " rclasifica.transportadopor =" & DBSet(Rs!transportadopor, "N") & " and "
    Sql1 = Sql1 & " rclasifica.numnotac = tmpNotas.numnotac "
    
    If Estercero Or Not vParamAplic.SeAgrupanNotas Then
        Sql1 = Sql1 & " and rclasifica.numnotac = " & DBSet(Rs!NumNotac, "N")
    End If
    
    Set Rs2 = New ADODB.Recordset
'    Rs2.Open Sql1, conn, adOpenDynamic, adLockOptimistic, adCmdText
    
    Rs2.Open Sql1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    SQL = SQL & "(" & DBSet(Albaran, "N") & ","
    SQL = SQL & DBSet(Rs!FechaEnt, "F") & ","
    SQL = SQL & DBSet(Rs!codvarie, "N") & ","
    SQL = SQL & DBSet(Rs!Codsocio, "N") & ","
    SQL = SQL & DBSet(Rs!codcampo, "N") & ","
    SQL = SQL & DBSet(Rs!TipoEntr, "N") & ","
    SQL = SQL & DBSet(Rs!Recolect, "N") & ","
    SQL = SQL & DBSet(Rs!transportadopor, "N") & ","
    SQL = SQL & DBSet(Rs2.Fields(0).Value, "N") & ","
    SQL = SQL & DBSet(Rs2.Fields(1).Value, "N") & ","
    SQL = SQL & DBSet(Rs2.Fields(2).Value, "N") & ","
    SQL = SQL & DBSet(Rs2.Fields(3).Value, "N") & ","
    SQL = SQL & DBSet(Rs2.Fields(4).Value, "N") & ","
    SQL = SQL & DBSet(Rs2.Fields(5).Value, "N") & ","
    SQL = SQL & DBSet(Rs2.Fields(6).Value, "N") & ","
    SQL = SQL & "0,"
    SQL = SQL & DBSet(Rs2.Fields(7).Value, "N")
    '[Monica]05/10/2016: nuevo campo de contrato para coopic
    If vParamAplic.Cooperativa <> 16 Then
        SQL = SQL & "," & ValorNulo & ")"
    Else
        SQL = SQL & "," & DBSet(Rs!contrato, "T") & ")"
    End If
    Set Rs2 = Nothing
    
    conn.Execute SQL
    
    '[Monica]10/03/2017: para el caso de natural, guardamos el nro de albaran que introdujeron en la nota de entrada
    If vParamAplic.Cooperativa = 9 Then
        AlbaranE = DevuelveValor("select numalbar from rclasifica where numnotac = " & DBSet(Rs!NumNotac, "N"))
        
        SQL = "update rhisfruta set albarentrada = "
        If AlbaranE = "0" Then
            SQL = SQL & ValorNulo
        Else
            SQL = SQL & DBSet(AlbaranE, "N")
        End If
        
        SQL = SQL & " where numalbar = " & DBSet(Albaran, "N")
        
        conn.Execute SQL
    End If
    
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabecera = False
        cadErr = Err.Description
    Else
        InsertarCabecera = True
    End If
End Function


Private Function InsertarEntradas(ByRef Rs As ADODB.Recordset, Albaran As Long, cadErr As String, Estercero As Boolean) As Boolean
'Insertando en tabla conta.cabfact
Dim SQL As String
Dim Sql1 As String
Dim RS1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim cad As String
Dim NumCajones As Currency
Dim Transporte As Currency
Dim vPrecio As String
Dim Precio As Currency

    On Error GoTo EInsertar
    
    cadErr = ""
    
'rhisfruta_entradas
'numalbar,numnotac,fechaent,horaentr,kilosbru,numcajon,kilosnet,observac,imptrans,impacarr,
'imprecol,imppenal
'
    SQL = "insert into rhisfruta_entradas (numalbar,numnotac,fechaent,horaentr,kilosbru,numcajon,"
    SQL = SQL & "observac,kilosnet,imptrans,impacarr,imprecol,imppenal,prestimado,codtrans,codtarif,codcapat,kilostra, tiporecol, horastra, numtraba) "

    SQL = SQL & "select " & Albaran & ",rclasifica.numnotac,fechaent,horaentr,kilosbru,numcajon,"
    SQL = SQL & "observac,rclasifica.kilosnet,imptrans,impacarr,imprecol,imppenal,prestimado,codtrans,codtarif, codcapat, kilostra, "
    '[Monica]28/02/2012: se graban tambien el tipo de recolecion, las horas y el nro de trabajadores
    SQL = SQL & " tiporecol, horastra, numtraba "
    SQL = SQL & " from rclasifica, tmpNotas "
    SQL = SQL & " where rclasifica.fechaent =" & DBSet(Rs!FechaEnt, "F") & " and "
    SQL = SQL & " rclasifica.codcampo =" & DBSet(Rs!codcampo, "N") & " and "
    SQL = SQL & " rclasifica.tipoentr =" & DBSet(Rs!TipoEntr, "N") & " and "
    SQL = SQL & " rclasifica.recolect =" & DBSet(Rs!Recolect, "N") & " and "
    SQL = SQL & " rclasifica.codsocio =" & DBSet(Rs!Codsocio, "N") & " and "
    SQL = SQL & " rclasifica.codvarie =" & DBSet(Rs!codvarie, "N") & " and "
    SQL = SQL & " rclasifica.transportadopor = " & DBSet(Rs!transportadopor, "N") & " and "
    
    '[Monica]30/01/2014: para el caso de Alzira se rompe tambien por capataz
    If vParamAplic.Cooperativa = 4 Then
        SQL = SQL & " rclasifica.codcapat =" & DBSet(Rs!codcapat, "N") & " and "
    End If
    
    SQL = SQL & " rclasifica.numnotac = tmpNotas.numnotac "

    
    If Estercero Or Not vParamAplic.SeAgrupanNotas Then
        SQL = SQL & " and rclasifica.numnotac = " & DBSet(Rs!NumNotac, "N")
    End If
    
    conn.Execute SQL
    
    '[Monica] 18/11/2010: en el caso de alzira grabamos los kilos transportados para la factura de acarreo recoleccion socio
    If vParamAplic.Cooperativa = 4 Then
        SQL = "update rhisfruta_entradas set kilostra = (select sum(kilosnet) from rclasifica_clasif, rcalidad "
        SQL = SQL & " where rclasifica_clasif.codvarie = rcalidad.codvarie and "
        SQL = SQL & " rclasifica_clasif.codcalid = rcalidad.codcalid and "
        SQL = SQL & " rcalidad.gastosrec = 1 and "
        SQL = SQL & " rclasifica_clasif.numnotac = rhisfruta_entradas.numnotac)"
        SQL = SQL & " where rhisfruta_entradas.numalbar = " & Albaran

        conn.Execute SQL
        
        SQL = "update rhisfruta set kilostra = (select sum(kilostra) from rhisfruta_entradas where numalbar = " & Albaran & ")"
        SQL = SQL & " where numalbar = " & Albaran
        conn.Execute SQL
        
    End If
    
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarEntradas = False
        cadErr = Err.Description
    Else
        InsertarEntradas = True
    End If
End Function


Private Function RecalculaPrecioEstimadoCabecera(Albaran As Long, cadErr As String) As Boolean
'Insertando en tabla conta.cabfact
Dim SQL As String
Dim RS1 As ADODB.Recordset
Dim cad As String
Dim Precio As Currency

    On Error GoTo EInsertar
    
    cadErr = ""
    
    SQL = "select count(*), sum(prestimado) from rhisfruta_entradas where numalbar = " & DBSet(Albaran, "N")
    
    Set RS1 = New ADODB.Recordset
    RS1.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    Precio = 0
    If Not RS1.EOF Then
        If DBLet(RS1.Fields(0).Value, "N") <> 0 Then
            Precio = Round2(DBLet(RS1.Fields(1).Value, "N") / DBLet(RS1.Fields(0).Value, "N"), 4)
        End If
    End If
    
    SQL = "update rhisfruta set prestimado = " & DBSet(Precio, "N") & " where numalbar = " & DBSet(Albaran, "N")
    conn.Execute SQL
    
EInsertar:
    If Err.Number <> 0 Then
        RecalculaPrecioEstimadoCabecera = False
        cadErr = Err.Description
    Else
        RecalculaPrecioEstimadoCabecera = True
    End If
End Function



Private Function InsertarClasificacion(ByRef Rs As ADODB.Recordset, Albaran As Long, cadErr As String, Estercero As Boolean) As Boolean
'Insertando en tabla conta.cabfact
Dim SQL As String
Dim Sql1 As String
Dim RS1 As ADODB.Recordset
Dim cad As String
Dim KilosMuestra As Currency
Dim TotalKilos As Currency
Dim Calidad As Currency
Dim Diferencia As Currency
Dim HayReg As Byte
Dim TipoClasif As Byte
Dim vTipoClasif As String
Dim vCalidDest As String
Dim CalidadClasif As String


    On Error GoTo EInsertar


    cadErr = ""
'rhisfruta_clasif
'numalbar , CodVarie, codcalid, KilosNet
'
    SQL = "insert into rhisfruta_clasif (numalbar, codvarie, codcalid, kilosnet)  "

    SQL = SQL & "select " & Albaran & ",rclasifica_clasif.codvarie, codcalid, sum(rclasifica_clasif.kilosnet) "
    SQL = SQL & " from rclasifica_clasif, rclasifica, tmpNotas "
    SQL = SQL & " where rclasifica.fechaent =" & DBSet(Rs!FechaEnt, "F") & " and "
    SQL = SQL & " rclasifica.codcampo =" & DBSet(Rs!codcampo, "N") & " and "
    SQL = SQL & " rclasifica.tipoentr =" & DBSet(Rs!TipoEntr, "N") & " and "
    SQL = SQL & " rclasifica.recolect =" & DBSet(Rs!Recolect, "N") & " and "
    SQL = SQL & " rclasifica.codvarie =" & DBSet(Rs!codvarie, "N") & " and "
    SQL = SQL & " rclasifica.codsocio =" & DBSet(Rs!Codsocio, "N") & " and "
    SQL = SQL & " rclasifica.transportadopor = " & DBSet(Rs!transportadopor, "N") & " and "
    SQL = SQL & " rclasifica.numnotac = rclasifica_clasif.numnotac and "
    
    '[Monica]30/01/2014: para el caso de Alzira se rompia tambien por capataz
    If vParamAplic.Cooperativa = 4 Then
        SQL = SQL & " rclasifica.codcapat =" & DBSet(Rs!codcapat, "N") & " and "
    End If
    
    SQL = SQL & " rclasifica.numnotac = tmpNotas.numnotac "
    
    If Estercero Or Not vParamAplic.SeAgrupanNotas Then
        SQL = SQL & " and rclasifica.numnotac = " & DBSet(Rs!NumNotac, "N")
    End If
    
    SQL = SQL & " group by 1,2,3"

    conn.Execute SQL

EInsertar:
    If Err.Number <> 0 Then
        InsertarClasificacion = False
        cadErr = Err.Description
    Else
        InsertarClasificacion = True
    End If
End Function


Private Function InsertarIncidencias(ByRef Rs As ADODB.Recordset, Albaran As Long, cadErr As String, Estercero As Boolean) As Boolean
Dim SQL As String
Dim Sql1 As String
Dim RS1 As ADODB.Recordset
Dim cad As String
Dim KilosMuestra As Currency
Dim TotalKilos As Currency
Dim Calidad As Currency
Dim Diferencia As Currency
Dim HayReg As Byte
Dim TipoClasif As Byte
Dim vTipoClasif As String
Dim vCalidDest As String
Dim CalidadClasif As String


    On Error GoTo EInsertar
    
    
    cadErr = ""
'rhisfruta_incidencia
'numalbar , numnotac, codincid

    SQL = "insert into rhisfruta_incidencia (numalbar, numnotac, codincid)  "

    SQL = SQL & "select " & Albaran & ",rclasifica_incidencia.numnotac, rclasifica_incidencia.codincid "
    SQL = SQL & " from rclasifica_incidencia, rclasifica, tmpNotas "
    SQL = SQL & " where rclasifica.fechaent =" & DBSet(Rs!FechaEnt, "F") & " and "
    SQL = SQL & " rclasifica.codcampo =" & DBSet(Rs!codcampo, "N") & " and "
    SQL = SQL & " rclasifica.tipoentr =" & DBSet(Rs!TipoEntr, "N") & " and "
    SQL = SQL & " rclasifica.recolect =" & DBSet(Rs!Recolect, "N") & " and "
    SQL = SQL & " rclasifica.codvarie =" & DBSet(Rs!codvarie, "N") & " and "
    SQL = SQL & " rclasifica.codsocio =" & DBSet(Rs!Codsocio, "N") & " and "
    SQL = SQL & " rclasifica.transportadopor =" & DBSet(Rs!transportadopor, "N") & " and "
    
    '[Monica]30/01/2014: para el caso de Alzira se agrupa tambien por capataz
    If vParamAplic.Cooperativa = 4 Then
        SQL = SQL & " rclasifica.codcapat =" & DBSet(Rs!codcapat, "N") & " and "
    End If
    
    SQL = SQL & " rclasifica.numnotac = rclasifica_incidencia.numnotac and  "
    SQL = SQL & " rclasifica.numnotac = tmpNotas.numnotac "
    
    If Estercero Or Not vParamAplic.SeAgrupanNotas Then
        SQL = SQL & " and rclasifica.numnotac = " & DBSet(Rs!NumNotac, "N")
    End If
    
    SQL = SQL & " group by 1,2,3"
    
    conn.Execute SQL
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarIncidencias = False
        cadErr = Err.Description
    Else
        InsertarIncidencias = True
    End If
End Function

Private Function CalculoGastosTransporte(Albaran As Long, cadErr As String, EsPicassent As Boolean) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim PrecTarifaAlm As Currency
Dim PrecTarifaAlm2 As Currency
Dim PrecTarifa As Currency
Dim ImpTrans As Currency
Dim TotImpTrans As Currency
Dim ImpGastoSocio As Currency
Dim NumF As String

On Error GoTo EInsertar
    
    ' calculamos los gastos de transporte para el socio y el importe de gastos de transporte de rhisfruta
    ' a partir de la entradas que ya hemos grabado en la rhisfruta_entradas

    cadErr = ""

    '[Monica]25/02/2011: Si no es Picassent
    If Not EsPicassent Then
        SQL = "select numnotac, rhisfruta_entradas.codtarif, rtarifatra.tipotarifa, sum(kilosnet) as kilos "
    Else
        SQL = "select numnotac, rhisfruta_entradas.codtarif, rtarifatra.tipotarifa, sum(kilostra) as kilos "
    End If
    
    SQL = SQL & " from rhisfruta_entradas, rtarifatra where numalbar = " & DBSet(Albaran, "N")
    SQL = SQL & " and rhisfruta_entradas.codtarif = rtarifatra.codtarif "
    SQL = SQL & " and rtarifatra.tipotarifa <> 2 " 'las tarifas que buscamos son del tipo 1 o 2 (no sin asignar)
    SQL = SQL & " group by 1, 2, 3 order by 1, 2, 3 "

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    PrecTarifaAlm = DevuelveValor("select preciokg from rtarifatra where codtarif = " & vParamAplic.TarifaTRA)
    PrecTarifaAlm2 = DevuelveValor("select preciokg from rtarifatra where codtarif = " & vParamAplic.TarifaTRA2)
    
    ImpTrans = 0
    TotImpTrans = 0
    ImpGastoSocio = 0
    
    While Not Rs.EOF
        PrecTarifa = DevuelveValor("select preciokg from rtarifatra where codtarif = " & DBSet(Rs!Codtarif, "N"))
        
        ImpTrans = Round2(PrecTarifa * DBLet(Rs!Kilos, "N"), 2)
        
        TotImpTrans = TotImpTrans + ImpTrans
            
        '[Monica]25/02/2011: Si no es Picassent
        If Not EsPicassent Then
            SQL = "update rhisfruta_entradas set imptrans = " & DBSet(ImpTrans, "N")
            SQL = SQL & " where numalbar = " & DBSet(Albaran, "N") & " and numnotac = " & DBSet(Rs!NumNotac, "N")
            
            conn.Execute SQL
        End If
        
        If DBLet(Rs!tipotarifa, "N") = 0 Then ' Tarifa de Tipo 1
            If (PrecTarifa - PrecTarifaAlm) > 0 Then
                ImpGastoSocio = ImpGastoSocio + Round2((DBLet(Rs!Kilos, "N") * (PrecTarifa - PrecTarifaAlm)), 2)
            End If
        Else ' Tarifa de tipo 2
            If (PrecTarifa - PrecTarifaAlm2) > 0 Then
                ImpGastoSocio = ImpGastoSocio + Round2((DBLet(Rs!Kilos, "N") * (PrecTarifa - PrecTarifaAlm2)), 2)
            End If
        End If
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    '[Monica]25/02/2011: Si no es Picassent
    If Not EsPicassent Then
        ' actualizamos cabecera
        SQL = "update rhisfruta set imptrans = " & DBSet(TotImpTrans, "N")
        SQL = SQL & " where numalbar = " & DBSet(Albaran, "N")
        
        conn.Execute SQL
    End If
    
    '[Monica] s�lo insertamos cuando el importe total de gasto socio es positivo
    If ImpGastoSocio > 0 Then
        NumF = ""
        NumF = SugerirCodigoSiguienteStr("rhisfruta_gastos", "numlinea", "numalbar = " & DBSet(Albaran, "N"))
        ' grabamos un registro en con los gastos del cliente
        SQL = "insert into rhisfruta_gastos (numalbar, numlinea, codgasto, importe) values (" & DBSet(Albaran, "N") & ","
        SQL = SQL & DBSet(NumF, "N") & "," & DBSet(vParamAplic.CodGastoTRA, "N") & "," & DBSet(ImpGastoSocio, "N") & ")"
        
        conn.Execute SQL
    End If
    
EInsertar:
    If Err.Number <> 0 Then
        CalculoGastosTransporte = False
        cadErr = Err.Description
    Else
        CalculoGastosTransporte = True
    End If
End Function




Private Function EliminarRegistro(ByRef Rs As ADODB.Recordset, cadErr As String, Estercero As Boolean) As Boolean
'Insertando en tabla conta.cabfact
Dim SQL As String
Dim RS1 As ADODB.Recordset
Dim cad As String
Dim NumCajones As Currency
Dim Transporte As Currency
Dim vPrecio As String

    On Error GoTo EEliminar
    

    cadErr = ""

    'RCLASIFICA_INCIDENCIA
    SQL = "delete from rclasifica_incidencia where numnotac in (select rclasifica.numnotac from rclasifica, tmpNotas "
    SQL = SQL & " where rclasifica.fechaent =" & DBSet(Rs!FechaEnt, "F") & " and "
    SQL = SQL & " rclasifica.codcampo =" & DBSet(Rs!codcampo, "N") & " and "
    SQL = SQL & " rclasifica.tipoentr =" & DBSet(Rs!TipoEntr, "N") & " and "
    SQL = SQL & " rclasifica.codsocio =" & DBSet(Rs!Codsocio, "N") & " and "
    SQL = SQL & " rclasifica.codvarie =" & DBSet(Rs!codvarie, "N") & " and "
    SQL = SQL & " rclasifica.transportadopor =" & DBSet(Rs!transportadopor, "N") & " and "
    SQL = SQL & " rclasifica.numnotac = tmpNotas.numnotac and "
    
    '[Monica]30/01/2014: en el caso de alzira se rompe por capataz
    If vParamAplic.Cooperativa = 4 Then
        SQL = SQL & " rclasifica.codcapat =" & DBSet(Rs!codcapat, "N") & " and "
    End If
    
    If Estercero Or Not vParamAplic.SeAgrupanNotas Then
        SQL = SQL & " rclasifica.numnotac =" & DBSet(Rs!NumNotac, "N") & " and "
    End If
    
    SQL = SQL & " rclasifica.recolect =" & DBSet(Rs!Recolect, "N") & ") "
    
    conn.Execute SQL

    'RCLASIFICA_CLASIF
    SQL = "delete from rclasifica_clasif where numnotac in (select rclasifica.numnotac from rclasifica, tmpNotas  "
    SQL = SQL & " where rclasifica.fechaent =" & DBSet(Rs!FechaEnt, "F") & " and "
    SQL = SQL & " rclasifica.codcampo =" & DBSet(Rs!codcampo, "N") & " and "
    SQL = SQL & " rclasifica.tipoentr =" & DBSet(Rs!TipoEntr, "N") & " and "
    SQL = SQL & " rclasifica.codsocio =" & DBSet(Rs!Codsocio, "N") & " and "
    SQL = SQL & " rclasifica.codvarie =" & DBSet(Rs!codvarie, "N") & " and "
    SQL = SQL & " rclasifica.transportadopor =" & DBSet(Rs!transportadopor, "N") & " and "
    SQL = SQL & " rclasifica.numnotac = tmpNotas.numnotac and  "
    
    '[Monica]30/01/2014: en el caso de alzira se rompe por capataz
    If vParamAplic.Cooperativa = 4 Then
        SQL = SQL & " rclasifica.codcapat =" & DBSet(Rs!codcapat, "N") & " and "
    End If
    
    If Estercero Or Not vParamAplic.SeAgrupanNotas Then
        SQL = SQL & " rclasifica.numnotac =" & DBSet(Rs!NumNotac, "N") & " and "
    End If
    
    SQL = SQL & " rclasifica.recolect =" & DBSet(Rs!Recolect, "N") & ") "
    
    conn.Execute SQL

    'RCLASIFICA
    SQL = "delete from rclasifica "
    SQL = SQL & " where rclasifica.fechaent =" & DBSet(Rs!FechaEnt, "F") & " and "
    SQL = SQL & " rclasifica.codcampo =" & DBSet(Rs!codcampo, "N") & " and "
    SQL = SQL & " rclasifica.tipoentr =" & DBSet(Rs!TipoEntr, "N") & " and "
    SQL = SQL & " rclasifica.codsocio =" & DBSet(Rs!Codsocio, "N") & " and "
    SQL = SQL & " rclasifica.codvarie =" & DBSet(Rs!codvarie, "N") & " and "
    SQL = SQL & " rclasifica.transportadopor =" & DBSet(Rs!transportadopor, "N") & " and "
    
    '[Monica]30/01/2014: en el caso de alzira se rompe por capataz
    If vParamAplic.Cooperativa = 4 Then
        SQL = SQL & " rclasifica.codcapat =" & DBSet(Rs!codcapat, "N") & " and "
    End If
    
    If Estercero Or Not vParamAplic.SeAgrupanNotas Then
        SQL = SQL & " rclasifica.numnotac =" & DBSet(Rs!NumNotac, "N") & " and "
    End If
    
    SQL = SQL & " rclasifica.recolect =" & DBSet(Rs!Recolect, "N") & " and "
    SQL = SQL & " rclasifica.numnotac in (select numnotac from tmpNotas) "
    
    conn.Execute SQL

EEliminar:
    If Err.Number <> 0 Then
        EliminarRegistro = False
        cadErr = Err.Description
    Else
        EliminarRegistro = True
    End If
End Function



Public Function HayRegSinClasificacion(ByVal cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim vSQL As String
    On Error GoTo eHayRegSinClasificacion
    
    
    HayRegSinClasificacion = True
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    cTabla = "(" & cTabla & ") inner join rclasifica_clasif on rclasifica.numnotac = rclasifica_clasif.numnotac"
    
    SQL = "select rclasifica.numnotac, rclasifica.codsocio, sum(rclasifica_clasif.kilosnet) from " & QuitarCaracterACadena(cTabla, "_1")
    vSQL = "select " & vUsu.Codigo & ", rclasifica.numnotac, rclasifica.codsocio, 0 from " & QuitarCaracterACadena(cTabla, "_1")
    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
        vSQL = vSQL & " WHERE " & cWhere
    End If
    
    SQL = SQL & " group by rclasifica.numnotac, rclasifica.codsocio "
    SQL = SQL & " having sum(rclasifica_clasif.kilosnet) is null "
    
    vSQL = vSQL & " group by rclasifica.numnotac, rclasifica.codsocio "
    vSQL = vSQL & " having sum(rclasifica_clasif.kilosnet) is null "
    
    Sql2 = "select count(*) from (" & SQL & ") as a "
    
    If TotalRegistros(Sql2) <> 0 Then
        Sql3 = "insert into tmpclasifica (codusu, numnotac, codsocio, codclase) "
        Sql3 = Sql3 & vSQL
     
        conn.Execute Sql3
    
        HayRegSinClasificacion = True
    Else
        HayRegSinClasificacion = False
    End If
    
    Exit Function
    
eHayRegSinClasificacion:
    MuestraError Err.Number, "Hay Registros Sin Clasificacion", Err.Description
End Function

Public Function HayRegSinGastosCorrectos(ByVal cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
Dim Sql2 As String
Dim vSQL As String
Dim Rs As ADODB.Recordset
Dim cad As String

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    cTabla = "(" & cTabla & ") inner join rclasifica_clasif on rclasifica.numnotac = rclasifica_clasif.numnotac"
    
    SQL = "select distinct rclasifica.numnotac, rclasifica.codsocio from " & QuitarCaracterACadena(cTabla, "_1")
    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cad = ""
    
    While Not Rs.EOF
        If Not CalculoGastosCorrectos(DBLet(Rs!NumNotac, "N")) Then
            cad = cad & "(" & vUsu.Codigo & "," & DBLet(Rs!NumNotac, "N") & "," & DBLet(Rs!Codsocio, "N") & ",1),"
        End If
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    ' quitamos la ultima coma
    If cad <> "" Then cad = Mid(cad, 1, Len(cad) - 1)
    
    If cad <> "" Then
        HayRegSinGastosCorrectos = True
    
        SQL = "insert into tmpclasifica (codusu, numnotac, codsocio, codclase) values "
        SQL = SQL & cad
     
        conn.Execute SQL
    Else
        HayRegSinGastosCorrectos = False
    End If
    
End Function

Public Function HayRegEnHcoEntradas(ByVal cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim vSQL As String
    
    On Error GoTo eHayRegEnHcoEntradas
    
    HayRegEnHcoEntradas = True
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    
    cTabla = "(" & cTabla & ") inner join rclasifica_clasif on rclasifica.numnotac = rclasifica_clasif.numnotac"
    
    SQL = "select count(*) from rhisfruta where numalbar in ("
    SQL = SQL & "select rclasifica.numnotac from " & QuitarCaracterACadena(cTabla, "_1")
    vSQL = "select distinct " & vUsu.Codigo & ", rclasifica.numnotac, rclasifica.codsocio, 2 from " & QuitarCaracterACadena(cTabla, "_1")
    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
        vSQL = vSQL & " WHERE " & cWhere
    End If
    
    SQL = SQL & ")"
    
    If TotalRegistros(SQL) <> 0 Then
        Sql3 = "insert into tmpclasifica (codusu, numnotac, codsocio, codclase) "
        Sql3 = Sql3 & vSQL
     
        conn.Execute Sql3
    
        HayRegEnHcoEntradas = True
    Else
        HayRegEnHcoEntradas = False
    End If
    
    Exit Function
    
eHayRegEnHcoEntradas:
    MuestraError Err.Number, "Hay Registros en el Hist�rico de Entradas", Err.Description
End Function


'Private Function CalculoGastosCorrectos(NumNota As String) As Boolean
'Dim Rs As ADODB.Recordset
'Dim RS1 As ADODB.Recordset
'Dim SQL As String
'Dim TotalEnvases As String
'Dim TotalCostes As String
'Dim Valor As Currency
'Dim GasRecol As Currency
'Dim GasAcarreo As Currency
'Dim KilosTria As Long
'Dim KilosNet As Long
'Dim EurDesta As Currency
'Dim EurRecol As Currency
'Dim PrecAcarreo As Currency
'Dim I As Integer
'
'    On Error Resume Next
'
'
'    SQL = "select * from rclasifica where numnotac = " & DBSet(NumNota, "N")
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    If Not Rs.EOF Then
'
'        GasRecol = 0
'        GasAcarreo = 0
'
'        If DBLet(Rs!tipoentr, "N") = 1 Then ' es venta campo
'            CalculoGastosCorrectos = True
'            Exit Function
'        End If
'
'        SQL = "select eurdesta, eurecole from variedades where codvarie = " & DBSet(Rs!CodVarie, "N")
'
'        Set RS1 = New ADODB.Recordset
'        RS1.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'        If Not Rs.EOF Then
'            EurDesta = DBLet(RS1.Fields(0).Value, "N")
'            EurRecol = DBLet(RS1.Fields(1).Value, "N")
'        End If
'
'        Set RS1 = Nothing
'
'    '    Sql = "select sum(kilosnet) from rclasifica_clasif where numnotac = " & DBSet(Text1(0).Text, "N")
'    '    KilosNet = TotalRegistros(Sql)
'
'        KilosNet = DBLet(Rs!KilosNet, "N")
'
'        'recolecta socio
'        If DBLet(Rs!Recolect, "N") = 1 Then
'            SQL = "select sum(kilosnet) from rclasifica_clasif, rcalidad  where numnotac = " & DBSet(NumNota, "N")
'            SQL = SQL & " and rclasifica_clasif.codvarie = rcalidad.codvarie "
'            SQL = SQL & " and rclasifica_clasif.codcalid = rcalidad.codcalid "
'            SQL = SQL & " and rcalidad.gastosrec = 1"
'
'            KilosTria = TotalRegistros(SQL)
'
'            GasRecol = Round2(KilosTria * EurRecol, 2)
'        Else
'        'recolecta cooperativa
'            If DBLet(Rs!tiporecol, "N") = 0 Then
'                'horas
'                'gastosrecol = horas * personas * rparam.(costeshora + costesegso)
'                GasRecol = Round2(HorasDecimal(Format(DBLet(Rs!horastra, "N"), "###,##0.00")) * DBLet(Rs!numtraba, "N") * (vParamAplic.CosteHora + vParamAplic.CosteSegSo), 2)
'            Else
'                'destajo
'                GasRecol = Round2(KilosNet * EurDesta, 2)
'            End If
'        End If
''12/05/2009
''        If DBLet(Rs!codtarif, "N") <> 0 Then
''            Sql = DevuelveDesdeBDNew(cAgro, "rtarifatra", "preciokg", "codtarif", Rs!codtarif, "N")
''            PrecAcarreo = CCur(Sql)
''        Else
''            PrecAcarreo = 0
''        End If
''12/05/2009 cambiado por esto pq si que hay tarifa 0
'        PrecAcarreo = 0
'        SQL = ""
'        SQL = DevuelveDesdeBDNew(cAgro, "rtarifatra", "preciokg", "codtarif", DBLet(Rs!codtarif, "N"), "N")
'        If SQL <> "" Then
'            PrecAcarreo = CCur(SQL)
'        End If
'
'        GasAcarreo = Round2(PrecAcarreo * KilosNet, 2)
'
'        CalculoGastosCorrectos = Not (((DBLet(Rs!imprecol, "N") <> GasRecol) Or (DBLet(Rs!impacarr, "N") <> GasAcarreo)))
'    End If
'
'End Function
'


Private Sub BorrarTMPNotas()
On Error Resume Next

    conn.Execute " DROP TABLE IF EXISTS tmpNotas;"
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Function CrearTMPNotas() As Boolean
' temporal para selecccionar unicamente las notas con clasificacion
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPNotas = False
    
    'tmpNotas
    SQL = "CREATE TEMPORARY TABLE tmpNotas ( "
    SQL = SQL & "`numnotac` int(7) NOT NULL ,"
    SQL = SQL & "`kilosnet` int(7) unsigned NOT NULL)"
       
    conn.Execute SQL
     
    CrearTMPNotas = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPNotas = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpNotas;"
        conn.Execute SQL
    End If
End Function




Private Function InsertarHcoHortonature(NumNotac As Long, Albaran As Long, cadErr As String) As Boolean
'Insertando en tabla conta.cabfact
Dim SQL As String
Dim Sql1 As String
Dim RS1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim cad As String
Dim NumCajones As Currency
Dim Transporte As Currency
Dim vPrecio As String
Dim Precio As Currency
Dim AlbaranE As String

    On Error GoTo EInsertar

    cadErr = ""

'rhisfruta
'numalbar,fecalbar,codvarie,codsocio,codcampo,tipoentr,recolect,kilosbru,numcajon,kilosnet,
'imptrans , impacarr, imprecol, imppenal, impreso
    AlbaranE = DevuelveValor("select numalbar from rclasifica where numnotac = " & NumNotac)
    If ComprobarCero(AlbaranE) = 0 Then
        InsertarHcoHortonature = True
        Exit Function
    End If
    
    ' insertamos cabecera
    SQL = "insert into ariagro2.rhisfruta (numalbar,fecalbar,codvarie,codsocio,codcampo,tipoentr,recolect,transportadopor,kilosbru,"
    SQL = SQL & "numcajon,kilosnet,imptrans,impacarr,imprecol,imppenal,impreso,kilostra,contrato ) "
    SQL = SQL & " select " & DBSet(AlbaranE, "N") & ",fecalbar,codvarie,codsocio,codcampo,tipoentr,recolect,transportadopor,kilosbru,numcajon,kilosnet,imptrans,impacarr,"
    SQL = SQL & " imprecol,imppenal,impreso,kilostra,contrato from " & vEmpresa.BDAriagro & ".rhisfruta where numalbar = " & DBSet(Albaran, "N")
    
    conn.Execute SQL

    ' insertamos entradas
    SQL = "insert into ariagro2.rhisfruta_entradas (numalbar,numnotac,fechaent,horaentr,kilosbru,numcajon,"
    SQL = SQL & " observac,kilosnet,imptrans,impacarr,imprecol,imppenal,prestimado,codtrans,codtarif,codcapat,kilostra, tiporecol, horastra, numtraba) "
    SQL = SQL & " select " & DBSet(AlbaranE, "N") & "," & DBSet(AlbaranE, "N") & ",fechaent,horaentr,kilosbru,numcajon,"
    SQL = SQL & " observac,kilosnet,imptrans,impacarr,imprecol,imppenal,prestimado,codtrans,codtarif,codcapat,kilostra, tiporecol, horastra, numtraba "
    SQL = SQL & " from " & vEmpresa.BDAriagro & ".rhisfruta_entradas where numalbar = " & DBSet(Albaran, "N")
    
    conn.Execute SQL
    
    ' insertamos clasificacion
    SQL = "insert into ariagro2.rhisfruta_clasif (numalbar, codvarie, codcalid, kilosnet)  "
    SQL = SQL & " select " & DBSet(AlbaranE, "N") & ", codvarie, codcalid, kilosnet from " & vEmpresa.BDAriagro & ".rhisfruta_clasif where numalbar = " & DBSet(Albaran, "N")

    conn.Execute SQL
    
    ' insertamos en rhisfruta_gastos
    SQL = "insert into ariagro2.rhisfruta_gastos (numalbar,numlinea,codgasto,importe) "
    SQL = SQL & " select " & DBSet(AlbaranE, "N") & ", numlinea, codgasto, importe from " & vEmpresa.BDAriagro & ".rhisfruta_gastos where numalbar = " & DBSet(Albaran, "N")
    
    conn.Execute SQL
    
    ' insertamos en rhisfruta_incidencia
    SQL = "insert into ariagro2.rhisfruta_incidencia (numalbar, numnotac, codincid) "
    SQL = SQL & " select " & DBSet(AlbaranE, "N") & ", " & DBSet(AlbaranE, "N") & ", codincid from " & vEmpresa.BDAriagro & ".rhisfruta_incidencia where numalbar = " & DBSet(Albaran, "N")
    
    conn.Execute SQL
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarHcoHortonature = False
        cadErr = Err.Description
    Else
        InsertarHcoHortonature = True
    End If
End Function


Private Function InsertarClasificaHortonature(NumNotac As Long, Albaran As Long, cadErr As String) As Boolean
'Insertando en tabla conta.cabfact
Dim SQL As String
Dim Sql1 As String
Dim RS1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim cad As String
Dim NumCajones As Currency
Dim Transporte As Currency
Dim vPrecio As String
Dim Precio As Currency
Dim AlbaranE As String

    On Error GoTo EInsertar

    cadErr = ""

'rhisfruta
'numalbar,fecalbar,codvarie,codsocio,codcampo,tipoentr,recolect,kilosbru,numcajon,kilosnet,
'imptrans , impacarr, imprecol, imppenal, impreso
    AlbaranE = DevuelveValor("select numalbar from rclasifica where numnotac = " & NumNotac)
    If ComprobarCero(AlbaranE) = 0 Then
        InsertarClasificaHortonature = True
        Exit Function
    End If
    
    ' insertamos cabecera
    SQL = "insert into ariagro2.rclasifica (numnotac,fechaent,horaentr,codvarie,codsocio,codcampo,tipoentr,recolect,codtrans,codcapat,codtarif,kilosbru,numcajon,kilosnet,observac,imptrans,"
    SQL = SQL & "impacarr,imprecol,imppenal,tiporecol,horastra,numtraba,numalbar,fecalbar,impreso,prestimado,transportadopor,kilostra,contrato) "
    SQL = SQL & " select " & DBSet(AlbaranE, "N") & ",fechaent,horaentr,codvarie,codsocio,codcampo,tipoentr,recolect,codtrans,codcapat,codtarif,kilosbru,numcajon,kilosnet,observac,imptrans,"
    SQL = SQL & "impacarr,imprecol,imppenal,tiporecol,horastra,numtraba,numalbar,fecalbar,impreso,prestimado,transportadopor,kilostra,contrato "
    SQL = SQL & "  from " & vEmpresa.BDAriagro & ".rclasifica where numnotac = " & DBSet(NumNotac, "N")
    
    conn.Execute SQL

    
    ' insertamos clasificacion
    SQL = "insert into ariagro2.rclasifica_clasif (numnotac, codvarie, codcalid)  "
    SQL = SQL & " select " & DBSet(AlbaranE, "N") & ", codvarie, codcalid from " & vEmpresa.BDAriagro & ".rclasifica_clasif where numnotac = " & DBSet(NumNotac, "N")

    conn.Execute SQL
    
    ' insertamos en rclasifica_incidencia
    SQL = "insert into ariagro2.rclasifica_incidencia (numnotac, codincid) "
    SQL = SQL & " select " & DBSet(AlbaranE, "N") & ", codincid from " & vEmpresa.BDAriagro & ".rclasifica_incidencia where numnotac = " & DBSet(Albaran, "N")
    
    conn.Execute SQL
    
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarClasificaHortonature = False
        cadErr = Err.Description
    Else
        InsertarClasificaHortonature = True
    End If
End Function





