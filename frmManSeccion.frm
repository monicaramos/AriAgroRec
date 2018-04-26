VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManSeccion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Secciones"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   9705
   Icon            =   "frmManSeccion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   120
      TabIndex        =   26
      Top             =   60
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   27
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
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   330
      Index           =   3
      Left            =   7335
      MaskColor       =   &H00000000&
      TabIndex        =   23
      ToolTipText     =   "Buscar Tipo Iva"
      Top             =   3900
      Visible         =   0   'False
      Width           =   195
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
      Height          =   330
      Index           =   9
      Left            =   6660
      MaxLength       =   3
      TabIndex        =   6
      Tag             =   "Tipo Iva Exento|N|N|||rseccion|codivaexe|000||"
      Text            =   "Ti"
      Top             =   3900
      Width           =   675
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
      Height          =   330
      Index           =   5
      Left            =   5760
      MaxLength       =   10
      TabIndex        =   5
      Tag             =   "Raiz Proveedor|T|S|||rseccion|raiz_proveedor|||"
      Text            =   "Rai.Prove"
      Top             =   3915
      Width           =   675
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   290
      Index           =   2
      Left            =   6435
      MaskColor       =   &H00000000&
      TabIndex        =   22
      ToolTipText     =   "Buscar Cta.Contable"
      Top             =   3915
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   330
      Index           =   1
      Left            =   5535
      MaskColor       =   &H00000000&
      TabIndex        =   17
      ToolTipText     =   "Buscar Cta.Contable"
      Top             =   3915
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Frame Frame2 
      Height          =   1755
      Left            =   105
      TabIndex        =   14
      Top             =   4740
      Width           =   9450
      Begin VB.TextBox txtAux 
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
         Index           =   10
         Left            =   3030
         MaxLength       =   30
         TabIndex        =   24
         Top             =   1320
         Width           =   6255
      End
      Begin VB.TextBox txtAux 
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
         Index           =   8
         Left            =   3030
         MaxLength       =   30
         TabIndex        =   20
         Top             =   930
         Width           =   6255
      End
      Begin VB.TextBox txtAux 
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
         Left            =   3030
         MaxLength       =   30
         TabIndex        =   18
         Top             =   180
         Width           =   6255
      End
      Begin VB.TextBox txtAux 
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
         Left            =   3030
         MaxLength       =   30
         TabIndex        =   15
         Top             =   555
         Width           =   6255
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de Iva Exento"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   25
         Top             =   1290
         Width           =   2670
      End
      Begin VB.Label Label3 
         Caption         =   "Raíz Cta.Contable Proveedor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   165
         TabIndex        =   21
         Top             =   930
         Width           =   3150
      End
      Begin VB.Label Label2 
         Caption         =   "Raíz Cta.Contable Asociado"
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
         Left            =   165
         TabIndex        =   19
         Top             =   570
         Width           =   2910
      End
      Begin VB.Label Label1 
         Caption         =   "Raíz Cta.Contable Socio"
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
         Left            =   165
         TabIndex        =   16
         Top             =   225
         Width           =   2670
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
      Height          =   330
      Index           =   3
      Left            =   3780
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "Raiz Cliente-Socio|T|S|||rseccion|raiz_cliente_socio|||"
      Text            =   "Rai.Cli-So"
      Top             =   3915
      Width           =   900
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   330
      Index           =   0
      Left            =   4635
      MaskColor       =   &H00000000&
      TabIndex        =   13
      ToolTipText     =   "Buscar Cta.Contable"
      Top             =   3915
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
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
      Height          =   330
      Index           =   0
      Left            =   135
      MaxLength       =   3
      TabIndex        =   0
      Tag             =   "Codigo|N|N|1|999|rseccion|codsecci|000|S|"
      Text            =   "Cod"
      Top             =   3930
      Width           =   585
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
      Height          =   330
      Index           =   1
      Left            =   810
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Nombre|T|N|||rseccion|nomsecci|||"
      Text            =   "Nombre"
      Top             =   3930
      Width           =   2325
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
      Height          =   330
      Index           =   4
      Left            =   4860
      MaxLength       =   10
      TabIndex        =   4
      Tag             =   "Raiz Cliente-Asoc|T|S|||rseccion|raiz_cliente_asociado|||"
      Text            =   "Rai.Cli-As"
      Top             =   3915
      Width           =   675
   End
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
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
      Height          =   330
      Index           =   2
      Left            =   3150
      MaxLength       =   2
      TabIndex        =   2
      Tag             =   "N.Conta|N|N|1|99|rseccion|empresa_conta|00||"
      Text            =   "C"
      Top             =   3915
      Width           =   585
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
      Left            =   7335
      TabIndex        =   7
      Top             =   6780
      Visible         =   0   'False
      Width           =   1035
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
      Left            =   8460
      TabIndex        =   8
      Top             =   6780
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmManSeccion.frx":000C
      Height          =   3820
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   6747
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
      Left            =   8460
      TabIndex        =   12
      Top             =   6780
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   6570
      Width           =   2385
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
         Left            =   45
         TabIndex        =   10
         Top             =   210
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   330
      Left            =   4410
      Top             =   45
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
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   9090
      TabIndex        =   28
      Top             =   150
      Width           =   405
      _ExtentX        =   714
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
         HelpContextID   =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmManSeccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MANOLO (Se lo copia)                          +-+-
' +-+- Menú: Bancos Propios (con un par)                    +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-

' +-+-+-+- DISSENY +-+-+-+-
' 1. Posar tots els controls al formulari
' 2. Posar els index correlativament
' 3. Si n'hi han botons de buscar repasar el ToolTipText
' 4. Alliniar els camps numérics a la dreta i el resto a l'esquerra
' 5. Posar els TAGs
' (si es INTEGER: si PK => mínim 1; si no PK => mínim 0; màxim => 99; format => 00)
' (si es DECIMAL; mínim => 0; màxim => 99.99; format => #,###,###,##0.00)
' (si es DATE; format => dd/mm/yyyy)
' 6. Posar els MAXLENGTHs
' 7. Posar els TABINDEXs

Option Explicit

Private Const IdPrograma = 2001

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'altre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String

Public DeConsulta As Boolean
Public CodigoActual As String

' *** adrede: per a quan busque suplements/desconters des de frmViagrc ***
Public ExpedBusca As Long
Public TipoSuplem As Integer
' *********************************************************************

' *** declarar els formularis als que vaig a cridar ***
'Private WithEvents frmB As frmBuscaGrid

Private CadenaConsulta As String
Private CadB As String

' ### [Monica] 08/09/2006
Private WithEvents frmCtas As frmCtasConta 'cuentas contables
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmTIva As frmTipIVAConta 'tipos de via de contabilidad
Attribute frmTIva.VB_VarHelpID = -1
Dim Indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos

Private kCampo As Integer

Dim vSeccion As CSeccion

Dim Modo As Byte
'----------- MODOS --------------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'--------------------------------------------------

Private Sub PonerModo(vModo)
Dim B As Boolean
Dim I As Integer
    
    Modo = vModo
'    PonerIndicador lblIndicador, Modo
    
    B = (Modo = 2)
    If B Then
        PonerContRegIndicador lblIndicador, adodc1, CadB
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    ' **** posar tots els controls (botons inclosos) que siguen del Grid
    For I = 0 To 5
        txtAux(I).visible = Not B
    Next I
    txtAux(9).visible = Not B
    For I = 0 To btnBuscar.Count - 1
        btnBuscar(I).visible = (Modo = 3 Or Modo = 4)
        btnBuscar(I).Enabled = (Modo = 3 Or Modo = 4)
    Next I
    ' **************************************************
    
    ' **** si n'hi han camps fora del grid, bloquejar-los ****
    For I = 6 To 8
        BloquearTxt txtAux(I), True
    Next I
    BloquearTxt txtAux(10), True
    ' ********************************************************

    cmdAceptar.visible = Not B
    cmdCancelar.visible = Not B
    DataGrid1.Enabled = B
    
    'Si es retornar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = B
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botons de menu según Modo
    PonerOpcionesMenu 'Activar/Desact botons de menu según permissos de l'usuari
    
    ' *** bloquejar tota la PK quan estem en modificar  ***
    BloquearTxt txtAux(0), (Modo = 4) 'codconce
    
    BloquearImgBuscar Me, Modo

End Sub

Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botons de la toolbar i del menu, según el modo en que estiguem
Dim B As Boolean

    ' *** adrede: per a que no es puga fer res si estic cridant des de frmViagrc ***

    B = (Modo = 2) And ExpedBusca = 0
    'Busqueda
    Toolbar1.Buttons(5).Enabled = B
    Me.mnBuscar.Enabled = B
    'Vore Tots
    Toolbar1.Buttons(6).Enabled = B
    Me.mnVerTodos.Enabled = B
    
    'Insertar
    Toolbar1.Buttons(1).Enabled = B And Not DeConsulta
    Me.mnNuevo.Enabled = B And Not DeConsulta
    
    B = (B And adodc1.Recordset.RecordCount > 0) And Not DeConsulta And ExpedBusca = 0
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    Me.mnModificar.Enabled = B

    'Eliminar
    Toolbar1.Buttons(3).Enabled = B
    Me.mnEliminar.Enabled = B
    'Imprimir
    Toolbar1.Buttons(8).Enabled = B
    Me.mnImprimir.Enabled = B

    ' ******************************************************************************
End Sub

Private Sub BotonAnyadir()
Dim NumF As String
Dim anc As Single
Dim I As Integer
    
    CargaGrid 'primer de tot carregue tot el grid
    CadB = ""
    '********* canviar taula i camp; repasar codEmpre ************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = SugerirCodigoSiguienteStr("rseccion", "codsecci")
        'NumF = SugerirCodigoSiguienteStr("sdexpgrp", "codsupdt", "codempre=" & vSesion.Empresa)
        'NumF = ""
    End If
    '***************************************************************
    'Situem el grid al final
    AnyadirLinea DataGrid1, adodc1

    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 206
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    
    ' *** valors per defecte a l'afegir (dins i fora del grid); repasar codEmpre ***
    txtAux(0).Text = NumF
    For I = 1 To 10
        txtAux(I).Text = ""
    Next I
    ' **************************************************

    LLamaLineas anc, 3
       
    ' *** posar el foco ***
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        PonerFoco txtAux(1) '**** 1r camp visible que NO siga PK ****
    Else
        PonerFoco txtAux(0) '**** 1r camp visible que siga PK ****
    End If
    ' ******************************************************
End Sub

Private Sub BotonVerTodos()
    CargaGrid ""
    CadB = ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    Dim I As Integer
    
    ' *** canviar per la PK (no posar codempre si està a Form_Load) ***
    CargaGrid "codsecci = -1"
    '*******************************************************************************

    ' *** canviar-ho pels valors per defecte al buscar (dins i fora del grid);
    For I = 0 To 9
        txtAux(I).Text = ""
    Next I

    LLamaLineas DataGrid1.Top + 240, 1
    
    ' *** posar el foco al 1r camp visible que siga PK ***
    PonerFoco txtAux(0)
    ' ***************************************************************
End Sub

Private Sub BotonModificar()
    Dim anc As Single
    Dim I As Integer

    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top  '545
    End If

    ' *** asignar als controls del grid, els valors de les columnes ***
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux(2).Text = ComprobarCero(Trim(DataGrid1.Columns(2).Text))
    txtAux(3).Text = DataGrid1.Columns(3).Text
    txtAux(4).Text = DataGrid1.Columns(4).Text
    txtAux(5).Text = DataGrid1.Columns(5).Text
    txtAux(9).Text = DataGrid1.Columns(6).Text
    
    ' ********************************************************

    LLamaLineas anc, 4 'modo 4
   
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco txtAux(1)
    ' *********************************************************
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
Dim I As Integer

    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo

    ' *** posar el Top a tots els controls del grid (botons també) ***
    'Me.imgFec(2).Top = alto
    For I = 0 To 5
        txtAux(I).Top = alto
    Next I
    txtAux(9).Top = alto
    
    btnBuscar(0).Top = alto - 10
    btnBuscar(1).Top = alto - 10
    btnBuscar(2).Top = alto - 10
    btnBuscar(3).Top = alto - 10
    ' ***************************************************
End Sub

Private Sub BotonEliminar()
Dim SQL As String
Dim temp As Boolean

    On Error GoTo Error2
    
    'Certes comprovacions
    If adodc1.Recordset.EOF Then Exit Sub
'    If Not SepuedeBorrar Then Exit Sub
    
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(adodc1.Recordset.Fields(1).Value), FormatoCampo(txtAux(1))) Then Exit Sub
    ' ***************************************************************************
    
    '*** canviar la pregunta, els noms dels camps i el DELETE; repasar codEmpre ***
    SQL = "¿Seguro que desea eliminar la Sección?"
    'SQL = SQL & vbCrLf & "Código: " & Format(adodc1.Recordset.Fields(0), "000")
    SQL = SQL & vbCrLf & "Código: " & adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Nombre: " & adodc1.Recordset.Fields(1)
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'N'hi ha que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        SQL = "Delete from rseccion where codsecci = " & adodc1.Recordset!codsecci
        
        conn.Execute SQL
        CargaGrid CadB
'        If CadB <> "" Then
'            CargaGrid CadB
'            lblIndicador.Caption = "RESULTADO BUSQUEDA"
'        Else
'            CargaGrid ""
'            lblIndicador.Caption = ""
'        End If
        temp = SituarDataTrasEliminar(adodc1, NumRegElim, True)
        PonerModoOpcionesMenu
        adodc1.Recordset.Cancel
    End If
    Exit Sub
    
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los txtAux
    PonerLongCamposGnral Me, Modo, 3
End Sub

Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
     Select Case Index
        Case 0, 1, 2 'Cuentas Contables (de contabilidad)
            If txtAux(2).Text = "" Then Exit Sub
            
            If CInt(txtAux(2).Text) = 0 Then Exit Sub
            
            If AbrirConexionConta2(txtAux(2).Text) Then
                Set vEmpresa = New Cempresa
                If vEmpresa.LeerNiveles Then
                    Select Case Index
                        Case 0
                            Indice = 3
                        Case 1
                            Indice = 4
                        Case 2
                            Indice = 5
                    End Select
                    Set frmCtas = New frmCtasConta
'                    frmCtas.Conexion = txtAux(2).Text
'                    frmCtas.CadBusqueda = DevuelveDesdeBDNew(cConta, "parametros", "grupovta", "", "", "T")
'                    frmCtas.Facturas = True
                    frmCtas.NumDigit = vEmpresa.DigitosNivelAnterior
                    frmCtas.DatosADevolverBusqueda = "0|1|"
                    frmCtas.CodigoActual = txtAux(Indice).Text
                    frmCtas.Show vbModal
                    Set frmCtas = Nothing
                    PonerFoco txtAux(Indice)
                End If
                Set vEmpresa = Nothing
                CerrarConexionConta2
            End If
            
       Case 3 'Tipo de iva exento
            If txtAux(2).Text = "" Then Exit Sub
            
            If CInt(txtAux(2).Text) = 0 Then Exit Sub
            
            If AbrirConexionConta2(txtAux(2).Text) Then
                Set vEmpresa = New Cempresa
                If vEmpresa.LeerNiveles Then
                    Set frmTIva = New frmTipIVAConta
                    frmTIva.DeConsulta = True
                    frmTIva.DatosADevolverBusqueda = "0|1|"
                    frmTIva.CodigoActual = txtAux(9).Text
                    frmTIva.Show vbModal
                    Set frmTIva = Nothing
                    PonerFoco txtAux(9)
                End If
                Set vEmpresa = Nothing
                CerrarConexionConta2
            End If
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1

End Sub

Private Sub cmdAceptar_Click()
Dim I As Long

    Select Case Modo
        Case 3 'INSERTAR
            If DatosOK Then
                'If InsertarDesdeForm(Me) Then
                If InsertarDesdeForm2(Me, 0) Then
                    CargaGrid
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                        cmdCancelar_Click
                        If Not adodc1.Recordset.EOF Then
                            ' *** filtrar per tota la PK; repasar codEmpre **
                            'adodc1.Recordset.Filter = "codempre = " & txtAux(0).Text & " AND codsupdt = " & txtAux(1).Text
                            adodc1.Recordset.Filter = "codsecci = " & txtAux(0).Text
                            ' ****************************************************
                        End If
                        cmdRegresar_Click
                    Else
                        BotonAnyadir
                    End If
                    CadB = ""
                End If
            End If
            
        Case 4 'MODIFICAR
            If DatosOK Then
                'If ModificaDesdeFormulario(Me) Then
                If ModificaDesdeFormulario2(Me, 0) Then
                    I = adodc1.Recordset.AbsolutePosition
                    TerminaBloquear
                    PonerModo 2
                    CargaGrid CadB
'                    If CadB <> "" Then
'                        CargaGrid CadB
'                        lblIndicador.Caption = "RESULTADO BUSQUEDA"
'                    Else
'                        CargaGrid
'                        lblIndicador.Caption = ""
'                    End If
                    adodc1.Recordset.Move I - 1
                    PonerFocoGrid Me.DataGrid1
                End If
            End If
            
        Case 1  'BUSQUEDA
            CadB = ObtenerBusqueda(Me)
            If CadB <> "" Then
                CargaGrid CadB
                PonerModo 2
'                lblIndicador.Caption = "RESULTADO BUSQUEDA"
                PonerFocoGrid Me.DataGrid1
            End If
    End Select
End Sub

Private Sub cmdCancelar_Click()
'On Error Resume Next

    Select Case Modo
        Case 3 'INSERTAR
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        Case 4 'MODIFICAR
            TerminaBloquear
        Case 1 'BUSQUEDA
            CargaGrid CadB
    End Select
    
    If Not adodc1.Recordset.EOF Then
        CargaForaGrid
    Else
        LimpiarCampos
    End If
    
    PonerModo 2
'    If CadB <> "" Then
'        lblIndicador.Caption = "RESULTADO BUSQUEDA"
'    Else
'        lblIndicador.Caption = ""
'    End If
    PonerFocoGrid Me.DataGrid1
'    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim I As Integer
Dim J As Integer
Dim Aux As String

    If adodc1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    cad = ""
    I = 0
    Do
        J = I + 1
        I = InStr(J, DatosADevolverBusqueda, "|")
        If I > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, I - J)
            J = Val(Aux)
            cad = cad & adodc1.Recordset.Fields(J) & "|"
        End If
    Loop Until I = 0
    ' *** adrede: per a tornar el TipoSuplem ***
    ' cad = cad & TipoSuplem & "|"
    ' ******************************************
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    'Posem el foco
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        PonerFoco txtAux(1)
    End If
End Sub

Private Sub Form_Load()
Dim I As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    '******* repasar si n'hi ha botó d'imprimir o no******
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'el 1 es separadors
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2   'Tots
        'el 4 i el 5 son separadors
        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
        'el 9 i el 10 son separadors
        .Buttons(8).Image = 10  'Imprimir
    End With
    '*****************************************************
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 12
    End With


'    chkVistaPrevia.Value = CheckValueLeer(Name)
    ' *** SI N'HI HAN COMBOS ***
    ' CargaCombo 0
    ' **************************
    
    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT codsecci, nomsecci, empresa_conta, raiz_cliente_socio, raiz_cliente_asociado, raiz_proveedor, codivaexe "
    CadenaConsulta = CadenaConsulta & " FROM rseccion "
    '************************************************************************
    
    CadB = ""
    CargaGrid
    
    ' ****** Si n'hi han camps fora del grid ******
    CargaForaGrid
    ' *********************************************
    
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        BotonAnyadir
    Else
        PonerModo 2
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
    If Modo = 4 Then TerminaBloquear
End Sub

' ### [Monica] 08/09/2006
Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    txtAux(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    txtAux(Indice + 3).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta
End Sub

Private Sub frmTIva_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(9).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtAux(10).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    printNou
End Sub

Private Sub mnModificar_Click()
    'Comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(adodc1.Recordset.Fields(1).Value), FormatoCampo(txtAux(1))) Then Exit Sub
    ' ***************************************************************************
    
    
    'Prepara para modificar
    '-----------------------
    If BLOQUEADesdeFormulario2(Me, adodc1, 1) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
                mnNuevo_Click
        Case 2
                mnModificar_Click
        Case 3
                mnEliminar_Click
        Case 5
                mnBuscar_Click
        Case 6
                mnVerTodos_Click
        Case 8 'Imprimir
                mnImprimir_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String)
    Dim I As Integer
    Dim SQL As String
    Dim tots As String
    
'    adodc1.ConnectionString = Conn
    ' *** si en Form_load ya li he posat clausula WHERE, canviar el `WHERE` de
    ' `SQL = CadenaConsulta & " WHERE " & vSQL` per un `AND`
    If vSQL <> "" Then
        SQL = CadenaConsulta & " WHERE " & vSQL  ' ### [Monica] 08/09/2006: antes habia AND
    Else
        SQL = CadenaConsulta
    End If
    '********************* canviar el ORDER BY *********************++
    'SQL = SQL & " ORDER BY codempre, codsupdt"
    SQL = SQL & " ORDER BY codsecci"
    '**************************************************************++
    
'    adodc1.RecordSource = SQL
'    adodc1.CursorType = adOpenDynamic
'    adodc1.LockType = adLockOptimistic
'    DataGrid1.ScrollBars = dbgNone
'    adodc1.Refresh
'    Set DataGrid1.DataSource = adodc1 ' per a que no ixca l'error de "la fila actual no está disponible"
       
    CargaGridGnral Me.DataGrid1, Me.adodc1, SQL, False
       
       
    ' *** posar només els controls del grid ***
    tots = "S|txtAux(0)|T|Cód.|550|;S|txtAux(1)|T|Denominación|3200|;S|txtAux(2)|T|Conta|700|;"
    tots = tots & "S|txtAux(3)|T|Raíz Socio|1200|;S|btnBuscar(0)|B|||;"
    tots = tots & "S|txtAux(4)|T|Raíz Asoc.|1200|;S|btnBuscar(1)|B|||;"
    tots = tots & "S|txtAux(5)|T|Raíz Prov.|1200|;S|btnBuscar(2)|B|||;"
    tots = tots & "S|txtAux(9)|T|Iva Ex.|800|;S|btnBuscar(3)|B|||;"
    
'    For i = 1 To 11
'        tots = tots & "N||||0|;"
'    Next i
    arregla tots, DataGrid1, Me, 350
    DataGrid1.ScrollBars = dbgAutomatic
    ' **********************************************************
    
    ' *** alliniar les columnes que siguen numèriques a la dreta ***
    DataGrid1.Columns(2).Alignment = dbgCenter
    DataGrid1.Columns(3).Alignment = dbgLeft
    DataGrid1.Columns(4).Alignment = dbgLeft
    DataGrid1.Columns(5).Alignment = dbgLeft
    DataGrid1.Columns(6).Alignment = dbgLeft
    ' *****************************
    
    
    ' *** Si n'hi han camps fora del grid ***
    If Not adodc1.Recordset.EOF Then
        CargaForaGrid
    Else
        LimpiarCampos
    End If
    ' **************************************
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
'    If Index = 3 And KeyAscii = 43 Then '+
'        KeyAscii = 0
'    Else
'        KEYpress KeyAscii
'    End If
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 12: KEYBusqueda KeyAscii, 0 'cuenta contable
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
    
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
'    imgBuscar_Click (indice)
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    '*** configurar el LostFocus dels camps (de dins i de fora del grid) ***
    Select Case Index
        Case 0
            PonerFormatoEntero txtAux(Index)
        
        Case 1
            If txtAux(Index).Text = "" Then Exit Sub
            txtAux(Index).Text = UCase(txtAux(Index).Text)
            
        Case 2 'base de datos de contabilidad
            PonerFormatoEntero txtAux(Index)
            If txtAux(Index).Text <> "" Then
                If Not AbrirConexionConta2(txtAux(Index).Text) Then
                     txtAux(Index).Text = ""
                     PonerFoco txtAux(Index)
                Else
                    CerrarConexionConta2
                End If
            End If
        
        Case 3 'cuenta contable
            If txtAux(Index).Text = "" Then Exit Sub
            If txtAux(2).Text = "" Then Exit Sub
            
            If AbrirConexionConta2(txtAux(2).Text) Then
                txtAux(6) = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", txtAux(3).Text, "T")
                Set vEmpresa = New Cempresa
                If vEmpresa.LeerNiveles Then
                    If Len(txtAux(Index)) <> vEmpresa.DigitosNivelAnterior Then
                        MsgBox "Longitud de cuenta incorrecta. Revise.", vbExclamation
                    End If
                End If
                Set vEmpresa = Nothing
                CerrarConexionConta2
            End If
        Case 4 'cuenta contable
            If txtAux(Index).Text = "" Then Exit Sub
            If txtAux(2).Text = "" Then Exit Sub
            
            If AbrirConexionConta2(txtAux(2).Text) Then
                txtAux(7) = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", txtAux(4).Text, "T")
                Set vEmpresa = New Cempresa
                If vEmpresa.LeerNiveles Then
                    If Len(txtAux(Index)) <> vEmpresa.DigitosNivelAnterior Then
                        MsgBox "Longitud de cuenta incorrecta. Revise.", vbExclamation
                    End If
                End If
                Set vEmpresa = Nothing
                CerrarConexionConta2
            End If
        Case 5 'cuenta contable
            If txtAux(Index).Text = "" Then Exit Sub
            If txtAux(2).Text = "" Then Exit Sub
            
            If AbrirConexionConta2(txtAux(2).Text) Then
                txtAux(8) = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", txtAux(5).Text, "T")
                Set vEmpresa = New Cempresa
                If vEmpresa.LeerNiveles Then
                    If Len(txtAux(Index)) <> vEmpresa.DigitosNivelAnterior Then
                        MsgBox "Longitud de cuenta incorrecta. Revise.", vbExclamation
                    End If
                End If
                Set vEmpresa = Nothing
                CerrarConexionConta2
            End If
        
        Case 9 ' tipo de iva
            If txtAux(Index).Text <> "" Then
                If AbrirConexionConta2(txtAux(2).Text) Then
                    txtAux(10).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", txtAux(9).Text, "N")
                    CerrarConexionConta2
                End If
            End If
        
    End Select
    '**************************************************************************
End Sub


Private Function DatosOK() As Boolean
Dim Datos As String
Dim B As Boolean
' *** només per ad este manteniment ***
Dim Rs As Recordset
Dim cad As String
'Dim exped As String
' *************************************

    B = CompForm(Me)
    If Not B Then Exit Function


    If B And (Modo = 3) Then
        'Estem insertant
        'aço es com posar: select codvarie from svarie where codvarie = txtAux(0)
        'la N es pa dir que es numèric
         
        ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
        Datos = DevuelveDesdeBD("codsecci", "rseccion", "codsecci", txtAux(0).Text, "N")
'       Datos = DevuelveDesdeBDNew(1, "sdexpgrp", "codsupdt", "codsupdt", txtAux(1).Text, "N", "", "codempre", CStr(vSesion.Empresa), "N")
         
        If Datos <> "" Then
            MsgBox "Ya existe el Código de Sección: " & txtAux(0).Text, vbExclamation
            B = False
            PonerFoco txtAux(1) '*** posar el foco al 1r camp visible de la PK de la capçalera ***
            Exit Function
        End If
        '*************************************************************************************
    End If

    ' *** Si cal fer atres comprovacions ***
    ' comprobamos que las raices de las cuentas contables tengan la longitud que toque
    If B And (Modo = 3 Or Modo = 4) Then
        If AbrirConexionConta2(txtAux(2)) Then
            Set vEmpresa = New Cempresa
            If vEmpresa.LeerNiveles Then
                If Len(txtAux(3)) <> vEmpresa.DigitosNivelAnterior Then
                    MsgBox "Longitud de cuenta Socio incorrecta. Revise.", vbExclamation
                    B = False
                End If
                If B And Len(txtAux(4)) <> vEmpresa.DigitosNivelAnterior Then
                    MsgBox "Longitud de cuenta Asociado incorrecta. Revise.", vbExclamation
                    B = False
                End If
                If B And Len(txtAux(5)) <> vEmpresa.DigitosNivelAnterior Then
                    MsgBox "Longitud de cuenta Proveedor incorrecta. Revise.", vbExclamation
                    B = False
                End If
            End If
            Set vEmpresa = Nothing
            CerrarConexionConta2
        End If
    End If
    ' *********************************************

    DatosOK = B
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me

End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim I As Byte

    If Modo <> 4 Then 'Modificar
        CargaForaGrid
    Else
        For I = 0 To txtAux.Count - 1
            txtAux(I).Text = ""
        Next I
    End If
    
    If Modo = 2 Then PonerContRegIndicador lblIndicador, adodc1, CadB
    
End Sub

Private Sub CargaForaGrid()
        If adodc1.Recordset.EOF Then Exit Sub
        
        If IsNull(Me.adodc1.Recordset.Fields(2).Value) Then Exit Sub
        
        
        If AbrirConexionConta2(Me.adodc1.Recordset.Fields(2).Value) Then
            txtAux(6).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Me.adodc1.Recordset.Fields(3).Value, "T")
            txtAux(7).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Me.adodc1.Recordset.Fields(4).Value, "T")
            txtAux(8).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Me.adodc1.Recordset.Fields(5).Value, "T")
            txtAux(10).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Me.adodc1.Recordset.Fields(6).Value, "N")
            CerrarConexionConta2
        End If
        ' *** posar als camps de fora del grid el valor de la columna corresponent ***
        
'        txtAux(6) = PonerNombreCuenta(txtAux(4), Modo, cContaFac)
        ' *** Si fora del grid n'hi han camps de descripció, posar-los valor ***
'        text2(12).Text = PonerNombreCuenta(txtAux(12), Modo)
        
        'txtAux2(4).Text = PonerNombreDeCod(txtAux(4), "poblacio", "despobla", "codpobla", "N")
'       If txtAux(4).Text <> "" Then _
'           txtAux2(4).Text = DevuelveDesdeBDNew(1, "supdtogr", "nomsuple", "codsuple", txtAux(4).Text, "N", "", "codempre", CStr(vSesion.Empresa), "N")
        ' **********************************************************************
 End Sub

Private Sub LimpiarCampos()
Dim I As Integer
On Error Resume Next

    ' *** posar a huit tots els camps de fora del grid ***
    For I = 5 To 7
        txtAux(I).Text = ""
    Next I
    ' ****************************************************
'    text2(12).Text = "" ' el nombre de la cuenta contable la ponemos a cero

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "rseccion"
        .Informe2 = "rManSeccion.rpt"
        If CadB <> "" Then
            '.cadRegSelec = Replace(SQL2SF(CadB), "clientes", "clientes_1")
            .cadRegSelec = SQL2SF(CadB)
        Else
            .cadRegSelec = ""
        End If
        ' *** repasar el nom de l'adodc ***
        '.cadRegActua = Replace(POS2SF(Data1, Me), "clientes", "clientes_1")
        .cadRegActua = POS2SF(adodc1, Me)
        '[Monica]13/07/2012: falla si hay un solo registro seleccionado y apretamos registros buscados
        If adodc1.Recordset.RecordCount = 1 Then .cadRegSelec = .cadRegActua
        
        ' *** repasar codEmpre ***
        .cadTodosReg = ""
        '.cadTodosReg = "{sbanco.codbanpr} = " & vSesion.Empresa
        ' *** repasar si li pose ordre o no ****
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|pOrden={rseccion.codsecci}|"
        '.OtrosParametros2 = "pEmpresa='" & vEmpresa.NomEmpre & "'|"
        ' *** posar el nº de paràmetres que he posat en OtrosParametros2 ***
        '.NumeroParametros2 = 1
        .NumeroParametros2 = 2
        ' ******************************************************************
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = False

        .Show vbModal
    End With
End Sub

'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGrid1_GotFocus()
'  WheelHook DataGrid1
'End Sub
'Private Sub DataGrid1_Lostfocus()
'  WheelUnHook
'End Sub

