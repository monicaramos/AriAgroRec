VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPOZMantaTickets 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tickets de Riego a Manta"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15735
   Icon            =   "frmPOZMantaTickets.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   15735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   3840
      MaxLength       =   20
      TabIndex        =   29
      Tag             =   "Concepto|T|N|||rpozticketsmanta|concepto|||"
      Top             =   5340
      Width           =   5475
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   12
      Left            =   14250
      MaxLength       =   10
      TabIndex        =   12
      Tag             =   "Fecha Pago|F|S|||rpozticketsmanta|fecpago|dd/mm/yyyy||"
      Top             =   3840
      Width           =   885
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   4
      Left            =   15180
      MaskColor       =   &H00000000&
      TabIndex        =   28
      ToolTipText     =   "Buscar fecha"
      Top             =   3840
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   3
      Left            =   14040
      MaskColor       =   &H00000000&
      TabIndex        =   27
      ToolTipText     =   "Buscar fecha"
      Top             =   3840
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   11
      Left            =   11400
      MaxLength       =   10
      TabIndex        =   9
      Tag             =   "Precio|N|N|||rpozticketsmanta|precio1|#,##0.0000||"
      Top             =   3840
      Width           =   885
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   -60
      MaxLength       =   7
      TabIndex        =   0
      Tag             =   "Nro Albaran|N|S|0|9999999|rpozticketsmanta|numalbar|0000000|S|"
      Text            =   "Albaran"
      Top             =   3780
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   10
      Left            =   810
      MaxLength       =   10
      TabIndex        =   1
      Tag             =   "Fecha Albaran|F|N|||rpozticketsmanta|fecalbar|dd/mm/yyyy|S|"
      Text            =   "Fecha"
      Top             =   3780
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   9
      Left            =   13140
      MaxLength       =   10
      TabIndex        =   11
      Tag             =   "Fecha Riego|F|S|||rpozticketsmanta|fecriego|dd/mm/yyyy||"
      Top             =   3840
      Width           =   885
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   8
      Left            =   12270
      MaxLength       =   10
      TabIndex        =   10
      Tag             =   "Importe|N|N|||rpozticketsmanta|importe|###,##0.00||"
      Top             =   3840
      Width           =   885
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   6750
      MaxLength       =   40
      TabIndex        =   26
      Top             =   3840
      Width           =   1485
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   7
      Left            =   10320
      MaxLength       =   8
      TabIndex        =   8
      Tag             =   "Hanegadas|N|N|||rpozticketsmanta|hanegada|#,##0.00||"
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   6
      Left            =   9630
      MaxLength       =   2
      TabIndex        =   7
      Tag             =   "Subparcela|T|S|||rpozticketsmanta|subparce||N|"
      Top             =   3840
      Width           =   675
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   0
      Left            =   3570
      MaskColor       =   &H00000000&
      TabIndex        =   24
      ToolTipText     =   "Buscar socio"
      Top             =   3780
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   1
      Left            =   2250
      MaskColor       =   &H00000000&
      TabIndex        =   23
      ToolTipText     =   "Buscar partida"
      Top             =   3780
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   2
      Left            =   6480
      MaskColor       =   &H00000000&
      TabIndex        =   22
      ToolTipText     =   "Buscar fecha"
      Top             =   3840
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   3
      Left            =   5700
      MaxLength       =   4
      TabIndex        =   4
      Tag             =   "Braçal|N|N|||rpozticketsmanta|codzonas|0000||"
      Text            =   "1234567"
      Top             =   3840
      Width           =   765
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   4
      Left            =   8280
      MaxLength       =   3
      TabIndex        =   5
      Tag             =   "Poligonol|N|S|||rpozticketsmanta|poligono|000||"
      Text            =   "1234567890"
      Top             =   3840
      Width           =   555
   End
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   5
      Left            =   8850
      MaxLength       =   7
      TabIndex        =   6
      Tag             =   "Parcela|N|S|||rpozticketsmanta|parcela|000000||"
      Text            =   "1234567"
      Top             =   3840
      Width           =   795
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   1
      Left            =   4950
      MaxLength       =   8
      TabIndex        =   3
      Tag             =   "Campo|N|N|1|99999999|rpozticketsmanta|codcampo|00000000|S|"
      Top             =   3810
      Width           =   705
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   3810
      MaxLength       =   30
      TabIndex        =   21
      Top             =   3810
      Width           =   1125
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   0
      Left            =   2490
      MaxLength       =   6
      TabIndex        =   2
      Tag             =   "Socio|N|N|||rpozticketsmanta|codsocio|000000|S|"
      Top             =   3780
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   13350
      TabIndex        =   13
      Tag             =   "   "
      Top             =   5280
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   14490
      TabIndex        =   14
      Top             =   5265
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4410
      Left            =   120
      TabIndex        =   17
      Top             =   675
      Width           =   15435
      _ExtentX        =   27226
      _ExtentY        =   7779
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
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
      Height          =   375
      Left            =   14490
      TabIndex        =   20
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   5190
      Width           =   2385
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
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
         Left            =   40
         TabIndex        =   16
         Top             =   240
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   2205
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   15735
      _ExtentX        =   27755
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar Facturados"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Impresión Ticket"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pago Ticket"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   3735
         TabIndex        =   19
         Top             =   90
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Concepto:"
      Height          =   285
      Left            =   3060
      TabIndex        =   30
      Top             =   5370
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Procesando Registro:"
      Height          =   225
      Left            =   9540
      TabIndex        =   25
      Top             =   5370
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBuscarFacturados 
         Caption         =   "&Buscar Facturados"
         Shortcut        =   ^R
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
         Caption         =   "&Impresion Ticket"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnPagoTicket 
         Caption         =   "&Pago Ticket"
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
Attribute VB_Name = "frmPOZMantaTickets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MANOLO  +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-

' **************** PER A QUE FUNCIONE EN UN ATRE MANTENIMENT ********************
' 0. Posar-li l'atribut Datasource a "adodc1" del Datagrid1. Canviar el Caption
'    del formulari
' 1. Canviar els TAGs i els Maxlength de TextAux(0) i TextAux(1)
' 2. En PonerModo(vModo) repasar els indexs del botons, per si es canvien
' 3. En la funció BotonAnyadir() canviar la taula i el camp per a SugerirCodigoSiguienteStr
' 4. En la funció BotonBuscar() canviar el nom de la clau primaria
' 5. En la funció BotonEliminar() canviar la pregunta, les descripcions de la
'    variable SQL i el contingut del DELETE
' 6. En la funció PonerLongCampos() posar els camps als que volem canviar el MaxLength quan busquem
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
' *******************************SI N'HI HA COMBO*******************************
' 0. Comprovar que en el SQL de Form_Load() es faça referència a la taula del Combo
' 1. Pegar el Combo1 al  costat dels TextAux. Canviar-li el TAG
' 2. En BotonModificar() canviar el camp del Combo
' 3. En CargaCombo() canviar la consulta i els noms del camps, o posar els valor
'    a ma si no es llig de cap base de datos els valors del Combo

Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmZon As frmManZonas 'zonas
Attribute frmZon.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes
Attribute frmMens.VB_VarHelpID = -1


'codi per al registe que s'afegix al cridar des d'atre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

Private CadenaConsulta As String
Private CadB As String

Dim Ordenacion As String

Dim Modo As Byte
'----------- MODOS --------------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'--------------------------------------------------
Dim PrimeraVez As Boolean
Dim Indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim I As Integer

Dim FechaAnt As String
Dim OK As Boolean
Dim CadB1 As String
Dim Filtro As Byte
Dim SQL As String

Dim CodTipoMov As String

Dim PagoTicket As Boolean

Dim CadB2 As String

Dim SoloFacturado As Boolean
Dim CadenaB As String

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe
Private Titulo As String 'Titulo informe que se pasa a frmImprimir
Private nomRPT As String 'nombre del fichero .rpt a imprimir
Private conSubRPT As Boolean 'si tiene subinformes para enlazarlos a las tablas correctas
'-------------------------------------

Dim Continuar As Boolean
Dim EsTicketContado As Boolean

Private Sub PonerModo(vModo)
Dim B As Boolean

    Modo = vModo
    
    B = (Modo = 2)
    If B Then
        PonerContRegIndicador lblIndicador, adodc1, CadB
        If Not adodc1.Recordset.EOF Then Text2(1).Text = DBLet(adodc1.Recordset!Concepto, "T")
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    For I = 0 To txtAux.Count - 1
        txtAux(I).visible = (Modo = 1)
        txtAux(I).Enabled = (Modo = 1)
    Next I
    
    'hdas
    txtAux(7).visible = (Modo = 1 Or (Modo = 4 And Not PagoTicket And Not SoloFacturado))
    txtAux(7).Enabled = (Modo = 1 Or (Modo = 4 And Not PagoTicket And Not SoloFacturado))
    'importe
    txtAux(8).visible = (Modo = 1 Or (Modo = 4 And Not PagoTicket And Not SoloFacturado))
    txtAux(8).Enabled = (Modo = 1 Or (Modo = 4 And Not PagoTicket And Not SoloFacturado))
    'precio
    txtAux(11).visible = (Modo = 1 Or (Modo = 4 And Not PagoTicket And Not SoloFacturado))
    txtAux(11).Enabled = (Modo = 1 Or (Modo = 4 And Not PagoTicket And Not SoloFacturado))
    
    'fecha riego
    txtAux(9).visible = (Modo = 1 Or Modo = 4)
    txtAux(9).Enabled = (Modo = 1 Or Modo = 4)
    
    'fecha de pago
    txtAux(12).visible = (Modo = 4 And PagoTicket And Not SoloFacturado)
    txtAux(12).Enabled = (Modo = 4 And PagoTicket And Not SoloFacturado)
    
    'concepto
    Text2(1).Enabled = (Modo = 1 Or (Modo = 4 And Not PagoTicket And Not SoloFacturado))
    
    
    For I = 0 To Me.btnBuscar.Count - 1
        btnBuscar(I).visible = (Modo = 1)
        btnBuscar(I).Enabled = (Modo = 1)
    Next I
    btnBuscar(3).visible = (Modo = 1 Or Modo = 4)
    btnBuscar(3).Enabled = (Modo = 1 Or Modo = 4)
    btnBuscar(4).visible = (Modo = 1 Or (Modo = 4 And PagoTicket))
    btnBuscar(4).Enabled = (Modo = 1 Or (Modo = 4 And PagoTicket))
    
    Text2(0).visible = (Modo = 1)
    Text2(2).visible = (Modo = 1)
    
    cmdAceptar.visible = Not B
    cmdCancelar.visible = Not B
    DataGrid1.Enabled = B
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = B
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu  'En funcion del usuario
    
    'Si estamos modo Modificar bloquear clave primaria
    BloquearTxt txtAux(0), (Modo = 4)
End Sub


Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botones del la toobar y del menu, segun el modo en que estemos
Dim B As Boolean

    B = (Modo = 2)
    'Busqueda
    Toolbar1.Buttons(2).Enabled = B
    Me.mnBuscar.Enabled = B
    'Ver Todos
    Toolbar1.Buttons(3).Enabled = B
    Me.mnVerTodos.Enabled = B
    'Busqueda facturados
    Toolbar1.Buttons(4).Enabled = B
    Me.mnBuscarFacturados.Enabled = B
    
    'Insertar
    Toolbar1.Buttons(6).Enabled = B And Not DeConsulta
    Me.mnNuevo.Enabled = B And Not DeConsulta
    
    B = (B And adodc1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(7).Enabled = B
    Me.mnModificar.Enabled = B
    'Eliminar
    Toolbar1.Buttons(8).Enabled = B
    Me.mnEliminar.Enabled = B
    'Imprimir
    Toolbar1.Buttons(11).Enabled = B
    Me.mnImprimir.Enabled = B
    'Pago de ticket
    Toolbar1.Buttons(12).Enabled = B
    Me.mnPagoTicket.Enabled = B
    
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
    CadenaB = " WHERE rpozticketsmanta.fecpago is null "
    
    CargaGrid 'primer de tot carregue tot el grid
    CadB = ""
    '******************** canviar taula i camp **************************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = SugerirCodigoSiguienteStr("rpozticketsmanta", "hidrante")
    End If
    '********************************************************************
    'Situamos el grid al final
    AnyadirLinea DataGrid1, adodc1
         
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 206
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    For I = 0 To txtAux.Count - 1
        txtAux(I).Text = ""
    Next I

    LLamaLineas anc, 3 'Pone el form en Modo=3, Insertar
       
    'Ponemos el foco
    PonerFoco txtAux(0)
End Sub

Private Sub BotonVerTodos()
Dim Sql2 As String
Dim SQL As String
    
    CadenaB = " WHERE rpozticketsmanta.fecpago is null "

    CargaGrid ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    ' ***************** canviar per la clau primaria ********
    CadenaB = " WHERE rpozticketsmanta.fecpago is null "
    
    CargaGrid "rpozticketsmanta.codsocio is null"
    '*******************************************************************************
    'Buscar
    For I = 0 To txtAux.Count - 1
        txtAux(I).Text = ""
    Next I
    
    Text2(0).Text = ""
    Text2(2).Text = ""
    
    SoloFacturado = False
    
    LLamaLineas DataGrid1.Top + 206, 1 'Pone el form en Modo=1, Buscar
    PonerFoco txtAux(2)
End Sub


Private Sub BotonBuscarFacturados()
    ' ***************** canviar per la clau primaria ********
    CadenaB = " WHERE not rpozticketsmanta.fecpago is null "
    
    CargaGrid "rpozticketsmanta.codsocio is null"
    '*******************************************************************************
    
    'Buscar
    For I = 0 To txtAux.Count - 1
        txtAux(I).Text = ""
    Next I
    
    Text2(0).Text = ""
    Text2(2).Text = ""
    
    SoloFacturado = True
    
    LLamaLineas DataGrid1.Top + 206, 1 'Pone el form en Modo=1, Buscar
    PonerFoco txtAux(2)
End Sub



Private Sub BotonModificar()
    Dim anc As Single
    Dim I As Integer
    
    Screen.MousePointer = vbHourglass
    
    PagoTicket = False
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 670 '545
    End If

    'Llamamos al form
    txtAux(2).Text = DataGrid1.Columns(0).Text
    txtAux(10).Text = DataGrid1.Columns(1).Text
    txtAux(0).Text = DataGrid1.Columns(2).Text
    Text2(2).Text = DataGrid1.Columns(3).Text
    txtAux(1).Text = DataGrid1.Columns(4).Text
    txtAux(3).Text = DataGrid1.Columns(5).Text
    Text2(0).Text = DataGrid1.Columns(6).Text
    txtAux(4).Text = DataGrid1.Columns(7).Text
    txtAux(5).Text = DataGrid1.Columns(8).Text
    txtAux(6).Text = DataGrid1.Columns(9).Text
    
    txtAux(7).Text = DataGrid1.Columns(10).Text 'hanegadas
    txtAux(8).Text = DataGrid1.Columns(12).Text 'importe
    txtAux(11).Text = DataGrid1.Columns(11).Text 'precio1
    txtAux(9).Text = DataGrid1.Columns(13).Text ' fecha de riego
    txtAux(12).Text = DataGrid1.Columns(14).Text ' Fecha de pago
    

    LLamaLineas anc, 4 'Pone el form en Modo=4, Modificar
   
    'Como es modificar
    If SoloFacturado Then
        PonerFoco txtAux(9)
    Else
        PonerFoco txtAux(7)
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub BotonPagoTicket()
    Dim anc As Single
    Dim I As Integer
    
    If Me.adodc1.Recordset.EOF Then Exit Sub
    
    If DBLet(adodc1.Recordset!fecpago) <> "" Then
        MsgBox "Este ticket ya está pagado. No podemos realizar de nuevo el proceso.", vbExclamation
        Exit Sub
    End If
    
    '[Monica]18/07/2014: preguntamos el tipo de ticket que es
    EsTicketContado = False
    Continuar = False
    
    If Not EsSocioContadoPOZOS(CStr(adodc1.Recordset!Codsocio)) Then
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 58
        frmMens.Show vbModal
        
        Set frmMens = Nothing
    Else
        EsTicketContado = True
        Continuar = True
    End If
    
    If Not Continuar Then Exit Sub
    
    
    
    Screen.MousePointer = vbHourglass
    
    PagoTicket = True
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 670 '545
    End If

    'Llamamos al form
    txtAux(2).Text = DataGrid1.Columns(0).Text
    txtAux(10).Text = DataGrid1.Columns(1).Text
    txtAux(0).Text = DataGrid1.Columns(2).Text
    Text2(2).Text = DataGrid1.Columns(3).Text
    txtAux(1).Text = DataGrid1.Columns(4).Text
    txtAux(3).Text = DataGrid1.Columns(5).Text
    Text2(0).Text = DataGrid1.Columns(6).Text
    txtAux(4).Text = DataGrid1.Columns(7).Text
    txtAux(5).Text = DataGrid1.Columns(8).Text
    txtAux(6).Text = DataGrid1.Columns(9).Text
    
    txtAux(7).Text = DataGrid1.Columns(10).Text 'hanegadas
    txtAux(8).Text = DataGrid1.Columns(12).Text 'importe
    txtAux(11).Text = DataGrid1.Columns(11).Text 'precio1
    txtAux(9).Text = DataGrid1.Columns(13).Text ' fecha de riego
    
    If txtAux(9).Text = "" Then txtAux(9).Text = Format(Now, "dd/mm/yyyy") ' fecha de riego
    
    txtAux(12).Text = Format(Now, "dd/mm/yyyy") ' Fecha de pago
    
    LLamaLineas anc, 4 'Pone el form en Modo=4, Modificar
   
    'Como es modificar
    PonerFoco txtAux(12)
    Screen.MousePointer = vbDefault
End Sub


Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    'Fijamos el ancho
    For I = 0 To txtAux.Count - 1
        txtAux(I).Top = alto
    Next I
    Text2(0).Top = alto
    Text2(2).Top = alto
    For I = 0 To Me.btnBuscar.Count - 1
        btnBuscar(I).Top = alto
    Next I
    ' ### [Monica] 12/09/2006
    
End Sub


Private Sub BotonEliminar()
Dim SQL As String
Dim temp As Boolean
Dim vTipoMov As CTiposMov

    On Error GoTo Error2
    
    
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar Then Exit Sub
        
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCampo(txtAux(0))) Then Exit Sub
    ' ***************************************************************************
    
    '*************** canviar els noms i el DELETE **********************************
    SQL = "¿Seguro que desea eliminar el Ticket?"
    SQL = SQL & vbCrLf & "Código: " & adodc1.Recordset.Fields(0)
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        
        'Devolvemos contador, si no estamos actualizando
        Set vTipoMov = New CTiposMov
        vTipoMov.DevolverContador "ALV", adodc1.Recordset!numalbar
        Set vTipoMov = Nothing
        
        SQL = "Delete from rpozticketsmanta where numalbar=" & adodc1.Recordset!numalbar
        conn.Execute SQL
        CargaGrid CadB
        
        
        temp = SituarDataTrasEliminar(adodc1, NumRegElim, True)
        PonerModoOpcionesMenu
        adodc1.Recordset.Cancel
    End If
    Exit Sub
    
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub

Private Function SepuedeBorrar() As Boolean
    SepuedeBorrar = False
    
    If DBLet(Me.adodc1.Recordset!fecpago) <> "" Then
        MsgBox "No puede borrar un ticket con fecha de pago. Revise.", vbExclamation
        Exit Function
    End If
    
    SepuedeBorrar = True
End Function


Private Sub BotonImprimir()
Dim cadAux As String
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim ConSubInforme As Boolean
Dim Fecha As String
    If Me.adodc1.Recordset.RecordCount = 0 Then Exit Sub
        
    InicializarVbles
        
    'Nº Ticket
'    Fecha = Year(Fecha) & "," & Month(Fecha) & "," & Day(Fecha)
    cadAux = "{rpozticketsmanta.numalbar} = " & Me.adodc1.Recordset!numalbar & " and {rpozticketsmanta.fecalbar} = date(""" & Me.adodc1.Recordset!Fecalbar & """)"
    
    If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
    
    indRPT = 47 'Impresion de recibos de mantenimiento de pozos
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
    'Nombre fichero .rpt a Imprimir
    
    '[Monica]30/07/2014: dependiendo de si el socio es de contado mandamos una impresion u otra
    If EsSocioContadoPOZOS(CStr(adodc1.Recordset!Codsocio)) Then
        cadNombreRPT = Replace(nomDocu, "Mto.", "TicketMantaCont.")
    Else
        cadNombreRPT = Replace(nomDocu, "Mto.", "TicketManta.")
    End If
    
    'Nombre fichero .rpt a Imprimir
    cadTitulo = "Reimpresión Tickets Consumo a Manta"
    ConSubInforme = True

    LlamarImprimir

End Sub
Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    CadParam = ""
    numParam = 0
End Sub

Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = CadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .EnvioEMail = False
        .ConSubInforme = True ' ConSubInforme
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los txtAux
    PonerLongCamposGnral Me, Modo, 3
End Sub

Private Sub btnBuscar_Click(Index As Integer)
    Select Case Index
        Case 0 'socios
            Set frmSoc = New frmManSocios
'            frmSoc.DeConsulta = True
            frmSoc.DatosADevolverBusqueda = "0|1|"
'            frmSoc.CodigoActual = Text1(1).Text
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco txtAux(1)
        Case 2 'partida
            
            Set frmZon = New frmManZonas
            frmZon.DeConsulta = False
            frmZon.DatosADevolverBusqueda = ""
            frmZon.Caption = "Braçals"
            frmZon.DeInformes = True
            
        Case 1, 3, 4 ' fecha
            Dim esq As Long
            Dim dalt As Long
            Dim menu As Long
            Dim obj As Object
        
            Set frmC = New frmCal
            
            Indice = Index
            
            esq = btnBuscar(Index).Left
            dalt = btnBuscar(Index).Top
                
            Set obj = btnBuscar(Index).Container
              
            While btnBuscar(Index).Parent.Name <> obj.Name
                esq = esq + obj.Left
                dalt = dalt + obj.Top
                Set obj = obj.Container
            Wend
            
            menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
        
            frmC.Left = esq + btnBuscar(Index).Parent.Left + 30
            frmC.Top = dalt + btnBuscar(Index).Parent.Top + btnBuscar(Index).Height + menu - 40
        
            
            Select Case Index
                Case 1
                    Indice = 10
                Case 3
                    Indice = 9
                Case 4
                    Indice = 12
            End Select
            
            btnBuscar(3).Tag = Index '<===
            ' *** repasar si el camp es txtAux o Text1 ***
            If txtAux(Indice).Text <> "" Then frmC.NovaData = txtAux(Indice).Text
            
            ' ********************************************
            frmC.Show vbModal
            Set frmC = Nothing
            ' *** repasar si el camp es txtAux o Text1 ***
            PonerFoco txtAux(Indice) '<===
            ' ********************************************
            
    End Select
    
End Sub


Private Sub cmdAceptar_Click()
    Dim I As Long
    Dim NReg As Long
    Dim SQL As String
    Dim Sql2 As String
    
    
    
    Select Case Modo
        Case 1 'BUSQUEDA
            CadB = ObtenerBusqueda(Me)
            If CadB <> "" Then
                
                CargaGrid CadB   '& AnyadeCadenaFiltro(True)
                PonerModo 2
'                lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
                PonerFocoGrid Me.DataGrid1
            End If
            
        Case 3 'INSERTAR
            If DatosOK Then
                If InsertarDesdeForm(Me) Then
                    CargaGrid
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                        cmdCancelar_Click
'                        If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveLast
                        If Not adodc1.Recordset.EOF Then
                            adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & NuevoCodigo)
                        End If
                        cmdRegresar_Click
                    Else
                        BotonAnyadir
                    End If
                    CadB = ""
                End If
            End If
            
        Case 4 'MODIFICAR
            OK = False
            If DatosOK Then
                If Not PagoTicket Then
                    If ModificaDesdeFormulario2(Me, 0) Then
                        OK = True
                        
                        TerminaBloquear
                        I = adodc1.Recordset.Fields(0)
                        PonerModo 2
                        CargaGrid "" 'CadB
                        adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & I & "")
                        PonerFocoGrid Me.DataGrid1
                    End If
                Else
                    If ModificarRegistro Then
                        PonerModo 2
                        CargaGrid "" 'CadB
                    End If
                End If
            End If

    End Select
End Sub

Private Function ModificarRegistro() As Boolean
Dim bol As Boolean
Dim MenError As String
Dim devuelve As String
Dim SQL As String
Dim Rs As ADODB.Recordset

    On Error GoTo EModifica

    bol = False
        
    'Aqui empieza transaccion
    conn.BeginTrans
    
    'Insertar en tablas de Albaranes el Pedido (albaran, albaran_variedad, albaran_calibre)
    bol = ModificaDesdeFormulario2(Me, 0)
    
    If bol Then bol = InsertarFactura(MenError)
    
    
EModifica:
    If Err.Number <> 0 Or Not bol Then
        MenError = "Modificando registro." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        bol = False
    End If
    If bol Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
    ModificarRegistro = bol
End Function


Private Function InsertarFactura(MenError As String) As Boolean
'Devuelve el mensaje de error si se produce
Dim bol As Boolean, Existe As Boolean
Dim devuelve As String
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim vSeccion As CSeccion
Dim CodTipom As String
Dim numfactu As Long
Dim vPorcIva As String
Dim PorcIva As Currency
Dim ImpoIva As Currency
Dim Concepto As String


    On Error GoTo EInsertarFactura
    
    bol = False
    InsertarFactura = bol
    
    'Obtener el Contador de ALBARAN
    '[Monica]02/07/2012: antes cogiamos el tipo de movimiento de parametros ahora lo cogemos de clientes
    'codTipoM = vParamAplic.CodTipomAlb ' "ALV"
    
    CodTipom = "RMT"
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    vPorcIva = ""
    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaPOZ, "N")
    PorcIva = CCur(ImporteSinFormato(vPorcIva))
    ImpoIva = 0
    
    Set vSeccion = Nothing
    
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipom) Then
        'Comprobar si mientras tanto se incremento el contador de albaranes
        Do
            numfactu = vTipoMov.ConseguirContador(CodTipom)
            devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "numfactu", CStr(numfactu), "N", , "codtipom", CodTipom, "T")
            If devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (CodTipom)
                numfactu = vTipoMov.ConseguirContador(CodTipom)
            Else
                Existe = False
            End If
        Loop Until Not Existe
            
    Else 'No existe el tipo de Movimiento
        Set vTipoMov = Nothing
        Exit Function
    End If
    
    Concepto = Text2(1).Text & " Nro-" & txtAux(2).Text & " Fecha Riego " & txtAux(9).Text
    
    
    'insertar en la tabla de recibos de pozos
    SQL = "insert into rrecibpozos (codtipom, numfactu, fecfactu, codsocio, hidrante, baseimpo, tipoiva, porc_iva, imporiva, "
    SQL = SQL & "totalfact , consumo, impcuota, lect_ant, fech_ant, lect_act, fech_act, consumo1, precio1, consumo2, precio2, "
    SQL = SQL & "concepto, contabilizado, porcdto, impdto, precio, numalbar, fecalbar, escontado) "
    SQL = SQL & " values ('" & CodTipom & "'," & DBSet(numfactu, "N") & "," & DBSet(txtAux(12).Text, "F") & "," & DBSet(txtAux(0).Text, "N") & ","
    SQL = SQL & ValorNulo & "," & DBSet(txtAux(8).Text, "N") & "," & DBSet(vParamAplic.CodIvaPOZ, "N") & "," & DBSet(PorcIva, "N") & "," & DBSet(ImpoIva, "N") & ","
    SQL = SQL & DBSet(txtAux(8).Text, "N") & "," & ValorNulo & "," & ValorNulo & ","
    SQL = SQL & ValorNulo & "," & ValorNulo & ","
    SQL = SQL & ValorNulo & "," & ValorNulo & ","
    SQL = SQL & ValorNulo & "," & ValorNulo & ","
    SQL = SQL & ValorNulo & "," & ValorNulo & ","
    SQL = SQL & DBSet(Concepto, "T") & ",0,"
    SQL = SQL & DBSet(0, "N") & ","
    SQL = SQL & DBSet(0, "N") & ","
    SQL = SQL & DBSet(CCur(ImporteSinFormato(txtAux(11).Text)), "N") & ","
    SQL = SQL & DBSet(txtAux(2).Text, "N") & ","
    SQL = SQL & DBSet(txtAux(10).Text, "F") & ","
    SQL = SQL & DBSet(EsTicketContado, "B") & ")"
    
    conn.Execute SQL
        
        
    ' Introducimos en la tabla de lineas de campos que intervienen en la factura para la impresion
    ' SOLO HABRA UN CAMPO
    SQL = "insert into rrecibpozos_cam (codtipom, numfactu, fecfactu, codcampo, hanegada, precio1, codzonas, poligono, parcela, subparce) values  "
    
    SQL = SQL & "('" & CodTipom & "'," & DBSet(numfactu, "N") & "," & DBSet(txtAux(12).Text, "F") & ","
    SQL = SQL & DBSet(txtAux(1).Text, "N") & "," & DBSet(txtAux(7).Text, "N") & "," & DBSet(txtAux(11).Text, "N") & ","
    SQL = SQL & DBSet(txtAux(3).Text, "N") & "," & DBSet(txtAux(4).Text, "N") & "," & DBSet(txtAux(5).Text, "N") & "," & DBSet(txtAux(5).Text, "T")
    SQL = SQL & ")"

    conn.Execute SQL

    vTipoMov.IncrementarContador (CodTipom)
    Set vTipoMov = Nothing
    
    bol = True
    
EInsertarFactura:
        If Err.Number <> 0 Then bol = False
        InsertarFactura = bol
End Function







Private Sub cmdCancelar_Click()
    On Error Resume Next
    
    Select Case Modo
        Case 1 'búsqueda
            CargaGrid CadB
        Case 3 'insertar
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        Case 4 'modificar
            TerminaBloquear
    End Select
    
    PonerModo 2
    
    PonerFocoGrid Me.DataGrid1
    If Err.Number <> 0 Then Err.Clear
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
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
Dim cad As String
    
    If adodc1.Recordset Is Nothing Then Exit Sub
    If adodc1.Recordset.EOF Then Exit Sub
    
    Me.Refresh
    DoEvents
    Screen.MousePointer = vbHourglass
    
    Ordenacion = "ORDER BY " & DataGrid1.Columns(0).DataField
    
    CadB = ""
    CargaGrid CadB
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Modo = 2 Then
        PonerContRegIndicador lblIndicador, adodc1, CadB
        Text2(1).Text = adodc1.Recordset!Concepto
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault

    If PrimeraVez Then
        PrimeraVez = False
        If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
            BotonAnyadir
        Else
            PonerModo 2
             If Me.CodigoActual <> "" Then
                SituarData Me.adodc1, "hidrante='" & CodigoActual & "'", "", True
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
Dim Sql2 As String

    PrimeraVez = True
    
    'Icono del formulario
    Me.Icon = frmPpal.Icon


    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'el 1 es separadors
        .Buttons(2).Image = 1   'Buscar
        .Buttons(3).Image = 2   'Todos
        .Buttons(4).Image = 21   'Buscar Facturados
        
        'el 4 i el 5 son separadors
        .Buttons(6).Image = 3   'Insertar
        .Buttons(7).Image = 4   'Modificar
        .Buttons(8).Image = 5   'Borrar
        'el 9 i el 10 son separadors
        .Buttons(11).Image = 10  'Impresión del ticket
        
        .Buttons(12).Image = 13  'Pago del Ticket
        
        .Buttons(13).Image = 11  'Salir
    End With

    '## A mano
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    CodTipoMov = "ALF"

    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT rpozticketsmanta.numalbar, rpozticketsmanta.fecalbar, rpozticketsmanta.codsocio, rsocios.nomsocio, rpozticketsmanta.codcampo, "
    CadenaConsulta = CadenaConsulta & "rpozticketsmanta.codzonas, rzonas.nomzonas, rpozticketsmanta.poligono, rpozticketsmanta.parcela, rpozticketsmanta.subparce,"
    CadenaConsulta = CadenaConsulta & "rpozticketsmanta.hanegada, rpozticketsmanta.precio1, rpozticketsmanta.importe, rpozticketsmanta.fecriego, rpozticketsmanta.fecpago, "
    CadenaConsulta = CadenaConsulta & " rpozticketsmanta.concepto "
    CadenaConsulta = CadenaConsulta & " FROM ((rpozticketsmanta INNER JOIN rsocios ON rpozticketsmanta.codsocio = rsocios.codsocio) "
    CadenaConsulta = CadenaConsulta & " INNER JOIN rzonas ON rpozticketsmanta.codzonas = rzonas.codzonas)"
'    CadenaConsulta = CadenaConsulta & " WHERE rpozticketsmanta.fecpago is null "
    '************************************************************************
    
    CadenaB = " WHERE rpozticketsmanta.fecpago is null "
    
    Ordenacion = " ORDER BY 1 "
    
    CadB = ""
    CargaGrid
    
    FechaAnt = ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
    
    If Modo = 4 Then TerminaBloquear
End Sub

Private Sub frmC_Selec(vFecha As Date)
    txtAux(Indice).Text = Format(vFecha, "dd/mm/yyyy") '<===
End Sub


Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion = "" Then
        Continuar = False
    Else
        EsTicketContado = (CadenaSeleccion = "1")
        Continuar = True
    End If
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1) 'codsocio
    FormateaCampo txtAux(1)
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2) 'nomsocio
End Sub

Private Sub mnActualizar_Click()
End Sub

Private Sub frmZon_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo de zona
    FormateaCampo txtAux(3)
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre de zona
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub


Private Sub mnBuscarFacturados_Click()
    BotonBuscarFacturados
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    BotonImprimir
End Sub

Private Sub mnModificar_Click()
    'Comprobaciones
    '--------------
    If adodc1.Recordset.EOF Then Exit Sub
    
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCampo(txtAux(0))) Then Exit Sub
    
    
    'Preparamos para modificar
    '-------------------------
    If BLOQUEADesdeFormulario2(Me, adodc1, 1) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    AbrirListadoPOZ 17
    CargaGrid ""
    PonerModo 2
End Sub

Private Sub mnPagoTicket_Click()
    BotonPagoTicket
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 2
                mnBuscar_Click
        Case 3
                mnVerTodos_Click
        Case 4
                mnBuscarFacturados_Click
        Case 6
                mnNuevo_Click
        Case 7
                mnModificar_Click
        Case 8
                mnEliminar_Click
        Case 11
                mnImprimir_Click
        Case 12
                mnPagoTicket_Click
        Case 13
                mnSalir_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String)
    Dim SQL As String
    Dim tots As String
    Dim Sql2 As String
    
    
    If vSQL <> "" Then
        SQL = CadenaConsulta & " " & CadenaB & " AND " & vSQL
    Else
        SQL = CadenaConsulta & " " & CadenaB
    End If
    '********************* canviar el ORDER BY *********************++
    SQL = SQL & " " & Ordenacion
    '**************************************************************++
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, SQL, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "S|txtAux(2)|T|Albarán|800|;S|txtAux(10)|T|Fecha|1100|;S|btnBuscar(1)|B|||;S|txtAux(0)|T|Socio|800|;S|btnBuscar(0)|B||195|;S|Text2(2)|T|Nombre|2450|;S|txtAux(1)|T|Campo|1000|;"
    tots = tots & "S|txtAux(3)|T|Cód|600|;S|btnBuscar(2)|B||195|;S|Text2(0)|T|Braçal|1500|;"
    tots = tots & "S|txtAux(4)|T|Pol|500|;S|txtAux(5)|T|Parc|800|;S|txtAux(6)|T|Sb|500|;"
    tots = tots & "S|txtAux(7)|T|Hdas|800|;S|txtAux(11)|T|Precio|800|;S|txtAux(8)|T|Importe|1000|;S|txtAux(9)|T|Fec.Riego|1100|;S|btnBuscar(3)|B|||;S|txtAux(12)|T|Fec.Pago|1100|;S|btnBuscar(4)|B|||;N|Text2(1)(8)|T|Importe|1000|;"
    
    arregla tots, DataGrid1, Me
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgLeft
    DataGrid1.Columns(2).Alignment = dbgLeft
    DataGrid1.Columns(4).Alignment = dbgLeft
    DataGrid1.Columns(6).Alignment = dbgLeft
    DataGrid1.Columns(7).Alignment = dbgLeft
    DataGrid1.Columns(8).Alignment = dbgCenter
    
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 2 'albaran
            PonerFormatoEntero txtAux(Index)
        
        Case 3, 5 ' lectura anterior / lectura actual
            PonerFormatoEntero txtAux(Index)
             
        Case 9, 10, 12 ' fechas
            '[Monica]28/08/2013: no comprobamos que la fecha esté en la campaña
            PonerFormatoFecha txtAux(Index)
            
        Case 7 'hanegadas
            PonerFormatoDecimal txtAux(Index), 10
            
        Case 1, 4, 5 ' campo, poligono y parcela
            PonerFormatoEntero txtAux(Index)
            
        Case 0 'socio
            If txtAux(Index).Text <> "" Then txtAux(Index).Text = Format(txtAux(Index).Text, "000000")
            Text2(2).Text = PonerNombreDeCod(txtAux(Index), "rsocios", "nomsocio", "codsocio", "N")

        Case 3 'zona
            If txtAux(Index).Text <> "" Then txtAux(Index).Text = Format(txtAux(Index).Text, "0000")
            Text2(0).Text = PonerNombreDeCod(txtAux(Index), "rzonas", "nomzonas", "codzonas", "N")
        
        Case 11 'precio
            PonerFormatoDecimal txtAux(Index), 7
        
        Case 8 ' importe
            PonerFormatoDecimal txtAux(Index), 3
    End Select
    If txtAux(7).Text <> "" And txtAux(11).Text <> "" Then
        txtAux(8).Text = Round2(ImporteSinFormato(ComprobarCero(txtAux(7).Text)) * ImporteSinFormato(ComprobarCero(txtAux(11).Text)), 2)
        txtAux(8).Text = Format(txtAux(8).Text, "#,###,##0.00")
    End If
End Sub

Private Function DatosOK() As Boolean
'Dim Datos As String
Dim B As Boolean
Dim SQL As String
Dim Mens As String
Dim FechaAnt As Date
Dim NroDig As Integer
Dim Inicio As Long
Dim Fin As Long
Dim Consumo As Long
Dim Limite As Long

    B = CompForm(Me)
    If Not B Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
         If ExisteCP(txtAux(0)) Then B = False
    End If
    
    'comprobamos que la fecha de factura que se va a generar está dentro del periodo de liquidacion
    If Modo = 4 And PagoTicket Then
        If txtAux(12).Text = "" Then
            MsgBox "Debe introducir una fecha de recibo. Reintroduzca.", vbInformation
            B = False
            PonerFoco txtAux(12)
        End If
    End If
    
    DatosOK = B
End Function

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub

'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGrid1_GotFocus()
'  WheelHook DataGrid1
'End Sub
'Private Sub DataGrid1_Lostfocus()
'  WheelUnHook
'End Sub


Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
   KEYpress KeyAscii
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me

End Sub

