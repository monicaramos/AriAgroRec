VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManTranspor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transportistas"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12495
   Icon            =   "frmManTranspor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   12495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux1 
      Alignment       =   1  'Right Justify
      Height          =   290
      Index           =   2
      Left            =   7770
      MaxLength       =   4
      TabIndex        =   15
      Tag             =   "IBAN|T|S|||rtransporte|iban|||"
      Top             =   3690
      Width           =   600
   End
   Begin VB.CheckBox chkAbonos 
      Caption         =   "Factura Interna de Transporte"
      Height          =   195
      Index           =   1
      Left            =   8760
      TabIndex        =   24
      Tag             =   "Se factura|N|N|0|1|rtransporte|esfacttrainterna||N|"
      Top             =   4830
      Width           =   2535
   End
   Begin VB.TextBox txtAux2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   8760
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   50
      Top             =   4440
      Width           =   3315
   End
   Begin VB.TextBox txtAux1 
      Height          =   285
      Index           =   1
      Left            =   7770
      MaxLength       =   10
      TabIndex        =   22
      Tag             =   "Trabajador|N|S|||rtransporte|codtraba|000000||"
      Top             =   4440
      Width           =   945
   End
   Begin VB.TextBox txtAux1 
      Alignment       =   1  'Right Justify
      Height          =   290
      Index           =   0
      Left            =   11070
      MaxLength       =   7
      TabIndex        =   21
      Tag             =   "contador Factura|N|N|0|9999999|rtransporte|contador|0000000||"
      Top             =   4020
      Width           =   1020
   End
   Begin VB.CheckBox chkAbonos 
      Caption         =   "Se emite factura"
      Height          =   195
      Index           =   0
      Left            =   6720
      TabIndex        =   23
      Tag             =   "Se factura|N|N|0|1|rtransporte|sefactura||N|"
      Top             =   4830
      Width           =   1515
   End
   Begin VB.TextBox txtAux1 
      Alignment       =   1  'Right Justify
      Height          =   290
      Index           =   15
      Left            =   8430
      MaxLength       =   4
      TabIndex        =   16
      Tag             =   "Banco|N|S|0|9999|rtransporte|codbanco|0000||"
      Top             =   3690
      Width           =   600
   End
   Begin VB.TextBox txtAux1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   18
      Left            =   10380
      MaxLength       =   10
      TabIndex        =   19
      Tag             =   "Cuenta Bancaria|T|S|||rtransporte|cuentaba|0000000000||"
      Top             =   3690
      Width           =   1665
   End
   Begin VB.TextBox txtAux1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   17
      Left            =   9750
      MaxLength       =   2
      TabIndex        =   18
      Tag             =   "Digito Control|T|S|||rtransporte|digcontr|00||"
      Top             =   3690
      Width           =   585
   End
   Begin VB.TextBox txtAux1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   16
      Left            =   9090
      MaxLength       =   4
      TabIndex        =   17
      Tag             =   "Sucursal|N|S|0|9999|rtransporte|codsucur|0000||"
      Top             =   3690
      Width           =   630
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   7770
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Tag             =   "Tipo IRPF|N|N|0|2|rtransporte|tipoirpf||N|"
      Top             =   4050
      Width           =   1920
   End
   Begin VB.TextBox txtAux1 
      Height          =   290
      Index           =   6
      Left            =   7770
      MaxLength       =   2
      TabIndex        =   14
      Tag             =   "Cod.Iva|N|N|||rtransporte|codiva|00||"
      Top             =   3330
      Width           =   945
   End
   Begin VB.TextBox txtAux1 
      Height          =   290
      Index           =   5
      Left            =   7770
      MaxLength       =   10
      TabIndex        =   13
      Tag             =   "Cta.Contable Proveedor|T|S|||rtransporte|codmacpro|||"
      Top             =   2970
      Width           =   1065
   End
   Begin VB.TextBox txtAux1 
      Height          =   285
      Index           =   4
      Left            =   7770
      MaxLength       =   10
      TabIndex        =   12
      Tag             =   "Forma Pago|N|N|||rtransporte|codforpa|000||"
      Top             =   2610
      Width           =   945
   End
   Begin VB.TextBox txtAux2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   8760
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   43
      Top             =   3330
      Width           =   3315
   End
   Begin VB.TextBox txtAux2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   8880
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   42
      Top             =   2970
      Width           =   3195
   End
   Begin VB.TextBox txtAux2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   8760
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   41
      Top             =   2610
      Width           =   3315
   End
   Begin VB.TextBox txtAux1 
      Height          =   285
      Index           =   14
      Left            =   7770
      MaxLength       =   35
      TabIndex        =   6
      Tag             =   "Domicilio|T|N|||rtransporte|dirtrans|||"
      Top             =   840
      Width           =   4320
   End
   Begin VB.TextBox txtAux1 
      Height          =   285
      Index           =   13
      Left            =   7770
      MaxLength       =   30
      TabIndex        =   7
      Tag             =   "Población|T|N|||rtransporte|pobtrans|||"
      Top             =   1185
      Width           =   4320
   End
   Begin VB.TextBox txtAux1 
      Height          =   285
      Index           =   12
      Left            =   7770
      MaxLength       =   30
      TabIndex        =   8
      Tag             =   "Provincia|T|N|||rtransporte|protrans|||"
      Top             =   1530
      Width           =   4320
   End
   Begin VB.TextBox txtAux1 
      Height          =   285
      Index           =   10
      Left            =   10860
      MaxLength       =   6
      TabIndex        =   5
      Tag             =   "Código Postal|T|N|||rtransporte|codpostal|||"
      Top             =   510
      Width           =   1230
   End
   Begin VB.TextBox txtAux1 
      Height          =   285
      Index           =   7
      Left            =   7770
      MaxLength       =   9
      TabIndex        =   4
      Tag             =   "CIF|T|N|||rtransporte|niftrans|||"
      Top             =   510
      Width           =   1095
   End
   Begin VB.TextBox txtAux1 
      Height          =   285
      Index           =   11
      Left            =   7770
      MaxLength       =   40
      TabIndex        =   11
      Tag             =   "Mail|T|S|||rtransporte|maitrans|||"
      Top             =   2250
      Width           =   4320
   End
   Begin VB.TextBox txtAux1 
      Height          =   285
      Index           =   8
      Left            =   7770
      MaxLength       =   10
      TabIndex        =   9
      Tag             =   "Teléfono|T|S|||rtransporte|teltran1|||"
      Top             =   1905
      Width           =   1695
   End
   Begin VB.TextBox txtAux1 
      Height          =   285
      Index           =   9
      Left            =   10380
      MaxLength       =   10
      TabIndex        =   10
      Tag             =   "Fax|T|S|||rtransporte|movtrans|||"
      Top             =   1905
      Width           =   1710
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   3960
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "Tara Vehiculo|N|S|||rtransporte|taravehi|###,##0||"
      Top             =   4545
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   2430
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "Matrícula|T|S|||rtransporte|matricula|||"
      Top             =   4545
      Width           =   1395
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9840
      TabIndex        =   25
      Tag             =   "   "
      Top             =   5220
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10980
      TabIndex        =   26
      Top             =   5220
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   900
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Descripción|T|N|||rtransporte|nomtrans|||"
      Top             =   4560
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   60
      MaxLength       =   10
      TabIndex        =   0
      Tag             =   "Código|T|N|||rtransporte|codtrans||S|"
      Top             =   4560
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmManTranspor.frx":000C
      Height          =   4410
      Left            =   120
      TabIndex        =   29
      Top             =   540
      Width           =   6435
      _ExtentX        =   11351
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
      Left            =   10980
      TabIndex        =   32
      Top             =   5190
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   90
      TabIndex        =   27
      Top             =   5100
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
         TabIndex        =   28
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
      TabIndex        =   30
      Top             =   0
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
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
            Object.Visible         =   0   'False
            Style           =   3
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
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   3735
         TabIndex        =   31
         Top             =   90
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Trabajador"
      Height          =   285
      Left            =   6690
      TabIndex        =   51
      Top             =   4440
      Width           =   765
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   3
      Left            =   7500
      ToolTipText     =   "Buscar Trabajador"
      Top             =   4455
      Width           =   240
   End
   Begin VB.Label Label4 
      Caption         =   "Contador"
      Height          =   255
      Left            =   10110
      TabIndex        =   49
      Top             =   4050
      Width           =   915
   End
   Begin VB.Label Label38 
      Caption         =   "IBAN Transp."
      Height          =   255
      Left            =   6690
      TabIndex        =   48
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   2
      Left            =   7500
      ToolTipText     =   "Buscar iva"
      Top             =   3330
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   7500
      ToolTipText     =   "Buscar F.Pago"
      Top             =   2625
      Width           =   240
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   7500
      ToolTipText     =   "Buscar cuenta"
      Top             =   2985
      Width           =   240
   End
   Begin VB.Label Label43 
      Caption         =   "Tipo IRPF"
      Height          =   285
      Left            =   6690
      TabIndex        =   47
      Top             =   4050
      Width           =   1035
   End
   Begin VB.Label Label13 
      Caption         =   "F.Pago"
      Height          =   285
      Left            =   6690
      TabIndex        =   46
      Top             =   2610
      Width           =   555
   End
   Begin VB.Label Label12 
      Caption         =   "Iva"
      Height          =   285
      Left            =   6690
      TabIndex        =   45
      Top             =   3330
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "Cta.Prov."
      Height          =   285
      Left            =   6690
      TabIndex        =   44
      Top             =   2970
      Width           =   690
   End
   Begin VB.Label Label1 
      Caption         =   "C.Postal"
      Height          =   255
      Index           =   8
      Left            =   9870
      TabIndex        =   40
      Top             =   540
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Población"
      Height          =   255
      Index           =   9
      Left            =   6690
      TabIndex        =   39
      Top             =   1215
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Provincia"
      Height          =   255
      Left            =   6690
      TabIndex        =   38
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Domicilio"
      Height          =   255
      Index           =   7
      Left            =   6690
      TabIndex        =   37
      Top             =   870
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "CIF"
      Height          =   255
      Left            =   6690
      TabIndex        =   36
      Top             =   540
      Width           =   1215
   End
   Begin VB.Image imgMail 
      Height          =   240
      Index           =   0
      Left            =   7500
      Top             =   2265
      Width           =   240
   End
   Begin VB.Label Label7 
      Caption         =   "Mail"
      Height          =   255
      Left            =   6690
      TabIndex        =   35
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Teléfono"
      Height          =   285
      Index           =   1
      Left            =   6690
      TabIndex        =   34
      Top             =   1905
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Móvil"
      Height          =   255
      Left            =   9840
      TabIndex        =   33
      Top             =   1935
      Width           =   495
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
Attribute VB_Name = "frmManTranspor"
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

'codi per al registe que s'afegix al cridar des d'atre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

Private CadenaConsulta As String
Private CadB As String

Private WithEvents frmCtas As frmCtasConta 'cuentas contables
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmTIva As frmTipIVAConta 'Tipos de iva de conta
Attribute frmTIva.VB_VarHelpID = -1
Private WithEvents frmFPa As frmForpaConta 'formas de pago de la contabilidad
Attribute frmFPa.VB_VarHelpID = -1
Private WithEvents frmTra As frmManTraba 'trabajadores
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'mensajes para pantalla de contadores
Attribute frmMens.VB_VarHelpID = -1


Private BuscaChekc As String

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
Dim PrimeraVez As Boolean
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim i As Integer

'Cambio en cuentas de la contabilidad
Dim IbanAnt As String
Dim NombreAnt As String
Dim BancoAnt  As String
Dim SucurAnt As String
Dim DigitoAnt As String
Dim CuentaAnt As String

Dim DirecAnt As String
Dim cPostalAnt As String
Dim PoblaAnt As String
Dim ProviAnt As String
Dim NifAnt As String
Dim forpaant As String


Dim EMaiAnt As String
Dim WebAnt As String





Private Sub chkAbonos_GotFocus(Index As Integer)
    PonerFocoChk Me.chkAbonos(Index)
End Sub

Private Sub chkAbonos_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkAbonos(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "chkAbonos(" & Index & ")|"
    Else
        If Index = 0 And (Modo = 3 Or Modo = 4) Then
            If chkAbonos(Index).Value = 0 Then
                chkAbonos(1).Value = 0
                chkAbonos(1).Enabled = False
            Else
                chkAbonos(1).Enabled = True
            End If
        End If
    End If
End Sub

Private Sub chkAbonos_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkAbonos_LostFocus(Index As Integer)
'    If Index = 1 And (Modo = 3 Or Modo = 4) Then
'        If chkAbonos(Index).Value = 1 Then Text1(25).Text = ""
'    End If
End Sub

Private Sub PonerModo(vModo)
Dim b As Boolean

    Modo = vModo
    
    b = (Modo = 2)
    If b Then
        PonerContRegIndicador lblIndicador, adodc1, CadB
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    For i = 0 To txtAux.Count - 1
        txtAux(i).visible = Not b
    Next i
    
    ' **** si n'hi han camps fora del grid, bloquejar-los ****
    BloquearTxt txtAux1(0), b
    BloquearTxt txtAux1(1), b
    BloquearTxt txtAux1(2), b
    For i = 4 To 18
        BloquearTxt txtAux1(i), b
    Next i
    
    BloquearCmb Me.Combo1(0), b
    For i = 0 To 3
        Me.imgBuscar(i).Enabled = Not b
        Me.imgBuscar(i).visible = Not b
    Next i
    
    BloquearChk Me.chkAbonos(0), (Modo = 0 Or Modo = 2)
    BloquearChk Me.chkAbonos(1), (Modo = 0 Or Modo = 2)
    
    
'    For i = 0 To Me.cmdAux.Count - 1
'        cmdAux(i).visible = Not b
'        cmdAux(i).Enabled = Not b
'    Next i
    
'    Combo1(0).visible = Not b
'    Combo1(0).Enabled = Not b
'
    
    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    DataGrid1.Enabled = b
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = b
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu  'En funcion del usuario
    
    'Si estamos modo Modificar bloquear clave primaria
    BloquearTxt txtAux(0), (Modo = 4)
End Sub


Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botones del la toobar y del menu, segun el modo en que estemos
Dim b As Boolean

    b = (Modo = 2)
    'Busqueda
    Toolbar1.Buttons(2).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(3).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(6).Enabled = b And Not DeConsulta
    Me.mnNuevo.Enabled = b And Not DeConsulta
    
    b = (b And adodc1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnModificar.Enabled = b
    'Eliminar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnEliminar.Enabled = b
    'Imprimir
    Toolbar1.Buttons(11).Enabled = b
    Me.mnImprimir.Enabled = b
    
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
    CargaGrid 'primer de tot carregue tot el grid
    CadB = ""
    '******************** canviar taula i camp **************************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = SugerirCodigoSiguienteStr("rtransporte", "codtrans")
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
    txtAux(0).Text = NumF
    FormateaCampo txtAux(0)
    For i = 1 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i
    Me.chkAbonos(0).Value = 0
    Me.chkAbonos(1).Value = 0

    For i = 4 To 6
        txtAux2(i).Text = ""
    Next i
    
    Combo1(0).ListIndex = 0
    
    LLamaLineas anc, 3 'Pone el form en Modo=3, Insertar
       
    'Ponemos el foco
    PonerFoco txtAux(0)
End Sub

Private Sub BotonVerTodos()
    CadB = ""
    CargaGrid ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    ' ***************** canviar per la clau primaria ********
    CargaGrid "rtransporte.codtrans = '-1'"
    '*******************************************************************************
    'Buscar
    For i = 0 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i

    For i = 0 To 2
        txtAux1(i).Text = ""
    Next i
    For i = 4 To 18
        txtAux1(i).Text = ""
    Next i
    Me.Combo1(0).ListIndex = -1
    
    chkAbonos(0).Value = 0
    chkAbonos(1).Value = 0
    
    
'    PosicionarCombo Combo1, "724"
    LLamaLineas DataGrid1.Top + 206, 1 'Pone el form en Modo=1, Buscar
    PonerFoco txtAux(0)
End Sub

Private Sub BotonModificar()
    Dim anc As Single
    Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 545
    End If

    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux(2).Text = DataGrid1.Columns(2).Text
    txtAux(3).Text = DataGrid1.Columns(3).Text
    
    For i = 4 To 7
        txtAux1(i).Text = DataGrid1.Columns(i + 1).Text
    Next i
    txtAux1(14).Text = DataGrid1.Columns(9).Text
    txtAux1(13).Text = DataGrid1.Columns(10).Text
    txtAux1(12).Text = DataGrid1.Columns(11).Text
    txtAux1(10).Text = DataGrid1.Columns(12).Text
    txtAux1(8).Text = DataGrid1.Columns(13).Text
    txtAux1(9).Text = DataGrid1.Columns(14).Text
    txtAux1(11).Text = DataGrid1.Columns(15).Text
    For i = 15 To 18
        txtAux1(i).Text = DataGrid1.Columns(i + 1).Text
    Next i
    txtAux1(1).Text = DataGrid1.Columns(22).Text
    txtAux1(2).Text = DataGrid1.Columns(24).Text
    ' ***** canviar-ho pel nom del camp del combo *********
'    SelComboBool DataGrid1.Columns(2).Text, Combo1(0)
    ' *****************************************************

    PosicionarCombo Me.Combo1(0), DataGrid1.Columns(4).Text
    
    'PosicionarCombo Me.Combo1(1), i

    '[Monica]08/11/2016: no modificabamos los datos de la cuenta
    NombreAnt = txtAux(1).Text
    IbanAnt = txtAux1(2).Text
    BancoAnt = txtAux1(15).Text
    SucurAnt = txtAux1(16).Text
    DigitoAnt = txtAux1(17).Text
    CuentaAnt = txtAux1(18).Text
    
    DirecAnt = txtAux1(14).Text
    cPostalAnt = txtAux1(10).Text
    PoblaAnt = txtAux1(13).Text
    ProviAnt = txtAux1(12).Text
    NifAnt = txtAux1(7).Text
    EMaiAnt = txtAux1(11).Text
    
    '[Monica]26/03/2015: antes no se grababa la forma de pago en la cuenta de cliente
    forpaant = txtAux1(4).Text

    LLamaLineas anc, 4 'Pone el form en Modo=4, Modificar
   
    'Como es modificar
    PonerFoco txtAux(1)
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
Dim jj As Integer

    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    'Fijamos el ancho
    For i = 0 To txtAux.Count - 1
        txtAux(i).Top = alto
    Next i
    
'    For jj = 0 To cmdAux.Count - 1
'        cmdAux(jj).visible = (Modo = 1 Or Modo = 3 Or Modo = 4)
'        cmdAux(jj).Top = txtAux(3).Top
'        cmdAux(jj).Height = txtAux(3).Height
'    Next jj
'    Combo1(0).visible = (Modo = 1 Or Modo = 3 Or Modo = 4)
'    Combo1(0).Top = txtAux(3).Top

End Sub

Private Sub BotonEliminar()
Dim Sql As String
Dim temp As Boolean

    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
'    If Not SepuedeBorrar Then Exit Sub
        
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCampo(txtAux(0))) Then Exit Sub
    ' ***************************************************************************
    
    '*************** canviar els noms i el DELETE **********************************
    Sql = "¿Seguro que desea eliminar el Transportista?"
    Sql = Sql & vbCrLf & "Código: " & adodc1.Recordset.Fields(0)
    Sql = Sql & vbCrLf & "Nombre: " & adodc1.Recordset.Fields(1)
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        
        'chivato
        CargarUnVehiculo CStr(adodc1.Recordset!codTrans), "D"
        
        Sql = "Delete from rtransporte where codtrans=" & DBSet(adodc1.Recordset!codTrans, "T")
        conn.Execute Sql
        CargaGrid CadB
'        If CadB <> "" Then
'            CargaGrid CadB
'            lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
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



Private Sub cmdAceptar_Click()
    Dim i As String

    Select Case Modo
        Case 1 'BUSQUEDA
            '[Monica]09/01/2015: nuevo tipo de datos para la busqueda sin asteriscos
            txtAux(0).Tag = "Código|TT|N|||rtransporte|codtrans||S|"
            CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
            txtAux(0).Tag = "Código|T|N|||rtransporte|codtrans||S|"
            
            If CadB <> "" Then
                CargaGrid CadB
                PonerModo 2
'                lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
                PonerFocoGrid Me.DataGrid1
            End If
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                
                    'chivato
                    CargarUnVehiculo txtAux(0).Text, "I"
                
                    CargaGrid
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                        cmdCancelar_Click
'                        If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveLast
                        If Not adodc1.Recordset.EOF Then
                            adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " ='" & Trim(NuevoCodigo) & "'")
                        End If
                        cmdRegresar_Click
                    Else
                        BotonAnyadir
                    End If
                    CadB = ""
                End If
            End If
            
        Case 4 'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me) Then
                    TerminaBloquear
                
                    'chivato
                    CargarUnVehiculo txtAux(0).Text, "U"
                
' De momento lo dejo comentado, pq no se permite crear la cuenta desde el transportista
'                    '[Monica]08/11/2016: Si han cambiado nombre o CCC pregunto si quieren cambiar los datos de la cuenta en la seccion de horto
'                    ModificarDatosCuentaContable
                
                
                    i = adodc1.Recordset.Fields(0)
                    PonerModo 2
                    CargaGrid CadB
'                    If CadB <> "" Then
'                        CargaGrid CadB
'                        lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
'                    Else
'                        CargaGrid
'                        lblIndicador.Caption = ""
'                    End If
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " ='" & i & "'")
                    PonerFocoGrid Me.DataGrid1
                End If
            End If
    End Select
End Sub

Private Sub cmdAux_Click(Index As Integer)
    Select Case Index
        Case 0
            If vSeccion Is Nothing Then Exit Sub
            
            indice = 4
            Set frmFPa = New frmForpaConta
            frmFPa.DatosADevolverBusqueda = "0|1|"
            frmFPa.CodigoActual = txtAux(indice)
        '    frmFpa.Conexion = cContaFacSoc
            frmFPa.Show vbModal
            Set frmFPa = Nothing
            PonerFoco txtAux(indice)
        
        Case 1 'cuentas contables de y proveedor
            If vSeccion Is Nothing Then Exit Sub
            
            indice = 5
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = txtAux(indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco txtAux(indice)
        
        
        Case 2 'codigo de iva
            If vSeccion Is Nothing Then Exit Sub
            
            indice = 6
            Set frmTIva = New frmTipIVAConta
            frmTIva.DeConsulta = True
            frmTIva.DatosADevolverBusqueda = "0|1|"
            frmTIva.CodigoActual = txtAux(6).Text
            frmTIva.Show vbModal
            Set frmTIva = Nothing
            PonerFoco txtAux(6)
     End Select
End Sub

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
    
'    If CadB <> "" Then
'        lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
''    Else
''        lblIndicador.Caption = ""
'    End If
    
    PonerFocoGrid Me.DataGrid1
    If Err.Number <> 0 Then Err.Clear
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

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Modo = 2 Then PonerContRegIndicador lblIndicador, adodc1, CadB
    
    CargaForaGrid
    
    
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
                SituarData Me.adodc1, "codtrans=" & DBSet(CodigoActual, "T"), "", True
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    PrimeraVez = True

    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'el 1 es separadors
        .Buttons(2).Image = 1   'Buscar
        .Buttons(3).Image = 2   'Todos
        'el 4 i el 5 son separadors
        .Buttons(6).Image = 3   'Insertar
        .Buttons(7).Image = 4   'Modificar
        .Buttons(8).Image = 5   'Borrar
        'el 9 i el 10 son separadors
        .Buttons(11).Image = 10  'imprimir
        .Buttons(12).Image = 11  'Salir
    End With

    '## A mano
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    
    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT rtransporte.codtrans, rtransporte.nomtrans, rtransporte.matricula, rtransporte.taravehi, "
    CadenaConsulta = CadenaConsulta & " rtransporte.tipoirpf, " ' CASE rtransporte.tipoirpf WHEN 0 THEN ""Módulos"" WHEN 1 THEN ""E.D."" WHEN 2 THEN ""Entidad"" END, "
    CadenaConsulta = CadenaConsulta & " rtransporte.codforpa, rtransporte.codmacpro, rtransporte.codiva, "
    CadenaConsulta = CadenaConsulta & " rtransporte.niftrans, rtransporte.dirtrans, rtransporte.pobtrans, "
    CadenaConsulta = CadenaConsulta & " rtransporte.protrans, rtransporte.codpostal,  "
    CadenaConsulta = CadenaConsulta & " rtransporte.teltran1, rtransporte.movtrans, rtransporte.maitrans, "
    CadenaConsulta = CadenaConsulta & " rtransporte.codbanco, rtransporte.codsucur, rtransporte.digcontr, rtransporte.cuentaba, "
    CadenaConsulta = CadenaConsulta & " rtransporte.sefactura, rtransporte.contador, rtransporte.codtraba, rtransporte.esfacttrainterna, rtransporte.iban  "
    CadenaConsulta = CadenaConsulta & " FROM rtransporte "
    CadenaConsulta = CadenaConsulta & " WHERE 1 = 1 "
    '************************************************************************
    
    ConexionConta
    
    CargaCombo
    
    CadB = ""
    CargaGrid
        
    ' Para el chivato
    Set dbAriagro = New BaseDatos
    dbAriagro.abrir_MYSQL vConfig.SERVER, vUsu.CadenaConexion, vConfig.User, vConfig.password

'    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'        BotonAnyadir
'    Else
'        PonerModo 2
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
    If Modo = 4 Then TerminaBloquear
    
    ' chivato
    Set dbAriagro = Nothing
    
    If Not vSeccion Is Nothing Then
        vSeccion.CerrarConta
        Set vSeccion = Nothing
    End If
    
End Sub

Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    txtAux1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codiva
    FormateaCampo txtAux1(indice)
    txtAux2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nomiva
End Sub

Private Sub frmFPa_DatoSeleccionado(CadenaSeleccion As String)
    txtAux1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codforpa
    FormateaCampo txtAux1(indice)
    txtAux2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nomforpa
End Sub

Private Sub frmTIva_DatoSeleccionado(CadenaSeleccion As String)
'Mantenimiento Tipo de iva
    txtAux1(6).Text = RecuperaValor(CadenaSeleccion, 1) 'codiva
    FormateaCampo txtAux1(6)
    txtAux2(6).Text = RecuperaValor(CadenaSeleccion, 2) 'nomiva
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
    txtAux1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo trabajador
    txtAux2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre trabajador
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    Select Case Index
        Case 0
            If vSeccion Is Nothing Then Exit Sub
            
            indice = 4
            Set frmFPa = New frmForpaConta
            frmFPa.DatosADevolverBusqueda = "0|1|"
            frmFPa.CodigoActual = txtAux1(indice)
        '    frmFpa.Conexion = cContaFacSoc
            frmFPa.Show vbModal
            Set frmFPa = Nothing
            PonerFoco txtAux1(indice)
        
        Case 1 'cuentas contables de y proveedor
            If vSeccion Is Nothing Then Exit Sub
            
            indice = 5
            Set frmCtas = New frmCtasConta
            frmCtas.NumDigit = 0
            frmCtas.DatosADevolverBusqueda = "0|1|"
            frmCtas.CodigoActual = txtAux1(indice).Text
            frmCtas.Show vbModal
            Set frmCtas = Nothing
            PonerFoco txtAux1(indice)
        
        
        Case 2 'codigo de iva
            If vSeccion Is Nothing Then Exit Sub
            
            indice = 6
            Set frmTIva = New frmTipIVAConta
            frmTIva.DeConsulta = True
            frmTIva.DatosADevolverBusqueda = "0|1|"
            frmTIva.CodigoActual = txtAux1(6).Text
            frmTIva.Show vbModal
            Set frmTIva = Nothing
            PonerFoco txtAux1(6)
            
            
        Case 3 ' codigo de trabajador si es tractorista
            indice = 1
            Set frmTra = New frmManTraba
            frmTra.DatosADevolverBusqueda = "0|2|"
            frmTra.Show vbModal
            Set frmTra = Nothing
            PonerFoco txtAux(indice)
     End Select

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
        Case 2
                mnBuscar_Click
        Case 3
                mnVerTodos_Click
        Case 6
                mnNuevo_Click
        Case 7
                mnModificar_Click
        Case 8
                mnEliminar_Click
        Case 11
                'MsgBox "Imprimir...under construction"
                mnImprimir_Click
        Case 12
                mnSalir_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String)
    Dim Sql As String
    Dim tots As String
    
'    adodc1.ConnectionString = Conn
    If vSQL <> "" Then
        Sql = CadenaConsulta & " AND " & vSQL
    Else
        Sql = CadenaConsulta
    End If
    '********************* canviar el ORDER BY *********************++
    Sql = Sql & " ORDER BY rtransporte.codtrans"
    '**************************************************************++
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, Sql, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "S|txtAux(0)|T|Código|1000|;S|txtAux(1)|T|Descripción|2850|;S|txtAux(2)|T|Matrícula|1000|;S|txtAux(3)|T|Tara Veh.|1000|;"
    tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
    tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;"
    tots = tots & "N||||0|;"
'    tots = tots & "N||||0|;S|Combo1(0)|C|Tipo IRPF|1200|;"
'    tots = tots & "S|txtaux(4)|T|F.Pago|800|;S|cmdAux(0)|B|||;"
'    tots = tots & "S|txtaux(5)|T|Cta.Prov.|1200|;S|cmdAux(1)|B|||;"
'    tots = tots & "S|txtaux(6)|T|Iva|800|;S|cmdAux(2)|B|||;"
    
    arregla tots, DataGrid1, Me
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgLeft
    DataGrid1.Columns(5).Alignment = dbgLeft
    DataGrid1.Columns(6).Alignment = dbgLeft
    DataGrid1.Columns(7).Alignment = dbgLeft
    DataGrid1.Columns(8).Alignment = dbgLeft
    

    If (Not adodc1.Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
        If Not vSeccion Is Nothing Then
        
            CargaForaGrid
        
            txtAux2(4).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", adodc1.Recordset!Codforpa, "N")
            If DBLet(adodc1.Recordset!codmacpro, "T") <> "" Then
                txtAux2(5).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", adodc1.Recordset!codmacpro, "T")
            End If
            
            txtAux2(6).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", CStr(adodc1.Recordset!CodIva), "N")
        End If
    Else
        txtAux2(4).Text = ""
        txtAux2(5).Text = ""
        txtAux2(6).Text = ""
    End If

'   DataGrid1.Columns(2).Alignment = dbgRight
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
'        Case 0
'            PonerFormatoEntero txtAux(Index)
        Case 0, 1, 2
            txtAux(Index).Text = UCase(txtAux(Index).Text)
'        Case 2
'            If Modo = 1 Then Exit Sub
'            PonerFormatoDecimal txtAux(Index), 4
            
        Case 3
            PonerFormatoEntero txtAux(Index)
            
            
        Case 4 ' FORMA DE PAGO
            If vSeccion Is Nothing Then Exit Sub
            
            If txtAux(Index).Text = "" Then Exit Sub
            
            If txtAux(Index).Text <> "" Then txtAux2(Index).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtAux(Index).Text, "N")
            If txtAux2(Index).Text = "" Then
                MsgBox "Forma de Pago  no existe en la contabilidad. Reintroduzca.", vbExclamation
            End If

        Case 5 ' CUENTA CONTABLE ( banco, retencion y aportacion )
            If vSeccion Is Nothing Then Exit Sub
            
            If txtAux(Index).Text = "" Then Exit Sub
        
            If txtAux(Index).Text <> "" Then txtAux2(Index).Text = PonerNombreCuenta(txtAux(Index), 2)
            If txtAux2(Index).Text = "" Then
                MsgBox "Número de Cuenta contable no existe en la contabilidad. Reintroduzca.", vbExclamation
            End If
            
        Case 6 'iva
            If vSeccion Is Nothing Then Exit Sub
            
            If txtAux(Index).Text = "" Then Exit Sub
            
            txtAux2(6).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", txtAux(Index).Text, "N")
    End Select
    
End Sub

Private Function DatosOk() As Boolean
'Dim Datos As String
Dim b As Boolean
Dim Sql As String
Dim Mens As String
Dim cta As String
Dim cadMen As String

    b = CompForm(Me)
    If Not b Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
         If ExisteCP(txtAux(0)) Then b = False
    End If
    
    '[Monica]09/03/2015: visualizamos las matriculas del nif introducido
    If Modo = 3 Then VisualizarMatriculas
    
'++[Monica] 24/10/2009 comprobamos que la cuenta CCC sea correcta
    If b And (Modo = 3 Or Modo = 4) Then
        If txtAux1(15).Text = "" Or txtAux1(16).Text = "" Or txtAux1(17).Text = "" Or txtAux1(18).Text = "" Then
            txtAux1(15).Text = ""
            txtAux1(16).Text = ""
            txtAux1(17).Text = ""
            txtAux1(18).Text = ""
        Else
            cta = Format(txtAux1(15).Text, "0000") & Format(txtAux1(16).Text, "0000") & Format(txtAux1(17).Text, "00") & Format(txtAux1(18).Text, "0000000000")
            If Val(ComprobarCero(cta)) = 0 Then
                cadMen = "El transportista no tiene asignada cuenta bancaria."
                MsgBox cadMen, vbExclamation
            End If
            If Not Comprueba_CC(cta) Then
                cadMen = "La cuenta bancaria del transportista no es correcta."
                MsgBox cadMen, vbExclamation
                b = False
            Else
'                '[Monica]20/11/2013: añadimos el tema de la comprobacion del IBAN
'                If Not Comprueba_CC_IBAN(cta, Text1(42).Text) Then
'                    cadMen = "La cuenta IBAN del cliente no es correcta. ¿ Desea continuar ?."
'                    If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'                        b = True
'                    Else
'                        PonerFoco Text1(42)
'                        b = False
'                    End If
'                End If

'       sustituido por lo de David
                BuscaChekc = ""
                If Me.txtAux1(2).Text <> "" Then BuscaChekc = Mid(txtAux1(2).Text, 1, 2)
                    
                If DevuelveIBAN2(BuscaChekc, cta, cta) Then
                    If Me.txtAux1(2).Text = "" Then
                        If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.txtAux1(2).Text = BuscaChekc & cta
                    Else
                        If Mid(txtAux1(2).Text, 3) <> cta Then
                            cta = "Calculado : " & BuscaChekc & cta
                            cta = "Introducido: " & Me.txtAux1(2).Text & vbCrLf & cta & vbCrLf
                            cta = "Error en codigo IBAN" & vbCrLf & cta & "Continuar?"
                            If MsgBox(cta, vbQuestion + vbYesNo) = vbNo Then
                                PonerFoco txtAux1(2)
                                b = False
                            End If
                        End If
                    End If
                End If
                
            End If
        End If
    End If
'++
    
    
    
    
    
    DatosOk = b
End Function

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub


Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "rtransporte"
        .Informe2 = "rManTranspor.rpt"
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
        '.cadTodosReg = "{itinerar.codempre} = " & codEmpre
        ' *** repasar si li pose ordre o no ****
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|pOrden={rtransporte.codtrans}|"
        '.OtrosParametros2 = "pEmpresa='" & vEmpresa.NomEmpre & "'|"
        ' *** posar el nº de paràmetres que he posat en OtrosParametros2 ***
        '.NumeroParametros2 = 1
        .NumeroParametros2 = 2
        ' ******************************************************************
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = False
        .SubInformeConta = ""
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

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
        If (Modo = 0 Or Modo = 2) Then Unload Me
    End If
End Sub

Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 0 To Combo1.Count - 1
        Combo1(i).Clear
    Next i
    
    'tipo irpf
    Combo1(0).AddItem "Módulo Agrario"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "E.D."
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Entidad"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    Combo1(0).AddItem "Modulo Transportista"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 3

End Sub

Private Sub ConexionConta()
    
    If vSeccion Is Nothing Then
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
            vSeccion.AbrirConta
        End If
    Else
        ' si el objeto existia: cerramos la conexion y volvemos crearlo
        vSeccion.CerrarConta
        Set vSeccion = Nothing
        
        
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
            vSeccion.AbrirConta
        End If
    End If
End Sub


Private Sub CargaForaGrid()
Dim i As Integer

    If DataGrid1.Columns.Count <= 2 Then Exit Sub
    
    ' *** posar als camps de fora del grid el valor de la columna corresponent ***
    
    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux(2).Text = DataGrid1.Columns(2).Text
    txtAux(3).Text = DataGrid1.Columns(3).Text
    
    For i = 4 To 7
        txtAux1(i).Text = DataGrid1.Columns(i + 1).Text
    Next i
    txtAux1(14).Text = DataGrid1.Columns(9).Text
    txtAux1(13).Text = DataGrid1.Columns(10).Text
    txtAux1(12).Text = DataGrid1.Columns(11).Text
    txtAux1(10).Text = DataGrid1.Columns(12).Text
    txtAux1(8).Text = DataGrid1.Columns(13).Text
    txtAux1(9).Text = DataGrid1.Columns(14).Text
    txtAux1(11).Text = DataGrid1.Columns(15).Text
    For i = 15 To 18
        txtAux1(i).Text = DataGrid1.Columns(i + 1).Text
    Next i
    
    If DataGrid1.Columns(20).Text <> "" Then
        Me.chkAbonos(0).Value = DataGrid1.Columns(20).Text
    End If
    
    If DataGrid1.Columns(23).Text <> "" Then
        Me.chkAbonos(1).Value = DataGrid1.Columns(23).Text
    End If
    
    txtAux1(0).Text = DataGrid1.Columns(21).Text ' contador de factura de transporte
    txtAux1(1).Text = DataGrid1.Columns(22).Text ' codigo de trabajador asociado si lo tiene
    
    txtAux1(2).Text = DataGrid1.Columns(24).Text ' iban
    
    txtAux2(1).Text = ""
    If txtAux1(1).Text <> "" Then
        txtAux2(1).Text = PonerNombreDeCod(txtAux1(1), "straba", "nomtraba", "codtraba", "N")
    End If
    
    If Not vSeccion Is Nothing Then
        txtAux2(4).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", adodc1.Recordset!Codforpa, "N")
        txtAux2(5).Text = ""
        If DBLet(adodc1.Recordset!codmacpro, "T") <> "" Then
            txtAux2(5).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", adodc1.Recordset!codmacpro, "T")
        End If
        
        txtAux2(6).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", CStr(adodc1.Recordset!CodIva), "N")
    End If
            
    PosicionarCombo Combo1(0), Me.adodc1.Recordset!TipoIRPF

End Sub

Private Sub txtAux1_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux1(Index)
End Sub

Private Sub txtAux1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub txtAux1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtAux1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Sql As String

    If Not PerderFocoGnral(txtAux1(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0 ' contador de factura
            PonerFormatoEntero txtAux1(Index)
            
        Case 1 ' codigo de trabajador
            If PonerFormatoEntero(txtAux1(Index)) Then
                txtAux2(Index).Text = PonerNombreDeCod(txtAux1(Index), "straba", "nomtraba")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe el Trabajador: " & txtAux1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmTra = New frmManTraba
                        frmTra.DatosADevolverBusqueda = "0|1|"
                        txtAux1(Index).Text = ""
                        TerminaBloquear
                        frmTra.Show vbModal
                        Set frmTra = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1
                    Else
                        txtAux1(Index).Text = ""
                    End If
                    PonerFoco txtAux1(Index)
                End If
            End If
            
        Case 3
            PonerFormatoEntero txtAux1(Index)
            
        Case 4 ' FORMA DE PAGO
            If vSeccion Is Nothing Then Exit Sub
            
            If txtAux1(Index).Text = "" Then Exit Sub
            
            If txtAux1(Index).Text <> "" Then txtAux2(Index).Text = DevuelveDesdeBDNew(cConta, "sforpa", "nomforpa", "codforpa", txtAux1(Index).Text, "N")
            If txtAux2(Index).Text = "" Then
                MsgBox "Forma de Pago  no existe en la contabilidad. Reintroduzca.", vbExclamation
            Else
                PonerFormatoEntero txtAux1(Index)
            End If

        Case 5 ' CUENTA CONTABLE ( banco, retencion y aportacion )
            If vSeccion Is Nothing Then Exit Sub
            
            If txtAux1(Index).Text = "" Then Exit Sub
        
            If txtAux1(Index).Text <> "" Then txtAux2(Index).Text = PonerNombreCuenta(txtAux1(Index), 2)
            If txtAux2(Index).Text = "" Then
                MsgBox "Número de Cuenta contable no existe en la contabilidad. Reintroduzca.", vbExclamation
            End If
            
        Case 6 'iva
            If vSeccion Is Nothing Then Exit Sub
            
            If txtAux1(Index).Text = "" Then Exit Sub
            
            txtAux2(6).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", txtAux1(Index).Text, "N")
            
        Case 15, 16 ' entidad y sucursal
            PonerFormatoEntero txtAux1(Index)
        
        Case 2 ' codigo de iban
            txtAux1(Index).Text = UCase(txtAux1(Index).Text)
            
        Case 7 ' nif del transportista
            '[Monica]09/03/2015: cuando me dan el nif sacamos una ventana con todas las matriculas y contadores
            VisualizarMatriculas
            
    End Select
    
    '[Monica]: calculo del iban si no lo ponen
    If Index = 15 Or Index = 16 Or Index = 17 Or Index = 18 Then
        Dim cta As String
        Dim CC As String
        If txtAux1(15).Text <> "" And txtAux1(16).Text <> "" And txtAux1(17).Text <> "" And txtAux1(18).Text <> "" Then
            
            cta = Format(txtAux1(15).Text, "0000") & Format(txtAux1(16).Text, "0000") & Format(txtAux1(17).Text, "00") & Format(txtAux1(18).Text, "0000000000")
            If Len(cta) = 20 Then
    '        Text1(42).Text = Calculo_CC_IBAN(cta, Text1(42).Text)
    
                If txtAux1(2).Text = "" Then
                    'NO ha puesto IBAN
                    If DevuelveIBAN2("ES", cta, cta) Then txtAux1(2).Text = "ES" & cta
                Else
                    CC = CStr(Mid(txtAux1(2).Text, 1, 2))
                    If DevuelveIBAN2(CStr(CC), cta, cta) Then
                        If Mid(txtAux1(2).Text, 3) <> cta Then
                            
                            MsgBox "Codigo IBAN distinto del calculado [" & CC & cta & "]", vbExclamation
                        End If
                    End If
                End If
            End If
        End If
    End If

   

End Sub

Private Sub VisualizarMatriculas()
Dim Sql As String
Dim Nregs As Integer

    Sql = "select count(*) from rtransporte where niftrans = " & DBSet(txtAux1(7).Text, "T")
    Nregs = TotalRegistros(Sql)
    If (Nregs > 0 And Modo = 3) Then
        Set frmMens = New frmMensajes
        frmMens.OpcionMensaje = 60
        frmMens.Label4(4).Caption = "Matrículas del tranportista NIF: " & txtAux1(7).Text
        frmMens.cadWHERE = "niftrans = " & DBSet(txtAux1(7).Text, "T")
        frmMens.Show vbModal
        Set frmMens = Nothing
    End If
End Sub




Private Sub ModificarDatosCuentaContable()
Dim Sql As String
Dim Cad As String

    On Error GoTo eModificarDatosCuentaContable

'    NombreAnt = txtAux(1).Text
'    IbanAnt = txtaux1(2).Text
'    BancoAnt = txtaux1(15).Text
'    SucurAnt = txtaux1(16).Text
'    DigitoAnt = txtaux1(17).Text
'    CuentaAnt = txtaux1(18).Text
'
'    DirecAnt = txtaux1(14).Text
'    cPostalAnt = txtaux1(10).Text
'    PoblaAnt = txtaux1(13).Text
'    ProviAnt = txtaux1(12).Text
'    NifAnt = txtaux1(7).Text
'    EMaiAnt = txtaux1(11).Text

    '[Monica]26/03/2015: antes no se grababa la forma de pago en la cuenta de cliente
    forpaant = txtAux1(4).Text


    If txtAux(1).Text <> NombreAnt Or txtAux1(15).Text <> BancoAnt Or txtAux1(16).Text <> SucurAnt Or txtAux1(17).Text <> DigitoAnt Or txtAux1(18).Text <> CuentaAnt Or _
       DirecAnt <> txtAux1(14).Text Or cPostalAnt <> txtAux1(10).Text Or PoblaAnt <> txtAux1(13).Text Or ProviAnt <> txtAux1(12).Text Or NifAnt <> txtAux1(7).Text Or _
       forpaant <> txtAux1(4).Text Or _
       EMaiAnt <> txtAux1(11).Text Or _
       IbanAnt <> txtAux1(2).Text Then

        Cad = "Se han producido cambios en datos del Transportista. " '& vbCrLf

'        If NombreAnt <> Text1(2).Text Then Cad = Cad & " Nombre,"
'        If DirecAnt <> Text1(4).Text Then Cad = Cad & " Direccion,"
'        If cPostalAnt <> Text1(5).Text Then Cad = Cad & " CPostal,"
'        If PoblaAnt <> Text1(18).Text Then Cad = Cad & " Población,"
'        If ProviAnt <> Text1(22).Text Then Cad = Cad & " Provincia,"
'        If NifAnt <> Text1(3).Text Then Cad = Cad & " NIF,"
''        If EMaiAnt <> Text1(12).Text Then Cad = Cad & " EMail,"
'        If BancoAnt <> Text1(1).Text Then Cad = Cad & " Banco,"
'        If SucurAnt <> Text1(28).Text Then Cad = Cad & " Sucursal,"
'        If DigitoAnt <> Text1(29).Text Then Cad = Cad & " Dig.Control,"
'        If CuentaAnt <> Text1(30).Text Then Cad = Cad & " Cuenta banco,"
'
'        Cad = Mid(Cad, 1, Len(Cad) - 1)

        Cad = Cad & vbCrLf & vbCrLf & "¿ Desea actualizarlos en la Contabilidad ?" & vbCrLf & vbCrLf

        If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then

            Sql = "update cuentas set nommacta = " & DBSet(Trim(txtAux(1).Text), "T")
            Sql = Sql & ", razosoci = " & DBSet(Trim(txtAux(1).Text), "T")
            Sql = Sql & ", dirdatos = " & DBSet(Trim(txtAux1(14).Text), "T")
            Sql = Sql & ", codposta = " & DBSet(Trim(txtAux1(10).Text), "T")
            Sql = Sql & ", despobla = " & DBSet(Trim(txtAux1(13).Text), "T")
            Sql = Sql & ", desprovi = " & DBSet(Trim(txtAux1(12).Text), "T")
            Sql = Sql & ", nifdatos = " & DBSet(Trim(txtAux1(7).Text), "T")
            Sql = Sql & ", maidatos = " & DBSet(Trim(txtAux1(11).Text), "T")

            '[Monica]26/03/2015: antes no grababamos la forma de pago de la cuenta
            Sql = Sql & ", forpa = " & DBSet(Trim(txtAux1(4).Text), "N", "S")


            If vParamAplic.ContabilidadNueva Then
                Dim vIban As String

                vIban = MiFormat(txtAux1(2).Text, "") & MiFormat(txtAux1(15).Text, "0000") & MiFormat(txtAux1(16).Text, "0000") & MiFormat(txtAux1(17).Text, "00") & MiFormat(txtAux1(18).Text, "0000000000")

                Sql = Sql & ", iban = " & DBSet(vIban, "T")
                Sql = Sql & ", codpais = 'ES' "
            Else
                Sql = Sql & ", entidad = " & DBSet(Trim(txtAux1(15).Text), "T", "S")
                Sql = Sql & ", oficina = " & DBSet(Trim(txtAux1(16).Text), "T", "S")
                Sql = Sql & ", cc = " & DBSet(Trim(txtAux1(17).Text), "T", "S")
                Sql = Sql & ", cuentaba = " & DBSet(Trim(txtAux1(18).Text), "T", "S")

                '[Monica]22/11/2013: tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    Sql = Sql & ", iban = " & DBSet(Trim(txtAux1(2).Text), "T", "S")
                End If
            End If
            Sql = Sql & " where codmacta = " & DBSet(Trim(txtAux1(5).Text), "T")

            ConnConta.Execute Sql

'            MsgBox "Datos de Cuenta modificados correctamente.", vbExclamation

        End If
    End If


    '[Monica]30/08/2013: modificamos los datos de tesoreria sobre los cobros y pagos pendientes
    If txtAux1(15).Text <> BancoAnt Or txtAux1(16).Text <> SucurAnt Or txtAux1(17).Text <> DigitoAnt Or txtAux1(18).Text <> CuentaAnt _
        Or txtAux1(2).Text <> IbanAnt Or txtAux1(4).Text <> forpaant Then
        Cad = "Se han producido cambios en la Cta.Bancaria del transportista."
        Cad = Cad & vbCrLf & vbCrLf & "¿ Desea actualizar los Cobros y Pagos pendientes en Tesoreria ?" & vbCrLf & vbCrLf

        If HayCobrosPagosPendientes(txtAux1(5).Text) Then
            If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                If ActualizarCobrosPagosPdtes(txtAux1(5).Text, txtAux1(15).Text, txtAux1(16).Text, txtAux1(17).Text, txtAux1(18).Text, txtAux1(2).Text, txtAux1(4).Text) Then
'                    MsgBox "Datos en Tesoreria modificados correctamente.", vbExclamation
                End If
            End If
        End If
    End If

    Exit Sub

eModificarDatosCuentaContable:
    MuestraError Err.Number, "Modificar Datos Cuenta Contable", Err.Description
End Sub




