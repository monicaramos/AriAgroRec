VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPagoAnticiposCoopic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6495
   Icon            =   "frmPagoAnticiposCoopic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameHorasTrabajadas 
      Height          =   6555
      Left            =   0
      TabIndex        =   10
      Top             =   45
      Width           =   6435
      Begin VB.Frame FramePago 
         Height          =   1185
         Left            =   270
         TabIndex        =   23
         Top             =   4320
         Width           =   6000
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
            Index           =   2
            Left            =   2655
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Tag             =   "Tipo|N|N|||straba|codsecci||N|"
            Top             =   675
            Width           =   1665
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
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
            Left            =   2325
            Locked          =   -1  'True
            TabIndex        =   24
            Text            =   "Text5"
            Top             =   270
            Width           =   3375
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
            Left            =   1395
            MaxLength       =   6
            TabIndex        =   6
            Top             =   270
            Width           =   870
         End
         Begin VB.Label Label3 
            Caption         =   "Concepto Transferencia "
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
            Left            =   135
            TabIndex        =   26
            Top             =   675
            Width           =   2415
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   0
            Left            =   1095
            MouseIcon       =   "frmPagoAnticiposCoopic.frx":000C
            MousePointer    =   4  'Icon
            ToolTipText     =   "Buscar banco"
            Top             =   270
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Banco "
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
            Index           =   0
            Left            =   135
            TabIndex        =   25
            Top             =   225
            Width           =   675
         End
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
         Index           =   16
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   2
         Top             =   2340
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
         Index           =   17
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   3
         Top             =   2745
         Width           =   1350
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Index           =   19
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Index           =   18
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text5"
         Top             =   1305
         Width           =   3375
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
         Index           =   18
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   0
         Top             =   1305
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
         Index           =   19
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1665
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
         Index           =   20
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   4
         Top             =   3420
         Width           =   1350
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
         Index           =   1
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "Tipo|N|N|||straba|codsecci||N|"
         Top             =   3870
         Width           =   1710
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
         Left            =   5010
         TabIndex        =   9
         Top             =   5910
         Width           =   1065
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
         Left            =   3885
         TabIndex        =   8
         Top             =   5910
         Width           =   1065
      End
      Begin MSComctlLib.ProgressBar Pb2 
         Height          =   240
         Left            =   405
         TabIndex        =   12
         Top             =   5580
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   4860
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1395
         Picture         =   "frmPagoAnticiposCoopic.frx":015E
         Top             =   2340
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   1395
         Picture         =   "frmPagoAnticiposCoopic.frx":01E9
         Top             =   2745
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   14
         Left            =   1440
         MouseIcon       =   "frmPagoAnticiposCoopic.frx":0274
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1305
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   15
         Left            =   1440
         MouseIcon       =   "frmPagoAnticiposCoopic.frx":03C6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1665
         Width           =   240
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
         Index           =   24
         Left            =   420
         TabIndex        =   22
         Top             =   2115
         Width           =   600
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
         Index           =   25
         Left            =   690
         TabIndex        =   21
         Top             =   2715
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
         Index           =   26
         Left            =   690
         TabIndex        =   20
         Top             =   2400
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
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
         Index           =   27
         Left            =   420
         TabIndex        =   19
         Top             =   1035
         Width           =   1065
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
         TabIndex        =   18
         Top             =   1680
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
         Index           =   29
         Left            =   690
         TabIndex        =   17
         Top             =   1320
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Anticipo"
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
         Index           =   30
         Left            =   405
         TabIndex        =   16
         Top             =   3105
         Width           =   1470
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   6
         Left            =   1380
         Picture         =   "frmPagoAnticiposCoopic.frx":0518
         Top             =   3420
         Width           =   240
      End
      Begin VB.Label Label8 
         Caption         =   "Secci�n "
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
         Left            =   405
         TabIndex        =   15
         Top             =   3870
         Width           =   885
      End
      Begin VB.Label Label7 
         Caption         =   "Pago Anticipos"
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
         Left            =   405
         TabIndex        =   11
         Top             =   405
         Width           =   5925
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPagoAnticiposCoopic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MONICA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public OpcionListado As Byte
    '==== Listados BASICOS ====
    '=============================
    ' 1 .- Pago de Recibos de valsur y alzira
    ' 2 .- Pago de Recibos de natural de monta�a
    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar n� oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(Cliente As String, Observaciones As String)

Private Conexion As Byte
'1.- Conexi�n a BD Ariges  2.- Conexi�n a BD Conta

Private HaDevueltoDatos As Boolean

Private WithEvents frmBan As frmBasico2 'Banco propio
Attribute frmBan.VB_VarHelpID = -1
Private WithEvents frmTra As frmManTraba 'mantenimiento de trabajadores
Attribute frmTra.VB_VarHelpID = -1

Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe
Private ConSubInforme As Boolean

Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'n� de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'C�digo para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String
Dim Repetir As Boolean

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
    
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
    
Dim cadSelect1 As String
Dim cadSelect2 As String
Dim cTabla As String
Dim Sql As String

    
    If Not DatosOK Then Exit Sub
    
    cadSelect = ""
               
               
    'D/H TRABAJADOR
    cDesde = Trim(txtCodigo(18).Text)
    cHasta = Trim(txtCodigo(19).Text)
    nDesde = txtNombre(18).Text
    nHasta = txtNombre(19).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{horas.codtraba}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHTrabajador=""") Then Exit Sub
    End If
            
    'D/H fecha
    cDesde = Trim(txtCodigo(16).Text)
    cHasta = Trim(txtCodigo(17).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{horas.fechahora}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
    End If
                       
    'Tipo de seccion
    AnyadirAFormula cadFormula, "{straba.codsecci} = " & Me.Combo1(1).ListIndex
    AnyadirAFormula cadSelect, "{straba.codsecci} = " & Me.Combo1(1).ListIndex
    
            
    tabla = "horas INNER JOIN straba ON horas.codtraba = straba.codtraba "
                       
    AnyadirAFormula cadFormula, "isnull({horas.fecharec})"
    AnyadirAFormula cadSelect, "horas.fecharec is null"
    
    AnyadirAFormula cadSelect, "horas.intconta = 0"
    
    
    '[Monica]08/02/2017: los que han trabajado y se dan de baja no se anticipan, se manda la nomina hasta el momento
    AnyadirAFormula cadSelect, "(straba.fechabaja is null or straba.fechabaja = '')"
    
    If vParamAplic.Cooperativa = 0 Then
        AnyadirAFormula cadSelect, "straba.codforpa in (select codforpa from forpago where forpago.tipoforp = 1)"
        
        '[Monica]20/04/2018: solo vamos a anticipar a los del banco que toque y los que no tengan categoria
        Dim SqlNue As String
        SqlNue = "(horas.codcateg in (select codcateg from rcategorias where codbanpr = " & DBSet(txtCodigo(0).Text, "N")
        SqlNue = SqlNue & ") or horas.codcateg is null or horas.codcateg = '')"
        
        AnyadirAFormula cadSelect, SqlNue
    End If
    
                       
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    Repetir = False
    
    If HayRegParaInforme(tabla, cadSelect) Then
        ProcesarCambiosCoopic (cadSelect)
    Else
        Repetir = True
        If MsgBox("�Desea repetir el �ltimo anticipo de esa fecha?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            RepetirProcesoCoopic
        End If
    End If
    
    cmdCancel_Click (0)
    
End Sub




Private Sub RepetirProcesoCoopic()
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim cad As String
Dim I As Integer
Dim HayReg As Integer
Dim B As Boolean
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Mens As String

Dim ImpHoras As Currency
Dim ImpHorasE As Currency
Dim ImpBruto As Currency
Dim IRPF As Currency
Dim SegSoc As Currency
Dim SegSoc1 As Currency
Dim Neto As Currency
Dim Bruto As Currency
Dim CuentaPropia As String
Dim CodigoOrden34 As String

Dim Neto34 As Currency
Dim Bruto34 As Currency
Dim Jornadas As Currency
Dim Diferencia As Currency
Dim BaseSegso As Currency
Dim Complemento As Currency
Dim TSegSoc As Currency
Dim TSegSoc1 As Currency
Dim Max As Long

Dim Sql5 As String
Dim RS5 As ADODB.Recordset

Dim IdContador As Long
Dim TieneEmbargo As String

On Error GoTo eRepetirProcesoCoopic
    
    BorrarTMP
    CrearTMP
        
    Sql = "select max(idcontador) from rrecibosnomina where fechahora = " & DBSet(txtCodigo(20).Text, "F") & " and hayembargo = 0 "
    IdContador = DevuelveValor(Sql)
    
    Sql = "select count(*) from rrecibosnomina where idcontador = " & DBSet(IdContador, "N") & " and hayembargo = 0"
    If TotalRegistros(Sql) = 0 Then
        Mens = "No hay anticipos, debe realizar el proceso."
        B = False
    Else
        If vParamAplic.Cooperativa = 0 Then
        
            Sql = "select * from rrecibosnomina where idcontador = " & DBSet(IdContador, "N")
            
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not Rs.EOF
            
                IncrementarProgres Pb2, 1
                Mens = "Calculando Importes" & vbCrLf & vbCrLf & "Trabajador: " & Rs!CodTraba & vbCrLf
                
                TieneEmbargo = DevuelveValor("select hayembargo from straba where codtraba = " & DBSet(Rs!CodTraba, "N"))
                If TieneEmbargo = "0" Then
                    Sql3 = "insert into tmpImpor (codtraba, importe) values ("
                    Sql3 = Sql3 & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(Rs!Neto34, "N") & ")"
                
                    conn.Execute Sql3
                End If
                    
                Rs.MoveNext
            Wend
            
            Set Rs = Nothing
        
        
            Dim vSeccion As CSeccion
            If vSeccion Is Nothing Then
                Set vSeccion = New CSeccion
                If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
                    vSeccion.AbrirConta
                End If
            End If
        
            '[Monica]22/11/2013: iban
            Sql = "select codbanco, codsucur, digcontr, cuentaba, codorden34, iban from banpropi where codbanpr = " & DBSet(txtCodigo(0).Text, "N")
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            CodigoOrden34 = ""
            
            If Rs.EOF Then
                cad = ""
            Else
                If IsNull(Rs!CodBanco) Then
                    cad = ""
                Else
                    '[Monica]22/11/2013: iban
                    cad = Format(Rs!CodBanco, "0000") & "|" & Format(DBLet(Rs!CodSucur, "T"), "0000") & "|" & DBLet(Rs!digcontr, "T") & "|" & Format(DBLet(Rs!CuentaBa, "T"), "0000000000") & "|" & DBLet(Rs!IBAN, "T") & "|"
                End If
                CodigoOrden34 = DBLet(Rs!codorden34, "T")
            End If
            
            Set Rs = Nothing
            
            CuentaPropia = cad
            
            '[Monica]02/02/2018: Catadau ha de generar el fichero
            If vEmpresa.AplicarNorma19_34Nueva = 1 Then
                If HayXML Then
                    B = GeneraFicheroNorma34SEPA_XML(vParam.CifEmpresa, CDate(txtCodigo(20).Text), CuentaPropia, "", "Pago N�mina", Combo1(2).ListIndex, CodigoOrden34)
                Else
                    B = GeneraFicheroNorma34SEPA(vParam.CifEmpresa, CDate(txtCodigo(20).Text), CuentaPropia, "", "Pago N�mina", Combo1(2).ListIndex, CodigoOrden34)
                End If
            Else
                B = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtCodigo(20).Text), CuentaPropia, 9, 0, "Pago N�mina", CodigoOrden34, Combo1(2).ListIndex)
            End If
            
            
            vSeccion.CerrarConta
            Set vSeccion = Nothing

            
        'antes
        '    b = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtcodigo(20).Text), CuentaPropia, 9, 0, "Pago N�mina", CodigoOrden34, Combo1(0).ListIndex)
        Else
            B = GeneraFicheroA3(IdContador, txtCodigo(20).Text)
        End If
    End If
    
    If B Then
        Mens = "Copiar Fichero"
        If vParamAplic.Cooperativa = 0 Then
            If CopiarFichero Then
                B = True
            Else
                B = False
            End If
        Else
            If CopiarFicheroA3("AnticipoA3.txt", txtCodigo(20).Text) Then
                B = True
            Else
                B = False
            End If
        End If
    End If

eRepetirProcesoCoopic:
    If Err.Number <> 0 Then
        Mens = Err.Description
        B = False
    End If
    If B Then
        MsgBox "Proceso realizado correctamente.", vbExclamation
        cmdCancel_Click (1)
    Else
        MsgBox "Error " & Mens, vbExclamation
    End If
End Sub


Private Function DireccionesOk(cTabla As String, cWhere As String) As Boolean
Dim Sql As String
Dim cadResult As String
Dim Rs As ADODB.Recordset

    On Error GoTo eDireccionesOk
    
    DireccionesOk = False

    Sql = "Select straba.* FROM " & cTabla & "  WHERE " & cWhere
    Sql = Sql & " and (domtraba is null or domtraba = '' or codpobla is null or codpobla = ''  or pobtraba is null or pobtraba is null or protraba is null or protraba = '') "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cadResult = ""
    While Not Rs.EOF
        cadResult = cadResult & DBLet(Rs!CodTraba) & ","
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    If cadResult <> "" Then
        cadResult = Mid(cadResult, 1, Len(cadResult) - 1)
    
        MsgBox "Los siguientes trabajadores no tienen la direcci�n correcta: " & vbCrLf & vbCrLf & cadResult, vbExclamation
    
    End If
    
    
    DireccionesOk = (cadResult = "")
    Exit Function
eDireccionesOk:
    MuestraError Err.Number, "Direcciones Correctas", Err.Description
End Function


Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub


Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Combo1(1).ListIndex = 0
        Combo1(2).ListIndex = 0
        PonerFocoCmb Combo1(1)
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
    

' ### [Monica] 09/11/2006    he sustituido el anterior
    For H = 14 To 15 'imgBuscar.Count - 1
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    
    'Ocultar todos los Frames de Formulario
    Me.FrameHorasTrabajadas.visible = False
    
    CargaCombo
        
    '###Descomentar
'    CommitConexion
    H = 5700
    W = 6435
    
    If vParamAplic.Cooperativa = 0 Then H = H + 1200
    
    FrameHorasTrabajadasVisible True, H, W
    indFrame = 0
    Me.cmdCancel(0).Cancel = True
        
    tabla = "horas"
        
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.Width = W + 70
    Me.Height = H + 350
    
    Me.Combo1(1).ListIndex = 0
    
    Me.FramePago.visible = (vParamAplic.Cooperativa = 0)
    Me.FramePago.Enabled = (vParamAplic.Cooperativa = 0)
    
    If vParamAplic.Cooperativa <> 0 Then
        Me.CmdAceptar(0).Top = Me.CmdAceptar(0).Top - 1200
        Me.cmdCancel(0).Top = Me.cmdCancel(0).Top - 1200
        Me.Pb2.Top = Me.Pb2.Top - 1200
    End If
    
    Pb2.visible = False
End Sub



Private Sub frmBan_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de banco propio
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtCodigo(CByte(imgFecha(2).Tag)).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 14, 15  'Banco propio
            AbrirFrmManTraba (Index)
        Case 0
            indCodigo = 0
            AbrirFrmManBanco (Index)
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub AbrirFrmManTraba(Indice As Integer)
    indCodigo = Indice + 4
    Set frmTra = New frmManTraba
    frmTra.DatosADevolverBusqueda = "0|2|"
    frmTra.Show vbModal
    Set frmTra = Nothing
End Sub


Private Sub imgFecha_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object
    Dim Indice As Integer

    Set frmC = New frmCal
    
    esq = imgFecha(Index).Left
    dalt = imgFecha(Index).Top
        
    Set obj = imgFecha(Index).Container
      
      While imgFecha(Index).Parent.Name <> obj.Name
            esq = esq + obj.Left
            dalt = dalt + obj.Top
            Set obj = obj.Container
      Wend
    
    menu = Me.Height - Me.ScaleHeight 'ac� tinc el heigth del men� i de la toolbar

    frmC.Left = esq + imgFecha(Index).Parent.Left + 30
    frmC.Top = dalt + imgFecha(Index).Parent.Top + imgFecha(Index).Height + menu - 40

    Select Case Index
        Case 2, 3, 6
            Indice = Index + 14
    End Select

    imgFecha(2).Tag = Indice '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtCodigo(Indice).Text <> "" Then frmC.NovaData = txtCodigo(Indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtCodigo(CByte(imgFecha(2).Tag)) '<===
    ' ********************************************
End Sub

Private Sub ListView1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
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
'15/02/2007
'    KEYpress KeyAscii
'ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 18: KEYBusqueda KeyAscii, 14 'trabajador desde
            Case 19: KEYBusqueda KeyAscii, 15 'trabajador hasta
            Case 2: KEYFecha KeyAscii, 16 'fecha desde
            Case 3: KEYFecha KeyAscii, 17 'fecha hasta
            Case 6: KEYFecha KeyAscii, 20 'fecha recibo
            
            Case 0: KEYBusqueda KeyAscii, 0 'banco
        
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
    imgFecha_Click (Indice)
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
            
        Case 16, 17, 20   'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 18, 19 ' trabajador
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "straba", "nomtraba", "codtraba", "N")
            
        Case 0 ' banco
            If txtCodigo(Index).Text <> "" Then PonerFormatoEntero txtCodigo(Index)
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "banpropi", "nombanpr", "codbanpr", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
        
    End Select
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String

    'Llamamos a al form
    cad = ""
    Conexion = cAgro    'Conexi�n a BD: Ariges
'    Select Case OpcionListado
'        Case 7 'Traspaso de Almacenes
'            cad = cad & "N� Trasp|scatra|codtrasp|N|0000000|40�Almacen Origen|scatra|almaorig|N|000|20�Almacen Destino|scatra|almadest|N|000|20�Fecha|scatra|fechatra|F||20�"
'            Tabla = "scatra"
'            titulo = "Traspaso Almacenes"
'        Case 8 'Movimientos de Almacen
'            cad = cad & "N� Movim.|scamov|codmovim|N|0000000|40�Almacen|scamov|codalmac|N|000|30�Fecha|scamov|fecmovim|F||30�"
'            Tabla = "scamov"
'            titulo = "Movimientos Almacen"
'        Case 9, 12, 13, 14, 15, 16, 17 '9: Movimientos Articulos
'                   '12: Inventario Articulos
'                   '14:Actualizar Diferencias de Stock Inventariado
'                   '16: Listado Valoracion stock inventariado
'            cad = cad & "C�digo|sartic|codartic|T||30�Denominacion|sartic|nomartic|T||70�"
'            Tabla = "sartic"
'            titulo = "Articulos"
'    End Select
          
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vtabla = tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
        'frmB.vDevuelve = "0|1|"
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = Titulo
        frmB.vSelElem = 0
'        frmB.vConexionGrid = Conexion
'        frmB.vBuscaPrevia = 1
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
'        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub FrameHorasTrabajadasVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de horas trabajadas
    Me.FrameHorasTrabajadas.visible = visible
    If visible = True Then
        Me.FrameHorasTrabajadas.Top = -90
        Me.FrameHorasTrabajadas.Left = 0
        Me.FrameHorasTrabajadas.Height = H
        Me.FrameHorasTrabajadas.Width = W
        W = Me.FrameHorasTrabajadas.Width
        H = Me.FrameHorasTrabajadas.Height
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
        .Titulo = cadTitulo
        .EnvioEMail = False
        .NombreRPT = cadNombreRPT
        .Opcion = 0
        .ConSubInforme = True
        .Show vbModal
    End With
End Sub

Private Function PonerGrupo(numGrupo As Byte, cadgrupo As String) As Byte
Dim campo As String
Dim nomCampo As String

    campo = "pGroup" & numGrupo & "="
    nomCampo = "pGroup" & numGrupo & "Name="
    PonerGrupo = 0

    Select Case cadgrupo
'        Case "Codigo"
'            cadParam = cadParam & campo & "{" & Tabla & ".codclien}" & "|"
'            cadParam = cadParam & nomcampo & " {" & "scoope" & ".nomcoope}" & "|"
'            cadParam = cadParam & "pTitulo1" & "=""C�digo""" & "|"
'            numParam = numParam + 3
'
'        Case "Alfabetico"
'            cadParam = cadParam & campo & "{" & Tabla & ".tipsocio}" & "|"
'            cadParam = cadParam & nomcampo & " {" & "tiposoci" & ".nomtipso}" & "|"
'            cadParam = cadParam & "pTitulo1" & "=""Colectivo""" & "|"
'            numParam = numParam + 3
'
        
        'Informe de variedades
        Case "Clase"
            CadParam = CadParam & campo & "{" & tabla & ".codclase}" & "|"
            CadParam = CadParam & nomCampo & " {" & "clases" & ".nomclase}" & "|"
            CadParam = CadParam & "pTitulo1" & "=""Producto""" & "|"
            numParam = numParam + 3
            
        Case "Producto"
            CadParam = CadParam & campo & "{" & tabla & ".codprodu}" & "|"
            CadParam = CadParam & nomCampo & " {" & "productos" & ".nomprodu}" & "|"
            CadParam = CadParam & "pTitulo1" & "=""Clase""" & "|"
            numParam = numParam + 3

        'Informe de calibres
        Case "Variedad"
            CadParam = CadParam & campo & "{" & tabla & ".codvarie}" & "|"
            CadParam = CadParam & nomCampo & " {" & "variedades" & ".nomvarie}" & "|"
            CadParam = CadParam & "pTitulo1" & "=""Variedad""" & "|"
            numParam = numParam + 3
            
        Case "Calibre"
            CadParam = CadParam & campo & "{" & tabla & ".codcalib}" & "|"
            CadParam = CadParam & nomCampo & " {" & "calibres" & ".nomcalib}" & "|"
            CadParam = CadParam & "pTitulo1" & "=""Calibre""" & "|"
            numParam = numParam + 3
            
'        'Informe de Horas Trabajadas
'        Case "Trabajador"
'            cadParam = cadParam & campo & "{" & Tabla & ".codtraba}" & "|"
'            cadParam = cadParam & nomcampo & " {" & "straba" & ".nomtraba}" & "|"
'            cadParam = cadParam & "pTitulo1" & "=""Fecha""" & "|"
'            numParam = numParam + 3
'
'        Case "Fecha"
'            cadParam = cadParam & "pGroup1=" & "{" & Tabla & ".fechahora}" & "|"
'            cadParam = cadParam & "pGroup1Name=" & " {" & "horas" & ".fechahora}" & "|"
'            cadParam = cadParam & "pTitulo1" & "=""Trabajadores""" & "|"
'            numParam = numParam + 3
        

End Select

End Function

Private Function PonerOrden(cadgrupo As String) As Byte
Dim campo As String
Dim nomCampo As String

    PonerOrden = 0

    Select Case cadgrupo
        Case "Codigo"
            CadParam = CadParam & "Orden" & "= {" & tabla
            Select Case OpcionListado
                Case 10
                    CadParam = CadParam & ".codclien}|"
                Case 11
                    CadParam = CadParam & ".codprove}|"
            End Select
            Tipo = "C�digo"
        Case "Alfab�tico"
            CadParam = CadParam & "Orden" & "= {" & tabla
            Select Case OpcionListado
                Case 10
                    CadParam = CadParam & ".nomclien}|"
                Case 11
                    CadParam = CadParam & ".nomprove}|"
            End Select
            Tipo = "Alfab�tico"
    End Select
    
    numParam = numParam + 1

End Function

Private Sub AbrirFrmManBanco(Indice As Integer)
    Set frmBan = New frmBasico2
    
    AyudaBancosCom frmBan, txtCodigo(indCodigo)
    
    Set frmBan = Nothing
    
    PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
        .OtrosParametros = CadParam
        .NumeroParametros = numParam
        .Opcion = OpcionListado
        .Show vbModal
    End With
    
    Unload Me
End Sub

Private Sub AbrirEMail()
    If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
End Sub


' ********* si n'hi han combos a la cap�alera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim I As Integer

' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
'    For I = 0 To Combo1.Count - 1
'        Combo1(I).Clear
'    Next I

    Combo1(1).Clear
    
    Combo1(1).AddItem "Campo"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Almac�n"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    
    Combo1(2).Clear
    
    Combo1(2).AddItem "N�mina"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 0
    Combo1(2).AddItem "Pensi�n"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 1
    Combo1(2).AddItem "Otros Conceptos"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 2
    
End Sub

Private Function DatosOK() As Boolean
Dim B As Boolean
Dim Sql As String
'Dim Datos As String

    On Error GoTo EDatosOK

    B = True

    
    If txtCodigo(20).Text = "" Then
        MsgBox "Debe introducir una Fecha de Anticipo.", vbExclamation
        txtCodigo(20).Text = ""
        PonerFoco txtCodigo(20)
        B = False
    End If
    
    '[Monica]05/02/2018: para Catadau metemos los datos de banco pq hacen transferencia
    If vParamAplic.Cooperativa = 0 Then
        If txtCodigo(0).Text = "" Then
            MsgBox "Debe introducir un Banco.", vbExclamation
            txtCodigo(0).Text = ""
            PonerFoco txtCodigo(0)
            B = False
        End If
    
        If Combo1(2).ListIndex = -1 Then
            MsgBox "Debe introducir un Concepto de Transferencia", vbExclamation
            PonerFocoCmb Combo1(2)
            B = False
        End If
    End If
    
    
    DatosOK = B
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function ActualizarRegistros(tabla As String, cWhere As String) As Boolean
Dim Sql As String
    On Error GoTo eActualizarRegistros
    
    ActualizarRegistros = False
    
    cWhere = QuitarCaracterACadena(cWhere, "{")
    cWhere = QuitarCaracterACadena(cWhere, "}")
    cWhere = QuitarCaracterACadena(cWhere, "_1")

    Sql = "update horas, straba set fecharec = " & DBSet(txtCodigo(20).Text, "F")
    Sql = Sql & " where " & cWhere
    Sql = Sql & " and horas.codtraba = straba.codtraba"
'    (codtraba, fechahora) in (select horas.codtraba, horas.fechahora from " & tabla & " where " & cWhere & ")"
    
    conn.Execute Sql
        
    ActualizarRegistros = True
    
    Exit Function

eActualizarRegistros:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error en la actualizacion de Registros" & vbCrLf & Err.Description
    End If
End Function

Public Sub BorrarTMP()
On Error Resume Next

    conn.Execute " DROP TABLE IF EXISTS tmpImpor;"
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Function CrearTMP() As Boolean
'Crea una temporal donde inserta la clave primaria de las
'facturas seleccionadas para facturar y trabaja siempre con ellas
Dim Sql As String
    
    On Error GoTo ECrear
    
    CrearTMP = False
    
    Sql = "CREATE TEMPORARY TABLE tmpImpor ( "
    Sql = Sql & "codtraba int(6) unsigned NOT NULL default '0',"
    Sql = Sql & "importe decimal(12,2)  NOT NULL default '0')"
    
    conn.Execute Sql
     
    CrearTMP = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMP = False
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmpImpor;"
        conn.Execute Sql
    End If
End Function

Public Function CopiarFichero() As Boolean
Dim nomFich As String

On Error GoTo ecopiarfichero

    CopiarFichero = False
    ' abrimos el commondialog para indicar donde guardarlo
'    Me.CommonDialog1.InitDir = App.path

    Me.CommonDialog1.DefaultExt = "txt"
    
    CommonDialog1.Filter = "Archivos txt|txt|"
    CommonDialog1.FilterIndex = 1
    
    ' copiamos el primer fichero
    CommonDialog1.FileName = "norma34.txt"
    Me.CommonDialog1.ShowSave
    
    If CommonDialog1.FileName <> "" Then
        FileCopy App.Path & "\norma34.txt", CommonDialog1.FileName
    End If
    
    CopiarFichero = True
    Exit Function

ecopiarfichero:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    End If
    Err.Clear
End Function


Private Sub ProcesarCambiosCoopic(cadWHERE As String)
Dim Sql As String
Dim Sql2 As String
Dim Sql3 As String
Dim cad As String
Dim I As Integer
Dim HayReg As Integer
Dim B As Boolean
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Mens As String

Dim ImpHoras As Currency
Dim ImpHorasE As Currency
Dim ImpBruto As Currency
Dim ImpBruto2 As Currency
Dim IRPF As Currency
Dim SegSoc As Currency
Dim SegSoc1 As Currency
Dim Neto As Currency
Dim Bruto As Currency
Dim Retencion As Currency
Dim CuentaPropia As String
Dim CodigoOrden34 As String

Dim Neto34 As Currency
Dim Bruto34 As Currency
Dim Jornadas As Currency
Dim Diferencia As Currency
Dim BaseSegso As Currency
Dim Complemento As Currency
Dim TSegSoc As Currency
Dim TSegSoc1 As Currency
Dim Max As Long

Dim Sql5 As String
Dim RS5 As ADODB.Recordset

Dim Dias As Long
Dim Anticipo As Currency

On Error GoTo eProcesarCambiosCoopic
    
    BorrarTMP
    CrearTMP

    conn.BeginTrans
    
    If cadWHERE <> "" Then
        cadWHERE = QuitarCaracterACadena(cadWHERE, "{")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "}")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "_1")
    End If
        
    Sql = "select count(distinct horas.codtraba) from (horas inner join straba on horas.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Pb2.visible = True
    CargarProgres Pb2, Rs.Fields(0).Value
    
    Rs.Close
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
    Sql3 = "select max(idcontador) from rrecibosnomina"
    Max = DevuelveValor(Sql3) + 1
    
    Sql = "select horas.codtraba, 0, sum(if(horasdia is null,0,horasdia)), sum(if(compleme is null,0,compleme)), sum(if(penaliza is null,0,penaliza)), sum(if(importe is null,0,importe)) from (horas inner join straba on horas.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWHERE
    Sql = Sql & " group by horas.codtraba, 2 "
    Sql = Sql & " order by 1, 2"
        
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Dim AntCodTraba As Long
    Dim ActCodTraba As Long
    Dim TIRPF As Currency
    Dim TImpbruto As Currency
    Dim TImpBruto2 As Currency
    Dim TRetencion As Currency
    Dim TNeto34 As Currency
    Dim TSegSo As Currency
    
    TIRPF = 0
    TImpbruto = 0
    TImpBruto2 = 0
    TRetencion = 0
    TNeto34 = 0
    TSegSo = 0
    
    If Not Rs.EOF Then
        AntCodTraba = DBLet(Rs!CodTraba, "N")
        ActCodTraba = AntCodTraba
        Sql2 = "select salarios.*, straba.dtoreten, straba.dtosegso, straba.dtosirpf, straba.pluscapataz, straba.hayembargo from salarios, straba where straba.codtraba = " & DBSet(Rs!CodTraba, "N")
        Sql2 = Sql2 & " and salarios.codcateg = straba.codcateg "
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    End If
    
    HayReg = False
    
    While Not Rs.EOF
        HayReg = True
        
        ActCodTraba = DBLet(Rs!CodTraba, "N")
        
        If AntCodTraba <> ActCodTraba Then
            IncrementarProgres Pb2, 1
            Mens = "Calculando Importes" & vbCrLf & vbCrLf & "Trabajador: " & AntCodTraba & vbCrLf
            
            
            '[Monica]23/03/2016: si el importe es negativo no entra
            If TNeto34 >= 0 Then
        
                '[Monica]25/05/2018: anticipos pendientes de descuento
                Sql = "update horasanticipos set descontado = 1, fechahora = " & DBSet(txtCodigo(1).Text, "F") & ", idcontador = " & DBSet(Max, "N")
                Sql = Sql & " where codtraba = " & DBSet(AntCodTraba, "N") & " and descontado = 0 "
                conn.Execute Sql
        
        
        
                Sql3 = "insert into rrecibosnomina (codtraba, fechahora, importe, base34, porcsegso1, porcsegso2, porcirpf, "
                Sql3 = Sql3 & "importesegso1, importesegso2, importeirpf, complemento, neto34, idcontador, hayembargo) values ("
                Sql3 = Sql3 & DBSet(AntCodTraba, "N") & ","
                Sql3 = Sql3 & DBSet(txtCodigo(20).Text, "F") & "," & DBSet(ImporteSinFormato(CStr(TImpbruto)), "N") & ","
                Sql3 = Sql3 & DBSet(ImporteSinFormato(CStr(TImpBruto2)), "N") & ","
                '[Monica]05/01/2012: SegSoc pasa a ser porcentaje
                'Sql3 = Sql3 & DBSet(0, "N") & ","
                Sql3 = Sql3 & DBSet(Rs2!dtosegso, "N") & ","
                Sql3 = Sql3 & DBSet(Rs2!dtoreten, "N") & ","
                Sql3 = Sql3 & DBSet(Rs2!dtosirpf, "N") & ","
                Sql3 = Sql3 & DBSet(TSegSo, "N") & "," & DBSet(TRetencion, "N") & "," & DBSet(TIRPF, "N") & ","
                Sql3 = Sql3 & DBSet(0, "N") & ","
                Sql3 = Sql3 & DBSet(TNeto34, "N") & ","
                Sql3 = Sql3 & DBSet(Max, "N") & "," & DBSet(Rs2!HayEmbargo, "N") & ")"
                
                conn.Execute Sql3
        
                Sql3 = "insert into tmpinformes (codusu, codigo1, fecha1, importe1, importe2) values (" & vUsu.Codigo & "," & DBSet(AntCodTraba, "N") & ","
                Sql3 = Sql3 & DBSet(txtCodigo(20).Text, "F") & "," & DBSet(TNeto34, "N") & "," & DBSet(Rs2!HayEmbargo, "N") & ")"
                
                conn.Execute Sql3
        
                '[Monica]26/09/2016: si no hay embargo le pagamos
                If DBLet(Rs2!HayEmbargo) = 0 Then
                    Sql3 = "insert into tmpImpor (codtraba, importe) values ("
                    Sql3 = Sql3 & DBSet(AntCodTraba, "N") & "," & DBSet(ImporteSinFormato(CStr(TNeto34)), "N") & ")"
                    
                    conn.Execute Sql3
                End If
            End If
            
            TIRPF = 0
            TImpbruto = 0
            TImpBruto2 = 0
            TRetencion = 0
            TNeto34 = 0
            TSegSo = 0
            
            AntCodTraba = ActCodTraba
            ActCodTraba = DBSet(Rs!CodTraba, "N")
        
            Set Rs2 = Nothing
            
            Sql2 = "select salarios.*, straba.dtoreten, straba.dtosegso, straba.dtosirpf, straba.pluscapataz, straba.hayembargo from salarios, straba where straba.codtraba = " & DBSet(ActCodTraba, "N")
            Sql2 = Sql2 & " and salarios.codcateg = straba.codcateg "
            
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        End If
        
        ImpHoras = Round2(DBLet(Rs.Fields(2).Value, "N") * DBLet(Rs2!impsalar, "N"), 2)
                                    ' importe + pluscapataz + complemento - penalizacion
                                    
        ' en coopic llevamos en el bruto el plus del capataz
        ' y no hay imphoras
        ImpBruto = Round2(DBLet(Rs.Fields(5).Value, "N") + DBLet(Rs.Fields(3).Value, "N") - DBLet(Rs.Fields(4).Value, "N"), 2)
        
        TImpbruto = TImpbruto + ImpBruto
        
        IRPF = Round2(ImpBruto * DBLet(Rs2!dtosirpf, "N") / 100, 2)
        TIRPF = TIRPF + IRPF

'[Monica]05/01/2012: SegSoc pasa a ser porcentaje
        SegSoc = Round2(ImpBruto * DBLet(Rs2!dtosegso, "N") / 100, 2)
        
'[Monica]05/01/2012: SegSoc pasa a ser porcentaje
'        ImpBruto2 = ImpBruto - DBLet(Rs2!dtosegso, "N")
        ImpBruto2 = ImpBruto - DBLet(SegSoc, "N")
        TImpBruto2 = TImpBruto2 + ImpBruto2
        
'[Monica]05/01/2012: SegSoc pasa a ser porcentaje
'        TSegSo = TSegSo + DBLet(Rs2!dtosegso, "N")
        TSegSo = TSegSo + SegSoc
        
        Retencion = Round2(ImpBruto2 * DBLet(Rs2!dtoreten, "N") / 100, 2)
        TRetencion = TRetencion + Retencion
        
        Neto34 = ImpBruto2 - IRPF - Retencion
        
        
        '[Monica]25/05/2018: anticipos pendientes de descuento, como en natural
        Anticipo = AnticiposPendientes(Rs!CodTraba)
        Neto34 = Neto34 - Anticipo
        
        
        TNeto34 = TNeto34 + Neto34
        
        Rs.MoveNext
    Wend
    
    If HayReg Then
        IncrementarProgres Pb2, 1
        Mens = "Calculando Importes" & vbCrLf & vbCrLf & "Trabajador: " & AntCodTraba & vbCrLf
        
        '[Monica]23/03/2016: si el importe es negativo no entra
        If TNeto34 >= 0 Then
        
            '[Monica]25/05/2018: anticipos pendientes de descuento
            Sql = "update horasanticipos set descontado = 1, fechahora = " & DBSet(txtCodigo(20).Text, "F") & ", idcontador = " & DBSet(Max, "N")
            Sql = Sql & " where codtraba = " & DBSet(AntCodTraba, "N") & " and descontado = 0 "
            conn.Execute Sql
                        
        
        
        
            Sql3 = "insert into rrecibosnomina (codtraba, fechahora, importe, base34, porcsegso1, porcsegso2, porcirpf, "
            Sql3 = Sql3 & "importesegso1, importesegso2, importeirpf, complemento, neto34, idcontador, hayembargo) values ("
            Sql3 = Sql3 & DBSet(AntCodTraba, "N") & ","
            Sql3 = Sql3 & DBSet(txtCodigo(20).Text, "F") & "," & DBSet(ImporteSinFormato(CStr(TImpbruto)), "N") & ","
            Sql3 = Sql3 & DBSet(ImporteSinFormato(CStr(TImpBruto2)), "N") & ","
            '[Monica]05/01/2012: SegSoc pasa a ser porcentaje
            'Sql3 = Sql3 & DBSet(0, "N") & ","
            Sql3 = Sql3 & DBSet(Rs2!dtosegso, "N") & ","
            Sql3 = Sql3 & DBSet(Rs2!dtoreten, "N") & ","
            Sql3 = Sql3 & DBSet(Rs2!dtosirpf, "N") & ","
            Sql3 = Sql3 & DBSet(TSegSo, "N") & "," & DBSet(TRetencion, "N") & "," & DBSet(TIRPF, "N") & ","
            Sql3 = Sql3 & DBSet(0, "N") & ","
            Sql3 = Sql3 & DBSet(TNeto34, "N") & ","
            Sql3 = Sql3 & DBSet(Max, "N") & "," & DBSet(Rs2!HayEmbargo, "N") & ")"
            
            conn.Execute Sql3
    
            Sql3 = "insert into tmpinformes (codusu, codigo1, fecha1, importe1, importe2) values (" & vUsu.Codigo & "," & DBSet(AntCodTraba, "N") & ","
            Sql3 = Sql3 & DBSet(txtCodigo(20).Text, "F") & "," & DBSet(TNeto34, "N") & "," & DBSet(Rs2!HayEmbargo, "N") & ")"
            
            conn.Execute Sql3
            
            
            '[Monica]26/09/2016: si no hay embargo le pagamos
            If DBLet(Rs2!HayEmbargo) = 0 Then
                
                Sql3 = "insert into tmpImpor (codtraba, importe) values ("
                Sql3 = Sql3 & DBSet(AntCodTraba, "N") & "," & DBSet(ImporteSinFormato(CStr(TNeto34)), "N") & ")"
                
                conn.Execute Sql3
            End If
        End If
        
        Set Rs2 = Nothing
    End If
    
    Dim vSeccion As CSeccion
    If vSeccion Is Nothing Then
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
            vSeccion.AbrirConta
        End If
    End If
    
    ' generamos el fichero plano del anticipo
    If vParamAplic.Cooperativa = 0 Then
'*************
        '[Monica]22/11/2013: iban
        Sql = "select codbanco, codsucur, digcontr, cuentaba, codorden34, iban from banpropi where codbanpr = " & DBSet(txtCodigo(0).Text, "N")
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        CodigoOrden34 = ""
        
        If Rs.EOF Then
            cad = ""
        Else
            If IsNull(Rs!CodBanco) Then
                cad = ""
            Else
                '[Monica]22/11/2013: iban
                cad = Format(Rs!CodBanco, "0000") & "|" & Format(DBLet(Rs!CodSucur, "T"), "0000") & "|" & DBLet(Rs!digcontr, "T") & "|" & Format(DBLet(Rs!CuentaBa, "T"), "0000000000") & "|" & DBLet(Rs!IBAN, "T") & "|"
            End If
            CodigoOrden34 = DBLet(Rs!codorden34, "T")
        End If
        
        Set Rs = Nothing
        
        CuentaPropia = cad
        
        '[Monica]02/02/2018: Catadau ha de generar el fichero
        If vEmpresa.AplicarNorma19_34Nueva = 1 Then
            If HayXML Then
                B = GeneraFicheroNorma34SEPA_XML(vParam.CifEmpresa, CDate(txtCodigo(20).Text), CuentaPropia, "", "Pago N�mina", Combo1(2).ListIndex, CodigoOrden34)
            Else
                B = GeneraFicheroNorma34SEPA(vParam.CifEmpresa, CDate(txtCodigo(20).Text), CuentaPropia, "", "Pago N�mina", Combo1(2).ListIndex, CodigoOrden34)
            End If
        Else
            B = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtCodigo(20).Text), CuentaPropia, 9, 0, "Pago N�mina", CodigoOrden34, Combo1(2).ListIndex)
        End If
        
    'antes
    '    b = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtcodigo(20).Text), CuentaPropia, 9, 0, "Pago N�mina", CodigoOrden34, Combo1(0).ListIndex)
    
'**************CASO DE CATADAU
    Else
        B = GeneraFicheroA3(Max, txtCodigo(20).Text)
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
'antes
'    b = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtcodigo(20).Text), CuentaPropia, 9, 0, "Pago N�mina", CodigoOrden34, Combo1(0).ListIndex)
    If B Then
        Mens = "Copiar fichero"
        
        If vParamAplic.Cooperativa = 0 Then
            CopiarFichero
        Else
            CopiarFicheroA3 "AnticipoA3.txt", txtCodigo(20).Text
        End If
        
        If B Then
            CadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
            CadParam = CadParam & "pFechaRecibo=""" & txtCodigo(20).Text & """|pFechaPago=""" & txtCodigo(20).Text & """|" & "pImpagados=0|"
            numParam = 4
            
            '[Monica]23/04/2018: mostramos el banco
            If vParamAplic.Cooperativa = 0 Then
                CadParam = CadParam & "pBanco=""" & txtCodigo(0).Text & " " & txtNombre(0).Text & """|"
                numParam = numParam + 1
            End If
                
            cadFormula = "{tmpinformes.codusu}=" & vUsu.Codigo & " and {tmpinformes.importe2} = 0"
            cadNombreRPT = "rListadoPagos.rpt"
            cadTitulo = "Impresion de Pagos"
            ConSubInforme = True

            LlamarImprimir
            
            '[Monica]17/10/2016: impresion de los impagados de Picassent
            Sql = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo & " and importe2 = 1"
            If CInt(DevuelveValor(Sql)) <> 0 Then
                CadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
                CadParam = CadParam & "pFechaRecibo=""" & txtCodigo(20).Text & """|pFechaPago=""" & txtCodigo(20).Text & """|" & "pImpagados=1|"
                numParam = 4
                '[Monica]23/04/2018: mostramos el banco
                If vParamAplic.Cooperativa = 0 Then
                    CadParam = CadParam & "pBanco=""" & txtCodigo(0).Text & " " & txtNombre(0).Text & """|"
                    numParam = numParam + 1
                End If
                cadFormula = "{tmpinformes.codusu}=" & vUsu.Codigo & " and {tmpinformes.importe2} = 1"
                cadNombreRPT = "rListadoPagos.rpt"
                cadTitulo = "Impresion de Impagos"
                ConSubInforme = True
    
                LlamarImprimir
            End If
            
            If Not Repetir Then
                If MsgBox("�Proceso realizado correctamente para actualizar?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    Sql = "update horas, straba, forpago set horas.intconta = 1, horas.fecharec = " & DBSet(txtCodigo(20).Text, "F") & " where horas.codtraba = straba.codtraba and straba.codforpa = forpago.codforpa and " & cadWHERE
                    conn.Execute Sql
                Else
                    Sql = "delete from rrecibosnomina where fechahora = " & DBSet(txtCodigo(20).Text, "F")
                    Sql = Sql & " and idcontador = " & DBSet(Max, "N")
                    
                    conn.Execute Sql
                End If
            End If
        Else
            B = False
        End If
    End If

eProcesarCambiosCoopic:
    If Err.Number <> 0 Then
        Mens = Err.Description
        B = False
    End If
    If B Then
        conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        cmdCancel_Click (1)
    Else
        conn.RollbackTrans
        MsgBox "Error " & Mens, vbExclamation
    End If
End Sub


Private Function AnticiposPendientes(CodTraba As String) As Currency
Dim Sql As String

    Sql = "select sum(importe) from horasanticipos where codtraba = " & DBSet(CodTraba, "N")
    Sql = Sql & " and descontado = 0 "
    
    AnticiposPendientes = DevuelveValor(Sql)
    
End Function
