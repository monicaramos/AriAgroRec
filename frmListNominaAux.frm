VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListNominaAux 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   12525
   Icon            =   "frmListNominaAux.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   12525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FramePaseABanco 
      Height          =   4995
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   6435
      Begin VB.TextBox txtCodigo 
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
         Index           =   66
         Left            =   1710
         MaxLength       =   30
         TabIndex        =   9
         Top             =   3045
         Width           =   4290
      End
      Begin VB.Frame FrameConcep 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   540
         TabIndex        =   13
         Top             =   4410
         Width           =   5715
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
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Tag             =   "Tipo|N|N|||straba|codsecci||N|"
         Top             =   2505
         Width           =   1665
      End
      Begin VB.CommandButton CmdAcepPaseBanco 
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
         Index           =   2
         Left            =   3915
         TabIndex        =   10
         Top             =   3960
         Width           =   1065
      End
      Begin VB.CommandButton CmdCancel 
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
         Left            =   5085
         TabIndex        =   11
         Top             =   3960
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
         Index           =   60
         Left            =   1755
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1335
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
         Index           =   58
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   7
         Tag             =   "Código|N|N|0|9999|rcapataz|codcapat|0000|S|"
         Top             =   1815
         Width           =   840
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
         Index           =   58
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   1815
         Width           =   3405
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   240
         Left            =   480
         TabIndex        =   1
         Top             =   3645
         Visible         =   0   'False
         Width           =   5670
         _ExtentX        =   10001
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   5190
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
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
         Index           =   87
         Left            =   405
         TabIndex        =   14
         Top             =   3015
         Width           =   1125
      End
      Begin VB.Label Label16 
         Caption         =   "Pase a Banco"
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
         Left            =   450
         TabIndex        =   5
         Top             =   405
         Width           =   5835
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   22
         Left            =   1425
         Picture         =   "frmListNominaAux.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   1335
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Pago"
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
         Index           =   78
         Left            =   450
         TabIndex        =   4
         Top             =   1035
         Width           =   1155
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
         Index           =   76
         Left            =   450
         TabIndex        =   3
         Top             =   1785
         Width           =   675
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   24
         Left            =   1425
         MouseIcon       =   "frmListNominaAux.frx":0097
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar banco"
         Top             =   1815
         Width           =   240
      End
      Begin VB.Label Label14 
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
         Left            =   450
         TabIndex        =   2
         Top             =   2235
         Width           =   1875
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6750
      Top             =   4980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmListNominaAux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: LAURA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public Opcionlistado As Byte
    ' 36 .- Pase a Banco de movimientos de Asesoria
    

    
    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(Cliente As String, Observaciones As String)

Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean



Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1

Private WithEvents frmBan As frmBasico2 'Banco propio
Attribute frmBan.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
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
Dim Tipo As String


Dim PrimeraVez As Boolean
Dim Contabilizada As Byte
Dim ConSubInforme As Boolean


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub



Private Sub CargarTemporalNotas(cTabla As String, cWhere As String)
Dim SQL As String
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
                                           'nroparte, numnota,  codsocio
    SQL = "insert into tmpinformes (codusu, importe1, importe2, importe3) "
    SQL = SQL & " select " & vUsu.Codigo & ", rpartes_variedad.nroparte, rhisfruta_entradas.numnotac, rhisfruta.codsocio from rpartes_variedad, rhisfruta, rhisfruta_entradas "
    SQL = SQL & " where rhisfruta.numalbar = rhisfruta_entradas.numalbar and rhisfruta_entradas.numnotac = rpartes_variedad.numnotac and rpartes_variedad.nroparte in "
    SQL = SQL & "(select rpartes.nroparte from " & cTabla
    If cWhere <> "" Then SQL = SQL & " where " & cWhere
    SQL = SQL & ") "
    SQL = SQL & " union "
    SQL = SQL & " select " & vUsu.Codigo & ", rpartes_variedad.nroparte, rclasifica.numnotac, rclasifica.codsocio from rpartes_variedad, rclasifica "
    SQL = SQL & " where  rclasifica.numnotac = rpartes_variedad.numnotac and rpartes_variedad.nroparte in "
    SQL = SQL & "(select rpartes.nroparte from " & cTabla
    If cWhere <> "" Then SQL = SQL & " where " & cWhere
    SQL = SQL & ") "
    
    conn.Execute SQL
    
    


End Sub

Private Function CargarTemporalPicassent(cadWhere As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim cad As String
Dim i As Integer
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

Dim Anticipado As Currency

On Error GoTo eProcesarCambiosPicassent
    
    CargarTemporalPicassent = False
    
    conn.Execute "delete from tmpinformes where codusu = " & vUsu.Codigo
    
    If cadWhere <> "" Then
        cadWhere = QuitarCaracterACadena(cadWhere, "{")
        cadWhere = QuitarCaracterACadena(cadWhere, "}")
        cadWhere = QuitarCaracterACadena(cadWhere, "_1")
    End If
        
    SQL = "select count(distinct horas.codtraba) from (horas inner join straba on horas.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Rs.Close
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    SQL = "select horas.codtraba,  sum(horasdia), sum(compleme), sum(penaliza), sum(importe) from (horas inner join straba on horas.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWhere
    SQL = SQL & " group by horas.codtraba "
    SQL = SQL & " order by 1 "
        
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Mens = "Calculando Importes" & vbCrLf & vbCrLf & "Trabajador: " & Rs!CodTraba & vbCrLf
        
        Sql2 = "select salarios.*, straba.dtoreten, straba.dtosegso, straba.dtosirpf, straba.pluscapataz from salarios, straba where straba.codtraba = " & DBSet(Rs!CodTraba, "N")
        Sql2 = Sql2 & " and salarios.codcateg = straba.codcateg "
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        ImpHoras = Round2(DBLet(Rs.Fields(1).Value, "N") * DBLet(Rs2!impsalar, "N"), 2)
                                    ' importe + pluscapataz + complemento - penalizacion
        ImpBruto = Round2(ImpHoras + DBLet(Rs.Fields(4).Value, "N") + DBLet(Rs2!PlusCapataz, "N") + DBLet(Rs.Fields(2).Value, "N") - DBLet(Rs.Fields(3).Value, "N"), 2)
                                                'codtraba,bruto,    anticipado,diferencia
        
        '[Monica]05/10/2010: el importe bruto es el que le he pagaria sin cargar ningun dto
        Sql5 = "select sum(importe) from rrecibosnomina where codtraba = " & DBSet(Rs!CodTraba, "N")
        Sql5 = Sql5 & " and fechahora >= " & DBSet(txtCodigo(44).Text, "F")
        Sql5 = Sql5 & " and fechahora <= " & DBSet(txtCodigo(48).Text, "F")
        ImpBruto = DevuelveValor(Sql5)
        
        '[Monica]05/10/2010: el importe anticipado es el importe liquido (antes sum(importe) era incorrecto)
        Sql5 = "select sum(neto34) from rrecibosnomina where codtraba = " & DBSet(Rs!CodTraba, "N")
        Sql5 = Sql5 & " and fechahora >= " & DBSet(txtCodigo(44).Text, "F")
        Sql5 = Sql5 & " and fechahora <= " & DBSet(txtCodigo(48).Text, "F")
                                                
        Anticipado = DevuelveValor(Sql5)
        Diferencia = ImpBruto - Anticipado
                                                
        Sql3 = "insert into tmpinformes (codusu, codigo1, importe1, importe2, importe3) values ("
        Sql3 = Sql3 & vUsu.Codigo & ","
        Sql3 = Sql3 & DBSet(Rs.Fields(0).Value, "N") & ","
        Sql3 = Sql3 & DBSet(ImpBruto, "N") & ","
        Sql3 = Sql3 & DBSet(Anticipado, "N") & ","
        Sql3 = Sql3 & DBSet(Diferencia, "N") & ")"
        
        conn.Execute Sql3

        Set Rs2 = Nothing
            
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    CargarTemporalPicassent = True
    Exit Function
    
eProcesarCambiosPicassent:
    If Err.Number <> 0 Then
        Mens = Err.Description
        MsgBox "Error " & Mens, vbExclamation
    End If
End Function



Private Sub CmdAcepPaseBanco_Click(Index As Integer)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim NomAlmac As String
Dim cTabla As String
Dim SQL As String

    If Not DatosOk Then Exit Sub
    
    
    InicializarVbles
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    

   'La forma de pago tiene que ser de tipo Transferencia
   AnyadirAFormula cadselect, "forpago.tipoforp = 1"
   
  

   tabla = "(tmpinformes INNER JOIN straba ON tmpinformes.codigo1 = straba.codtraba) INNER JOIN forpago ON straba.codforpa = forpago.codforpa "
              
   cTabla = tabla
   
   cTabla = QuitarCaracterACadena(cTabla, "{")
   cTabla = QuitarCaracterACadena(cTabla, "}")
   SQL = "Select count(*) FROM " & QuitarCaracterACadena(cTabla, "_1")
   If cadselect <> "" Then
       cadselect = QuitarCaracterACadena(cadselect, "{")
       cadselect = QuitarCaracterACadena(cadselect, "}")
       cadselect = QuitarCaracterACadena(cadselect, "_1")
       SQL = SQL & " WHERE " & cadselect
   End If
   
   If RegistrosAListar(SQL) = 0 Then
       MsgBox "No hay datos para mostrar en el Informe.", vbInformation
   Else
       ProcesoPaseABancoAnticipos (cadselect)
   End If
    
End Sub

Private Function TrabajadoresEnActivo(Fecha As String) As Boolean
Dim SQL As String

    On Error GoTo eTrabajadoresEnActivo

    TrabajadoresEnActivo = False

    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL


    SQL = "insert into tmpinformes (codusu, codigo1, nombre1, nombre2) "
    SQL = SQL & "select " & vUsu.Codigo & ", codtraba, nomtraba, niftraba "
    SQL = SQL & " from straba where fechaalta <= " & DBSet(Fecha, "F")
    SQL = SQL & " and fechabaja is null "
    conn.Execute SQL
    
    SQL = "delete from tmpinformes2 where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    ' en tmpinformes2 metemos en que cuadrilla estan
    SQL = "insert into tmpinformes2 (codusu, codigo1, importe1) "
    SQL = SQL & " select codusu, codigo1, rcuadrilla_trabajador.codcuadrilla "
    SQL = SQL & " from  tmpinformes left join rcuadrilla_trabajador on tmpinformes.codigo1 = rcuadrilla_trabajador.codtraba "
    SQL = SQL & " where tmpinformes.codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    SQL = "update tmpinformes2,  rcuadrilla, rcapataz      "
    SQL = SQL & " set tmpinformes2.nombre1 = rcapataz.nomcapat "
    SQL = SQL & " where tmpinformes2.codusu = " & vUsu.Codigo
    SQL = SQL & " and tmpinformes2.importe1 = rcuadrilla.codcuadrilla "
    SQL = SQL & " and rcuadrilla.codcapat = rcapataz.codcapat "
    conn.Execute SQL
    
    TrabajadoresEnActivo = True
    
    Exit Function
    
    
eTrabajadoresEnActivo:
    MuestraError Err.Number, "Carga Trabajadores en Activo", Err.Description
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
        Select Case Opcionlistado
               
            Case 36 ' Pase a banco de importes
                Combo1(0).ListIndex = 0
                txtCodigo(60).Text = Format(Now, "dd/mm/yyyy")
                PonerFoco txtCodigo(62)
                
                '[Monica]18/09/2013: anticipos para Natural
                FrameConcep.visible = (vParamAplic.Cooperativa = 9)
                FrameConcep.Enabled = (vParamAplic.Cooperativa = 9)
                If vParamAplic.Cooperativa = 9 Then
                    Label2(77).Caption = "Fecha"
                    txtCodigo(66).Text = "ANTICIPO " & UCase(MonthName(Month(Now))) & " " & Year(Now)
                End If
        End Select
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
    Set List = New Collection
    For H = 24 To 27
        List.Add H
    Next H
    For H = 1 To 10
        List.Add H
    Next H
    List.Add 12
    List.Add 13
    List.Add 14
    List.Add 15
    List.Add 18
    List.Add 19
    
'    For H = 0 To 34
'        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
'    Next H
    Me.imgBuscar(24).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    
   
    Set List = Nothing

    'Ocultar todos los Frames de Formulario
    Me.FramePaseABanco.visible = False
    
    '###Descomentar
'    CommitConexion
    
    Select Case Opcionlistado
    
    'LISTADOS DE MANTENIMIENTOS BASICOS
    '---------------------
    
    Case 36 ' pase a banco
        CargaCombo
    
        FramePaseaBancoVisible True, H, W
        indFrame = 0
        tabla = "tmpinformes"
    
    End Select
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.CmdCancel(0).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub


Private Sub FechaBajaVisible(Mostrar As Boolean)
    Label2(105).visible = Mostrar
    imgFecha(25).visible = Mostrar
    txtCodigo(78).visible = Mostrar
End Sub



Private Sub frmBan_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtCodigo(CByte(imgFecha(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub


Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        
        Case 24 ' banco
            AbrirFrmManBanco (Index + 34)
        
        
    End Select
    PonerFoco txtCodigo(indCodigo)
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
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmC.Left = esq + imgFecha(Index).Parent.Left + 30
    frmC.Top = dalt + imgFecha(Index).Parent.Top + imgFecha(Index).Height + menu - 40
    
    Select Case Index
        Case 0, 1, 13, 2, 3
            Indice = Index + 14
        Case 4, 5
            Indice = Index
        Case 7
            Indice = 11
        Case 6
            Indice = 29
        Case 8
            Indice = 30
        Case 9
            Indice = 35
        Case 10
            Indice = 26
        Case 12
            Indice = 37
        Case 11
            Indice = 33
        Case 14
            Indice = 46
        Case 15
            Indice = 44
        Case 16
            Indice = 48
        Case 17, 18
            Indice = Index + 35
        Case 19, 20
            Indice = Index + 37
        Case 21, 22
            Indice = Index + 38
        Case 25
            Indice = 78
        Case 26
            Indice = 82
        Case 28
            Indice = 81
    End Select
    
    imgFecha(0).Tag = Indice '<===
    
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtCodigo(Indice).Text <> "" Then frmC.NovaData = txtCodigo(Indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtCodigo(CByte(imgFecha(0).Tag)) '<===
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
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 2: KEYBusqueda KeyAscii, 0 'trabajador desde
            Case 3: KEYBusqueda KeyAscii, 1 'trabajador hasta
            Case 6:  KEYBusqueda KeyAscii, 2 'variedad desde
            Case 7:  KEYBusqueda KeyAscii, 3 'variedad hasta
            
            Case 18: KEYBusqueda KeyAscii, 14 'trabajador desde
            Case 19: KEYBusqueda KeyAscii, 15 'trabajador hasta
            
            Case 14: KEYFecha KeyAscii, 0 'fecha desde
            Case 15: KEYFecha KeyAscii, 1 'fecha hasta
            
            Case 24: KEYBusqueda KeyAscii, 20 'almacen para el calculo de horas productivas
        
            Case 9:  KEYBusqueda KeyAscii, 5 ' variedad
            Case 11: KEYFecha KeyAscii, 7 ' fecha
            Case 12: KEYBusqueda KeyAscii, 6 'capataz
        
            Case 35: KEYFecha KeyAscii, 9 ' fecha desde
            Case 26: KEYFecha KeyAscii, 10 ' fecha hasta
            
            Case 34:  KEYBusqueda KeyAscii, 4 'capataz
            Case 36: KEYBusqueda KeyAscii, 7 ' variedad
            
        
            Case 31: KEYBusqueda KeyAscii, 8 'capataz desde
            Case 32: KEYBusqueda KeyAscii, 9 'capataz hasta
            Case 29: KEYFecha KeyAscii, 6 'fecha desde
            Case 30: KEYFecha KeyAscii, 8 'fecha hasta
        
            Case 28:  KEYBusqueda KeyAscii, 11 ' variedad
            Case 37: KEYFecha KeyAscii, 12 ' fecha desde
            Case 33: KEYFecha KeyAscii, 11 ' fecha hasta
            Case 41: KEYBusqueda KeyAscii, 12 'trabajador desde
            Case 42: KEYBusqueda KeyAscii, 13 'trabajador hasta
        
            Case 47:  KEYBusqueda KeyAscii, 16 ' variedad
            Case 46: KEYFecha KeyAscii, 14 ' fecha desde
            Case 45: KEYBusqueda KeyAscii, 10 'capataz
        
            Case 44: KEYFecha KeyAscii, 15 ' fecha desde
            Case 48: KEYFecha KeyAscii, 16 ' fecha hasta
            Case 49: KEYBusqueda KeyAscii, 19 'trabajador desde
            Case 50: KEYBusqueda KeyAscii, 21 'trabajador hasta
        
            Case 38: KEYBusqueda KeyAscii, 17 'capataz desde
            Case 43: KEYBusqueda KeyAscii, 18 'capataz hasta
            Case 52: KEYFecha KeyAscii, 17 ' fecha desde
            Case 53: KEYFecha KeyAscii, 18 ' fecha hasta
        
            Case 54: KEYBusqueda KeyAscii, 22 'trabajador desde
            Case 55: KEYBusqueda KeyAscii, 23 'trabajador hasta
            Case 56: KEYFecha KeyAscii, 19 ' fecha desde
            Case 57: KEYFecha KeyAscii, 20 ' fecha hasta
            
            ' Pase a bancos
            Case 62: KEYBusqueda KeyAscii, 25 'trabajador desde
            Case 63: KEYBusqueda KeyAscii, 26 'trabajador hasta
            Case 59: KEYFecha KeyAscii, 21 ' fecha
            Case 60: KEYFecha KeyAscii, 22 ' fecha
            Case 58: KEYBusqueda KeyAscii, 24 'banco
        
            Case 64: KEYBusqueda KeyAscii, 27 'trabajador desde
            Case 65: KEYBusqueda KeyAscii, 28 'trabajador hasta
        
            Case 68: KEYBusqueda KeyAscii, 29 'trabajador desde
            Case 69: KEYBusqueda KeyAscii, 30 'trabajador hasta
        
        
            Case 72: KEYBusqueda KeyAscii, 31 'capataz desde
            Case 73: KEYBusqueda KeyAscii, 32 'capataz hasta
            Case 70: KEYFecha KeyAscii, 23 'fecha desde
            Case 71: KEYFecha KeyAscii, 24 'fecha hasta
        
            Case 76: KEYBusqueda KeyAscii, 33 'capataz desde
            Case 77: KEYBusqueda KeyAscii, 34 'capataz hasta
        
            Case 78: KEYFecha KeyAscii, 25 'fecha de baja del trabajador (coopic)
        
            Case 82: KEYFecha KeyAscii, 26 'fecha de creacion
            Case 80: KEYBusqueda KeyAscii, 36 'capataz hasta
        
            Case 81: KEYFecha KeyAscii, 28 'fecha de trabajadores activos
        
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
        Case 0, 1 ' Nro.Partes
            PonerFormatoEntero txtCodigo(Index)
    
        Case 4, 5, 14, 15, 16, 17, 27, 11, 29, 30 'FECHAS
            If txtCodigo(Index).Text <> "" Then
                PonerFormatoFecha txtCodigo(Index)
            End If
            
        Case 35, 26, 33, 37, 46, 44, 48, 52, 53, 56, 57, 59, 60, 70, 71, 78, 82, 81 'FECHAS
            If txtCodigo(Index).Text <> "" Then
                PonerFormatoFecha txtCodigo(Index)
            End If
         
            
            
        Case 58 'BANCO
            PonerFormatoEntero txtCodigo(Index)
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "banpropi", "nombanpr", "codbanpr", "N")
        
        Case 74, 75 ' Nro de parte
            PonerFormatoEntero txtCodigo(Index)
        
    End Select
End Sub





Private Sub FramePaseaBancoVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el pase a banco
    Me.FramePaseABanco.visible = visible
    If visible = True Then
        Me.FramePaseABanco.Top = -90
        Me.FramePaseABanco.Left = 0
        Me.FramePaseABanco.Height = 5990 '5130
        Me.FramePaseABanco.Width = 6435
        W = Me.FramePaseABanco.Width
        H = Me.FramePaseABanco.Height
    End If
End Sub


Private Sub InicializarVbles()
    cadFormula = ""
    cadselect = ""
    cadParam = ""
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
        If Not AnyadirAFormula(cadselect, devuelve) Then Exit Function
    Else
        devuelve2 = CadenaDesdeHastaBD(codD, codH, Codigo, TipCod)
        If devuelve2 = "Error" Then Exit Function
        If Not AnyadirAFormula(cadselect, devuelve2) Then Exit Function
    End If
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            cadParam = cadParam & AnyadirParametroDH(param, codD, codH, nomD, nomH)
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function

Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .EnvioEMail = False
        .ConSubInforme = ConSubInforme
        .Opcion = Opcionlistado
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmManBanco(Indice As Integer)
    indCodigo = Indice
    
    Set frmBan = New frmBasico2
    
    AyudaBancosCom frmBan, txtCodigo(indCodigo)
    
    Set frmBan = Nothing
    
    PonerFoco txtCodigo(indCodigo)
End Sub









Private Sub ProcesoPaseABanco(cadWhere As String)
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim cad As String
Dim i As Integer
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
Dim Neto As Currency
Dim Bruto As Currency
Dim CuentaPropia As String
Dim CodigoOrden34 As String
Dim Extra As String

Dim AntOpcion As Integer

On Error GoTo eProcesoPaseABanco
    
    BorrarTMPs
    CrearTMPs

    conn.BeginTrans
    
    SQL = "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute SQL
    
    
    If cadWhere <> "" Then
        cadWhere = QuitarCaracterACadena(cadWhere, "{")
        cadWhere = QuitarCaracterACadena(cadWhere, "}")
        cadWhere = QuitarCaracterACadena(cadWhere, "_1")
    End If
        
    SQL = "select count(distinct rrecasesoria.codtraba) from (rrecasesoria inner join straba on rrecasesoria.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWhere
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Pb1.visible = True
    CargarProgres Pb1, Rs.Fields(0).Value
    
    Rs.Close
    
    SQL = "select rrecasesoria.codtraba, sum(importe) importe from (rrecasesoria inner join straba on rrecasesoria.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWhere
    SQL = SQL & " group by rrecasesoria.codtraba "
    
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        IncrementarProgres Pb1, 1
        
        '[Monica]23/03/2016: si el importe es negativo no entra
        If DBLet(Rs!Importe) >= 0 Then
            
            Sql3 = "insert into tmpImpor (codtraba, importe) values ("
            Sql3 = Sql3 & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(ImporteSinFormato(CStr(Rs!Importe)), "N") & ")"
            
            conn.Execute Sql3
            
            
            Sql3 = "insert into tmpinformes (codusu, codigo1, fecha1, importe1) values (" & vUsu.Codigo & "," & DBSet(Rs.Fields(0).Value, "N") & ","
            Sql3 = Sql3 & DBSet(txtCodigo(59).Text, "F") & "," & DBSet(ImporteSinFormato(CStr(Rs!Importe)), "N") & ")"
                
            conn.Execute Sql3
            
        End If
        
        Set Rs2 = Nothing
            
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    '[Monica]22/11/2013: iban
    SQL = "select codbanco, codsucur, digcontr, cuentaba, codorden34, sufijoem, iban from banpropi where codbanpr = " & DBSet(txtCodigo(58).Text, "N")
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CodigoOrden34 = ""
    Extra = ""
    If Rs.EOF Then
        cad = ""
    Else
        If IsNull(Rs!CodBanco) Then
            cad = ""
        Else
            '[Monica]22/11/2013: iban
            cad = Format(Rs!CodBanco, "0000") & "|" & Format(DBLet(Rs!CodSucur, "T"), "0000") & "|" & DBLet(Rs!digcontr, "T") & "|" & Format(DBLet(Rs!CuentaBa, "T"), "0000000000") & "|" & DBLet(Rs!Iban, "T") & "|"
        End If
        CodigoOrden34 = DBLet(Rs!codorden34, "T")
        Extra = DBLet(Rs!sufijoem, "T") & "|" & vParam.NombreEmpresa & "|"
    End If
    
    Set Rs = Nothing
    
    CuentaPropia = cad
    
    '[Monica]22/11/2013: iban
    Dim vSeccion As CSeccion
    If vSeccion Is Nothing Then
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
            vSeccion.AbrirConta
        End If
    End If
    
    If vEmpresa.AplicarNorma19_34Nueva = 1 Then
        If HayXML Then
            B = GeneraFicheroNorma34SEPA_XML(vParam.CifEmpresa, CDate(txtCodigo(60).Text), CuentaPropia, "", "Pago Nómina", Combo1(0).ListIndex, CodigoOrden34)
        Else
            B = GeneraFicheroNorma34SEPA(vParam.CifEmpresa, CDate(txtCodigo(60).Text), CuentaPropia, "", "Pago Nómina", Combo1(0).ListIndex, CodigoOrden34)
        End If
    Else
        B = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtCodigo(60).Text), CuentaPropia, 9, 0, "Pago Nómina", CodigoOrden34, Combo1(0).ListIndex)
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
     
'antes
'    b = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtCodigo(60).Text), CuentaPropia, 9, 0, "Pago Nómina", CodigoOrden34, Combo1(0).ListIndex)
    If B Then
        B = CopiarFichero
        If B Then
            cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
            cadParam = cadParam & "pFechaRecibo=""" & txtCodigo(59).Text & """|pFechaPago=""" & txtCodigo(60).Text & """|"
            numParam = 3
            cadFormula = "{tmpinformes.codusu}=" & vUsu.Codigo
            cadNombreRPT = "rListadoPagos.rpt"
            cadTitulo = "Impresion de Pagos"
            ConSubInforme = False
            
            AntOpcion = Opcionlistado
            Opcionlistado = 0

            LlamarImprimir
            
            Opcionlistado = AntOpcion
            
            If MsgBox("¿Proceso realizado correctamente para actualizar?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                SQL = "update rrecasesoria, straba, forpago set rrecasesoria.idconta = 1 where rrecasesoria.codtraba = straba.codtraba and straba.codforpa = forpago.codforpa and " & cadWhere
                conn.Execute SQL
            End If
        End If
    End If

eProcesoPaseABanco:
    If Err.Number <> 0 Then
        Mens = Err.Description
        B = False
    End If
    If B Then
        conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        cmdCancel_Click (0)
    Else
        conn.RollbackTrans
        MsgBox "Error " & Mens, vbExclamation
    End If
End Sub


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


Private Function DatosOk() As Boolean
Dim B As Boolean
Dim SQL As String
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOk = False
    B = True
    
    If txtCodigo(60).Text = "" Then
        SQL = "Debe introducir obligatoriamente un valor en los campos de fecha. Reintroduzca. " & vbCrLf & vbCrLf
        MsgBox SQL, vbExclamation
        B = False
        PonerFoco txtCodigo(59)
    End If
    If B Then
        If txtCodigo(58).Text = "" Then
            SQL = "Debe introducir obligatoriamente un valor en el banco. Reintroduzca. " & vbCrLf & vbCrLf
            MsgBox SQL, vbExclamation
            B = False
            PonerFoco txtCodigo(58)
        End If
    End If
    '[Monica]18/09/2013: debe introducir el concepto
    If B And vParamAplic.Cooperativa = 9 Then
        If txtCodigo(66).Text = "" Then
            SQL = "Debe introducir obligatoriamente una descripción. Reintroduzca. " & vbCrLf & vbCrLf
            MsgBox SQL, vbExclamation
            B = False
            PonerFoco txtCodigo(66)
        End If
    End If
        
    DatosOk = B
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
'    For I = 0 To Combo1.Count - 1
'        Combo1(I).Clear
'    Next I
    Combo1(0).Clear
    
    Combo1(0).AddItem "Nómina"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Pensión"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Otros Conceptos"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    
End Sub


Private Sub ProcesoPaseABancoAnticipos(cadWhere As String)
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim cad As String
Dim i As Integer
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
Dim Neto As Currency
Dim Bruto As Currency
Dim CuentaPropia As String
Dim CodigoOrden34 As String
Dim Extra As String

Dim AntOpcion As Integer

On Error GoTo eProcesoPaseABanco
    
    BorrarTMPs
    CrearTMPs

    conn.BeginTrans
    
    
    SQL = "select tmpinformes.codigo1, sum(importe1) importe, sum(coalesce(importe4,0)) anticipo from (tmpinformes inner join straba on tmpinformes.codigo1 = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where tmpinformes.codusu =" & DBSet(vUsu.Codigo, "N")
    SQL = SQL & " group by tmpinformes.codigo1 "
    Set Rs = New ADODB.Recordset
    
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        IncrementarProgres Pb1, 1
        
        '[Monica]23/03/2016: si el importe es negativo no entra
        If DBLet(Rs!Importe) - DBLet(Rs!Anticipo) >= 0 Then
        
            Sql3 = "insert into tmpImpor (codtraba, importe) values ("
            Sql3 = Sql3 & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(ImporteSinFormato(CStr(DBLet(Rs!Importe, "N") - DBLet(Rs!Anticipo, "N"))), "N") & ")"
            
            conn.Execute Sql3
            
            
        End If
        
        Set Rs2 = Nothing
            
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    '[Monica]22/11/2013: iban
    SQL = "select codbanco, codsucur, digcontr, cuentaba, codorden34, sufijoem, iban from banpropi where codbanpr = " & DBSet(txtCodigo(58).Text, "N")
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CodigoOrden34 = ""
    Extra = ""
    If Rs.EOF Then
        cad = ""
    Else
        If IsNull(Rs!CodBanco) Then
            cad = ""
        Else
            '[Monica]22/11/2013: iban
            cad = Format(Rs!CodBanco, "0000") & "|" & Format(DBLet(Rs!CodSucur, "T"), "0000") & "|" & DBLet(Rs!digcontr, "T") & "|" & Format(DBLet(Rs!CuentaBa, "T"), "0000000000") & "|" & DBLet(Rs!Iban, "T") & "|"
        End If
        CodigoOrden34 = DBLet(Rs!codorden34, "T")
        Extra = DBLet(Rs!sufijoem, "T") & "|" & vParam.NombreEmpresa & "|"
    End If
    
    Set Rs = Nothing
    
    CuentaPropia = cad
    '[Monica]22/11/2013: iban
    Dim vSeccion As CSeccion
    If vSeccion Is Nothing Then
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
            vSeccion.AbrirConta
        End If
    End If
    
    '[Monica]30/04/2019: caso unicamente de frutas inma
    Dim AntCifEmpresa As String
    Dim AntNomEmpresa As String
    Dim AntProvincia As String
    Dim AntPoblacion As String
    Dim AntCPostal As String
    Dim AntDomicilio As String
    
    AntCifEmpresa = vParam.CifEmpresa
    AntNomEmpresa = vEmpresa.nomempre
    AntCPostal = vParam.CPostal
    AntProvincia = vParam.Provincia
    AntDomicilio = vParam.DomicilioEmpresa
    AntPoblacion = vParam.Poblacion
    
    vParam.CifEmpresa = DevuelveValor("select cifcoope from rcoope where codcoope = 2")
    vEmpresa.nomempre = DevuelveValor("select nomcoope from rcoope where codcoope = 2")
    vParam.CPostal = DevuelveValor("select codposta from rcoope where codcoope = 2")
    vParam.Provincia = DevuelveValor("select procoope from rcoope where codcoope = 2")
    vParam.DomicilioEmpresa = DevuelveValor("select domcoope from rcoope where codcoope = 2")
    vParam.Poblacion = DevuelveValor("select pobcoope from rcoope where codcoope = 2")
    
    If vEmpresa.AplicarNorma19_34Nueva = 1 Then
        If HayXML Then
            B = GeneraFicheroNorma34SEPA_XML(vParam.CifEmpresa, CDate(txtCodigo(60).Text), CuentaPropia, "", txtCodigo(66).Text, Combo1(0).ListIndex, CodigoOrden34)
        Else
            B = GeneraFicheroNorma34SEPA(vParam.CifEmpresa, CDate(txtCodigo(60).Text), CuentaPropia, "", txtCodigo(66).Text, Combo1(0).ListIndex, CodigoOrden34)
        End If
    Else
        B = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtCodigo(60).Text), CuentaPropia, 9, 0, txtCodigo(66).Text, CodigoOrden34, Combo1(0).ListIndex)
    End If
    
    vParam.CifEmpresa = AntCifEmpresa
    vEmpresa.nomempre = AntNomEmpresa
    vParam.CPostal = AntCPostal
    vParam.Provincia = AntProvincia
    vParam.DomicilioEmpresa = AntDomicilio
    vParam.Poblacion = AntPoblacion
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
     
'antes
'    b = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtCodigo(60).Text), CuentaPropia, 9, 0, txtCodigo(66).Text, CodigoOrden34, Combo1(0).ListIndex)
    If B Then
        B = CopiarFichero
    End If

eProcesoPaseABanco:
    If Err.Number <> 0 Then
        Mens = Err.Description
        B = False
    End If
    If B Then
        conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        cmdCancel_Click 0
    Else
        conn.RollbackTrans
        MsgBox "Error " & Mens, vbExclamation
    End If
End Sub

Private Sub BorrarTMPs()
On Error Resume Next

    conn.Execute " DROP TABLE IF EXISTS tmpImpor;"
    conn.Execute " DROP TABLE IF EXISTS tmpImporNeg;"
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function CrearTMPs() As Boolean
'Crea una temporal donde inserta la clave primaria de las
'facturas seleccionadas para facturar y trabaja siempre con ellas
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPs = False
    
    SQL = "CREATE TEMPORARY TABLE tmpImpor ( "
    SQL = SQL & "codtraba int(6) unsigned NOT NULL default '0',"
    SQL = SQL & "importe decimal(12,2)  NOT NULL default '0')"
    
    conn.Execute SQL
    
    SQL = "CREATE TEMPORARY TABLE tmpImporNeg ( "
    SQL = SQL & "codtraba int(6) unsigned NOT NULL default '0',"
    SQL = SQL & "concepto varchar(30),"
    SQL = SQL & "importe decimal(12,2)  NOT NULL default '0')"
    
    conn.Execute SQL
     
    CrearTMPs = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPs = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpImpor;"
        conn.Execute SQL
        SQL = " DROP TABLE IF EXISTS tmpImporNeg;"
        conn.Execute SQL
    End If
End Function

