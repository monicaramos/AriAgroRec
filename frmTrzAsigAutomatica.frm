VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTrzAsigAutomatica 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6660
   Icon            =   "frmTrzAsigAutomatica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCreacionPalets 
      Height          =   3705
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   6600
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
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1755
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
         Index           =   16
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1290
         Width           =   1350
      End
      Begin VB.CommandButton CmdAcepCreacionPalet 
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
         Left            =   3780
         TabIndex        =   4
         Top             =   2850
         Width           =   1065
      End
      Begin VB.CommandButton CmdCancelT 
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
         Left            =   4980
         TabIndex        =   6
         Top             =   2850
         Width           =   1065
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
         Left            =   540
         TabIndex        =   10
         Top             =   1350
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
         Index           =   28
         Left            =   540
         TabIndex        =   9
         Top             =   1740
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   360
         TabIndex        =   8
         Top             =   2565
         Width           =   5730
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   7
         Top             =   2340
         Width           =   5820
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1305
         Picture         =   "frmTrzAsigAutomatica.frx":000C
         Top             =   1755
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1305
         Picture         =   "frmTrzAsigAutomatica.frx":0097
         Top             =   1290
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
         Index           =   43
         Left            =   360
         TabIndex        =   5
         Top             =   945
         Width           =   600
      End
      Begin VB.Label Label9 
         Caption         =   "Creación automática de Palets"
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
         Index           =   3
         Left            =   360
         TabIndex        =   3
         Top             =   450
         Width           =   4725
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6075
      Top             =   4770
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTrzAsigAutomatica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: LAURA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(Cliente As String, Observaciones As String)

Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean


Private WithEvents frmArtADV As frmADVArticulos 'articulos de adv
Attribute frmArtADV.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'Mensajes
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmVol As frmTrzManCargas ' volcados por la linea 2
Attribute frmVol.VB_VarHelpID = -1


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


Dim vSeccion As CSeccion



Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me

End Sub


Private Sub CmdAcepCreacionPalet_Click()
Dim SQL As String
Dim fec As Date
Dim B As Boolean


    If txtCodigo(16).Text = "" Then
        MsgBox "Ha de introducir una fecha desde. Reintroduzca.", vbExclamation
        PonerFoco txtCodigo(16)
        Exit Sub
    End If
    If txtCodigo(17).Text = "" Then
        MsgBox "Ha de introducir una fecha hasta. Reintroduzca.", vbExclamation
        PonerFoco txtCodigo(17)
        Exit Sub
    End If
    
    If CDate(txtCodigo(16).Text) > CDate(txtCodigo(17).Text) Then
        MsgBox "Desde fecha no puede ser superior a hasta fecha. Reintroduzca.", vbExclamation
        PonerFoco txtCodigo(16)
        Exit Sub
    End If
    
    fec = CDate(txtCodigo(16))
    
    B = True
    
    While fec <= CDate(txtCodigo(17).Text) And B
        Label2(6).Caption = "Procesando fecha " & fec
        Label2(8).Caption = ""
        DoEvents
        
        SQL = "select * from trzlineas_cargas where fecha = " & DBSet(fec, "F")
        SQL = SQL & " and not idpalet in (select idpalet from palets where not idpalet is null) "
        
        If TotalRegistros(SQL) = 0 Then
            MsgBox "No se ha realizado ningún volcado esa fecha.", vbExclamation
        Else
            If FechaVolcadoCargada(fec) Then
                Exit Sub
            End If
        
            '[Monica]07/06/2019: mostramos todos los palets de la fecha que han sido pasados por la linea 2 y que van directamente a mermas
            '                    para que me cambien la linea si es para la basura
            Set frmVol = New frmTrzManCargas
            frmVol.NuevoCodigo = "not trzlineas_cargas.idpalet in (select idpalet from palets where not idpalet is null) and trzlineas_cargas.linea <> 1"
            frmVol.FechaCarga = fec
            frmVol.Show vbModal
            Set frmVol = Nothing
        
            If Not ComprobarExistenciasConAlbaranes(SQL, fec) Then
                Exit Sub
            End If
            
            If ProcesoCarga(SQL, fec) Then
                MsgBox "Proceso dia " & fec & " realizado correctamente.", vbExclamation
            Else
                B = False
            End If
        End If
        
        fec = DateAdd("d", 1, fec)
        
    Wend
End Sub

Private Function FechaVolcadoCargada(fec As Date) As Boolean
Dim SQL As String

    FechaVolcadoCargada = False
    
    SQL = "select count(*) from palets where fechaini = " & DBSet(fec, "F")
    If DevuelveValor(SQL) <> 0 Then
        MsgBox "Hay palets confeccionados con esa fecha. Revise.", vbExclamation
        FechaVolcadoCargada = True
    End If
    
    SQL = "select count(*) from trzmovim where fecha = " & DBSet(fec, "F")
    If DevuelveValor(SQL) <> 0 Then
        MsgBox "Hay movimientos de palets con esa fecha. Revise.", vbExclamation
        FechaVolcadoCargada = True
    End If
    

End Function

Private Function ComprobarExistenciasConAlbaranes(vSQL As String, fec As Date) As Boolean
Dim SQL As String
Dim KilosVolcados As Long
Dim CadVariedades As String
Dim Rs As ADODB.Recordset

    On Error GoTo eComprobarExistenciasConAlbaranes


    ComprobarExistenciasConAlbaranes = False


    Label2(8).Caption = "Comprobar existencias con albaranes"
    DoEvents

    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    
    ' kilos salidos en albaranes
    SQL = "insert into tmpinformes (codusu, codigo1, importe1) "
    SQL = SQL & " select " & vUsu.Codigo & ", codvarie, sum(coalesce(pesoneto)) pesoneto "
    SQL = SQL & " from albaran_variedad inner join albaran on albaran_variedad.numalbar = albaran.numalbar "
    SQL = SQL & " where albaran.fechaalb = " & DBSet(fec, "F")
    SQL = SQL & " group by 1,2 "
    SQL = SQL & " order by 1,2 "
    conn.Execute SQL
    
    SQL = "delete from tmpinformes2 where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    ' kilos volcados esa fecha + kilos que quedan
    SQL = "insert into tmpinformes2 (codusu, codigo1, importe1) "
    SQL = SQL & " select " & vUsu.Codigo & ", aaaa.codvarie, sum(aaaa.kilos) from  "
    SQL = SQL & " (select codvarie, sum(coalesce(numkilos,0)) kilos from trzpalets inner join trzlineas_cargas on trzpalets.idpalet = trzlineas_cargas.idpalet where trzlineas_cargas.fecha = " & DBSet(fec, "F")
    
    '[Monica]07/06/2019: cargamos sólo lo que se ha volcado por la linea 1 ( lo de la linea 2 es destrio )
    SQL = SQL & " and trzlineas_cargas.linea = 1 "
    
    SQL = SQL & " group by 1 "
    SQL = SQL & " union "
    SQL = SQL & " select codvarie, sum(coalesce(kilos,0)) kilos from trzmovim where numalbar is null and esmerma = 0"
    SQL = SQL & " group by 1) aaaa "
    SQL = SQL & " group by 1, 2 "
    conn.Execute SQL


    CadVariedades = ""

    ' montamos un cursor con las variedades que tengan mas kilos salidos que volcados
    SQL = "select tmpinformes.* from tmpinformes where codusu = " & vUsu.Codigo
    SQL = SQL & " order by codigo1"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    While Not Rs.EOF
        KilosVolcados = DevuelveValor("select importe1 from tmpinformes2 where codusu = " & vUsu.Codigo & " and codigo1 = " & DBSet(Rs!Codigo1, "N"))
        
        If KilosVolcados < DBLet(Rs!importe1) Then
            CadVariedades = CadVariedades & DBLet(Rs!Codigo1) & ","
        End If
    
        Rs.MoveNext
    Wend
    
    If CadVariedades <> "" Then
        CadVariedades = Mid(CadVariedades, 1, Len(CadVariedades) - 1)
        MsgBox "No hay suficientes kilos de las siguientes variedades: " & CadVariedades & " fecha " & fec, vbExclamation
        ComprobarExistenciasConAlbaranes = False
        Exit Function
    End If
    
    Set Rs = Nothing

    ComprobarExistenciasConAlbaranes = True
    Exit Function

eComprobarExistenciasConAlbaranes:
    MuestraError Err.Number, "Comprobar Existencias con Albaranes Salida", Err.Description
End Function

Private Function ProcesoCarga(vSQL As String, fec As Date) As Boolean
Dim vMens As String

    On Error GoTo eProcesoCarga
    
    ProcesoCarga = False
    
    conn.BeginTrans
    
    vMens = ""
    If CargarPaletsConfeccionados(vSQL, vMens, fec) Then
        If RepartoAlbaranes(vMens, fec) Then
            ProcesoCarga = True
            conn.CommitTrans
            Exit Function
        End If
    End If
    
eProcesoCarga:
    conn.RollbackTrans
    If vMens <> "" Then
        MuestraError Err.Number, vMens
    Else
        MsgBox "No se ha realizado el proceso de carga", vbExclamation
    End If
End Function

Private Function RepartoAlbaranes(vMens As String, fec As Date) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Salir As Boolean
Dim KilosVar As Long
Dim NumLinea As Integer
Dim resto As Long
Dim vcodigo As Long

    On Error GoTo eRepartoAlbaranes

    RepartoAlbaranes = False

    Label2(8).Caption = "Reparto de albaranes"
    DoEvents
    
    ' para todos los albaranes que han salido repartimos
    SQL = "select albaran.numalbar, albaran.codclien, codvarie, nrotraza, sum(numcajas), sum(pesoneto) pesoneto from albaran_variedad inner join albaran on albaran_variedad.numalbar = albaran.numalbar "
    SQL = SQL & " where albaran.fechaalb = " & DBSet(fec, "F")
    SQL = SQL & " group by 1,2,3,4  "
    SQL = SQL & " order by 1,2,3,4 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Sql2 = "select sum(kilos) from trzmovim where numalbar is null and codvarie = " & DBSet(Rs!Codvarie, "N") & " and esmerma = 0 "
        
        KilosVar = DBLet(Rs!PesoNeto)
        If DevuelveValor(Sql2) < DBLet(Rs!PesoNeto) Then
            MsgBox "No hay suficiente existencias de la variedad " & DBLet(Rs!Codvarie), vbExclamation
            Exit Function
        Else
            Sql2 = "select * from trzmovim where numalbar is null and codvarie = " & DBSet(Rs!Codvarie, "N") & " and esmerma = 0 "
            Sql2 = Sql2 & " order by fecha asc "
            
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            Salir = False
            
            NumLinea = DevuelveValor("select max(coalesce(numlinea, 0)) from albaran_palets where numalbar = " & DBSet(Rs!NumAlbar, "N"))
            
            While Not Rs2.EOF And Not Salir
                NumLinea = NumLinea + 1
                
                SQL = "insert into albaran_palets (numalbar, numlinea, numpalet) values ("
                SQL = SQL & DBSet(Rs!NumAlbar, "N") & "," & DBSet(NumLinea, "N") & "," & DBSet(Rs2!NumPalet, "N") & ")"
                
                conn.Execute SQL
            
                If DBLet(Rs2!Kilos) < KilosVar Then
                    
                    KilosVar = KilosVar - DBLet(Rs2!Kilos)
                    
                    SQL = "update trzmovim set numalbar = " & DBSet(Rs!NumAlbar, "N")
                    SQL = SQL & ", nrotraza = " & DBSet(Rs!nrotraza, "T")
                    SQL = SQL & ", codclien = " & DBSet(Rs!CodClien, "N")
                    SQL = SQL & " where codigo = " & DBSet(Rs2!Codigo, "N")
                    
                    conn.Execute SQL
                Else
                    resto = DBLet(Rs2!Kilos) - KilosVar
                
                    SQL = "update trzmovim set numalbar = " & DBSet(Rs!NumAlbar, "N")
                    SQL = SQL & ", kilos =  " & DBSet(KilosVar, "N")
                    SQL = SQL & ", nrotraza = " & DBSet(Rs!nrotraza, "T")
                    SQL = SQL & ", codclien = " & DBSet(Rs!CodClien, "N")
                    SQL = SQL & " where codigo = " & DBSet(Rs2!Codigo, "N")
                
                    conn.Execute SQL
                    
                    ' insertamos una linea con la diferencia que nos queda
                    If resto <> 0 Then
                        vcodigo = DevuelveValor("select max(coalesce(codigo,0)) from trzmovim")
                        vcodigo = vcodigo + 1
                        
                        SQL = "insert into trzmovim (codigo, numpalet, fecha, codvarie, kilos) values "
                        SQL = SQL & "(" & DBSet(vcodigo, "N") & "," & DBSet(Rs2!NumPalet, "N") & "," & DBSet(Rs2!Fecha, "F") & "," & DBSet(Rs!Codvarie, "N") & ","
                        SQL = SQL & DBSet(resto, "N") & ")"
                        
                        conn.Execute SQL
                    End If
                    
                    Salir = True
                End If
        
                Rs2.MoveNext
            Wend
            
            Set Rs2 = Nothing
        End If
        
        Rs.MoveNext
        
    Wend
    Set Rs = Nothing
    
    RepartoAlbaranes = True
    Exit Function
    
eRepartoAlbaranes:
    vMens = "Reparto de Albaranes " & Err.Description & vbCrLf
    
End Function

Private Function CargarPaletsConfeccionados(vSQL As String, vMens As String, fec As Date) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim SqlInsert As String
Dim SqlInsert2 As String
Dim SqlInsert3 As String
Dim SqlValues As String
Dim NroPalet As Long
Dim Marca As Integer
Dim Forfait As String
Dim Calibre As Integer
Dim vcodigo As Long

    On Error GoTo eCargarPaletsConfeccionados

    CargarPaletsConfeccionados = False
    
    Label2(8).Caption = "Carga palets confeccionados"
    DoEvents

    NroPalet = DevuelveValor("select max(numpalet) from palets")
   
    
    SqlInsert = "insert into palets (numpalet,fechaini,horaini,fechafin,horafin,codpalet,linconfe,tipmercan,"
    SqlInsert = SqlInsert & "fechaconf,horaiconf,horafconf,codlinconf,intorden,linentrada,linsalida,idpalet) values "
    
    SqlInsert2 = "insert into palets_variedad (numpalet,numlinea,codvarie,codvarco,codmarca,codforfait,pesobrut,pesoneto,numcajas) values "
    
    SqlInsert3 = "insert into palets_calibre (numpalet,numlinea,numline1,codvarie,codcalib,numcajas) values "
    
    Marca = DevuelveValor("select min(codmarca) from marcas")
    Forfait = DevuelveValor("select min(codforfait) from forfaits")
    vcodigo = DevuelveValor("select max(coalesce(codigo,0)) from trzmovim")
    
    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        NroPalet = NroPalet + 1
        
        SqlValues = "(" & DBSet(NroPalet, "N") & "," & DBSet(fec, "F") & "," & DBSet(fec & " 00:00:00", "FH") & ","
                                                                                   '[Monica]07/06/2019: antes en linea ponia 1
        SqlValues = SqlValues & DBSet(fec, "F") & "," & DBSet(fec & " 00:00:00", "FH") & ",1," & DBSet(Rs!Linea, "N") & ",0,"
        SqlValues = SqlValues & DBSet(fec, "F") & "," & DBSet(fec & " 00:00:00", "FH") & "," & DBSet(fec & " 00:00:00", "FH")
        SqlValues = SqlValues & "," & DBSet(Rs!Linea, "N") & ",1," & DBSet(Rs!Linea, "N") & "," & DBSet(Rs!Linea, "N") & "," '[Monica]07/06/2019: antes ponia 1,1,1,1
        SqlValues = SqlValues & DBSet(Rs!IdPalet, "N") & ")"
    
        conn.Execute SqlInsert & SqlValues
    
        SQL = "select * from trzpalets where idpalet = " & DBSet(Rs!IdPalet, "N")
        
        Set Rs1 = New ADODB.Recordset
        Rs1.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs1.EOF Then
            Calibre = DevuelveValor("select min(codcalib) from calibres where codvarie = " & DBSet(Rs1!Codvarie, "N"))
            
            
            'palets_variedad
            SqlValues = "(" & DBSet(NroPalet, "N") & ",1," & DBSet(Rs1!Codvarie, "N") & "," & DBSet(Rs1!Codvarie, "N") & "," & DBSet(Marca, "N") & ","
            SqlValues = SqlValues & DBSet(Forfait, "T") & "," & DBSet(Rs1!NumKilos, "N") & "," & DBSet(Rs1!NumKilos, "N") & "," & DBSet(Rs1!NumCajones, "N") & ")"
            
            conn.Execute SqlInsert2 & SqlValues
            
            'palets_calibre
            SqlValues = "(" & DBSet(NroPalet, "N") & ",1,1," & DBSet(Rs1!Codvarie, "N") & "," & DBSet(Calibre, "N") & "," & DBSet(Rs1!NumCajones, "N") & ")"
            
            conn.Execute SqlInsert3 & SqlValues
        End If
        
        ' metemos en la tabla de movimientos de traza
        vcodigo = vcodigo + 1
        
        SQL = "insert into trzmovim (codigo, numpalet, fecha, codvarie, kilos,esmerma) values "
        SQL = SQL & "(" & DBSet(vcodigo, "N") & "," & DBSet(NroPalet, "N") & "," & DBSet(fec, "F") & "," & DBSet(Rs1!Codvarie, "N") & ","
        SQL = SQL & DBSet(Rs1!NumKilos, "N") & ","
        
        '[Monica]07/06/2019: en el caso de que el palet vaya por la linea 2 lo marcamos como merma, antes no grababamos el campo esmerma
        If DBLet(Rs!Linea, "N") = 1 Then
            SQL = SQL & "0"
        Else
            SQL = SQL & "1"
        End If
        
        SQL = SQL & ")"
        
        conn.Execute SQL
        
        Set Rs1 = Nothing
        Rs.MoveNext
    Wend
    Set Rs = Nothing

    CargarPaletsConfeccionados = True
    
    Exit Function

eCargarPaletsConfeccionados:
    vMens = "Cargar Palets Confeccionados:" & vbCrLf & Err.Description
End Function


Private Sub CmdCancelT_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtCodigo(1)
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

    tabla = "rhisfruta"
    
    ' Necesitamos la conexion a la contabilidad de la seccion de adv
    ' para sacar los porcentajes de iva de los articulos y calcular
    ' los datos de la factura
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
'    Me.cmdCancel(indFrame).Cancel = True
'    Me.Width = w + 70
'    Me.Height = h + 350
End Sub



Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtCodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim SQL As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        SQL = " {variedades.codvarie} in (" & CadenaSeleccion & ")"
        Sql2 = " {variedades.codvarie} in [" & CadenaSeleccion & "]"
    Else
        SQL = " {variedades.codvarie} = -1 "
    End If
    If Not AnyadirAFormula(cadselect, SQL) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub


Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0 'VARIEDADES
            AbrirFrmArticuloADV (Index)
    
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub


Private Sub imgFec_Click(Index As Integer)
Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object
    Dim Indice As Integer

    Set frmC = New frmCal

    esq = imgFec(Index).Left
    dalt = imgFec(Index).Top

    Set obj = imgFec(Index).Container

      While imgFec(Index).Parent.Name <> obj.Name
            esq = esq + obj.Left
            dalt = dalt + obj.Top
            Set obj = obj.Container
      Wend

    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40

    Select Case Index
        Case 0
            Indice = 16
        Case 1
            Indice = 17
    End Select

    imgFec(0).Tag = Indice '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtCodigo(Indice).Text <> "" Then frmC.NovaData = txtCodigo(Indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtCodigo(CByte(imgFec(0).Tag)) '<===
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
            Case 0: KEYBusqueda KeyAscii, 0 'variedad desde
        End Select
    Else
        KEYpress KeyAscii
    End If

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
        Case 16, 17 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)

    End Select
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
        .EnvioEMail = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmArticuloADV(Indice As Integer)
    indCodigo = Indice
    Set frmArtADV = New frmADVArticulos
    frmArtADV.DatosADevolverBusqueda = "0|1|"
    frmArtADV.Show vbModal
    Set frmArtADV = Nothing
End Sub

Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
'        .SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
        .OtrosParametros = cadParam
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
    Unload Me
End Sub

Private Sub AbrirEMail()
    If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean
Dim SQL As String
Dim Sql2 As String
Dim vClien As cSocio
' añadido
Dim Mens As String
Dim numfactu As String
Dim numser As String
Dim Fecha As Date
Dim cad As String

    B = True
    
    If B Then
        If (txtCodigo(1).Text = "" Or txtCodigo(2).Text = "") Then
            MsgBox "El rango de albaranes debe de tener un valor. Reintroduzca.", vbExclamation
            B = False
        End If
    End If
    
    DatosOk = B

End Function




Private Function ActualizarRegistros(cTabla As String, cWhere As String) As Boolean
'Actualizar la marca de impreso
Dim SQL As String

    On Error GoTo eActualizarRegistros

    ActualizarRegistros = False
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    SQL = "update " & QuitarCaracterACadena(cTabla, "_1") & " set impreso = 1 "
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    conn.Execute SQL
    
    ActualizarRegistros = True
    Exit Function
    
eActualizarRegistros:
    MuestraError Err.Number, "Actualizando registros", Err.Description
End Function



Public Function GeneraRegistros(vDesde As String, vHasta As String) As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' ariges.scafac --> conta.cabfact
' ariges.slifac --> conta.linfact
'Actualizar la tabla ariges.scafac.inconta=1 para indicar que ya esta contabilizada
Dim B As Boolean
Dim cadMen As String
Dim SQL As String
Dim NumF As Currency
Dim vFactADV As CFacturaADV
Dim vSocio As cSocio
Dim Rs As ADODB.Recordset
Dim cadErr As String


Dim Desde As Long
Dim Hasta As Long
Dim i As Long

    On Error GoTo EContab

    conn.BeginTrans
    
    B = True
    
    Desde = CLng(vDesde)
    Hasta = CLng(vHasta)
    
    SQL = "select numalbar from rhisfruta where tipoentr <> 1 and numalbar >=  " & DBSet(Desde, "N")
    SQL = SQL & " and numalbar <= " & DBSet(Hasta, "N")
    SQL = SQL & " order by numalbar "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF And B
        Label2(2).Caption = Rs!NumAlbar
        DoEvents
    
        B = CalculoGastosTransporte(Rs!NumAlbar, cadErr)
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
'antes
'    I = vDesde
'    While I <= vHasta And b
'
'            If DevuelveValor("select tipoentr from rhisfruta where numalbar =" & I) <> 1 Then
'                '[Monica]27/04/2010 Calculo de costes de transporte, si es por tarifas y la entrada no es de venta campo
'                b = CalculoGastosTransporte(I, cadErr)
'            End If
'        I = I + 1
'    Wend
    
EContab:
    
    If Err.Number <> 0 Then
        B = False
        MuestraError Err.Number, "Modificando Registros", Err.Description & " " & cadErr
    End If
    If B Then
        conn.CommitTrans
        GeneraRegistros = True
    Else
        conn.RollbackTrans
        GeneraRegistros = False
    End If
End Function

Private Function CalculoGastosTransporte(Albaran As Long, cadErr As String) As Boolean
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

    If vParamAplic.Cooperativa <> 2 And vParamAplic.Cooperativa <> 16 Then
        SQL = "select numnotac, rhisfruta_entradas.codtarif, rtarifatra.tipotarifa, sum(kilosnet) as kilos "
    Else
        SQL = "select numnotac, rhisfruta_entradas.codtarif, rtarifatra.tipotarifa, sum(kilostra) as kilos "
    End If
    SQL = SQL & " from rhisfruta_entradas, rtarifatra where numalbar = " & DBSet(Albaran, "N")
    SQL = SQL & " and rtarifatra.tipotarifa <> 2 " 'las tarifas que buscamos son del tipo 1 o 2 (no sin asignar)
    SQL = SQL & " and rtarifatra.codtarif = rhisfruta_entradas.codtarif "
    SQL = SQL & " group by 1, 2, 3 "
    SQL = SQL & " order by 1, 2, 3 "

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
        If vParamAplic.Cooperativa <> 2 And vParamAplic.Cooperativa <> 16 Then
            SQL = "update rhisfruta_entradas set imptrans = " & DBSet(ImpTrans, "N")
            SQL = SQL & " where numalbar = " & DBSet(Albaran, "N") & " and numnotac = " & DBSet(Rs!NumNotac, "N")
            
            conn.Execute SQL
        End If
        If DBLet(Rs!tipotarifa, "N") = 0 Then
            ImpGastoSocio = ImpGastoSocio + Round2((DBLet(Rs!Kilos, "N") * (PrecTarifa - PrecTarifaAlm)), 2)
        Else
            ImpGastoSocio = ImpGastoSocio + Round2((DBLet(Rs!Kilos, "N") * (PrecTarifa - PrecTarifaAlm2)), 2)
        End If
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    If vParamAplic.Cooperativa <> 2 And vParamAplic.Cooperativa <> 16 Then
        ' actualizamos cabecera
        SQL = "update rhisfruta set imptrans = " & DBSet(TotImpTrans, "N")
        SQL = SQL & " where numalbar = " & DBSet(Albaran, "N")
        
        conn.Execute SQL
    End If
    
    If ImpGastoSocio <> 0 Then
        ' si existe registro en la tabla rhisfruta_gastos de concepto codgastotra actualizamos el importe
        SQL = "select count(*) from rhisfruta_gastos where numalbar = " & DBSet(Albaran, "N")
        SQL = SQL & " and codgasto = " & DBSet(vParamAplic.CodGastoTRA, "N")
        
        If TotalRegistros(SQL) = 0 Then
        
            NumF = ""
            NumF = SugerirCodigoSiguienteStr("rhisfruta_gastos", "numlinea", "numalbar = " & DBSet(Albaran, "N"))
            ' grabamos un registro en con los gastos del cliente
            SQL = "insert into rhisfruta_gastos (numalbar, numlinea, codgasto, importe) values (" & DBSet(Albaran, "N") & ","
            SQL = SQL & DBSet(NumF, "N") & "," & DBSet(vParamAplic.CodGastoTRA, "N") & "," & DBSet(ImpGastoSocio, "N") & ")"
            
            conn.Execute SQL
            
        Else
        
            ' acualizamos el registro que hay
            SQL = "update rhisfruta_gastos set importe = " & DBSet(ImpGastoSocio, "N")
            SQL = SQL & " where numalbar = " & DBSet(Albaran, "N")
            SQL = SQL & " and codgasto = " & DBSet(vParamAplic.CodGastoTRA, "N")
            
            conn.Execute SQL
        
        End If
    End If
EInsertar:
    If Err.Number <> 0 Then
        CalculoGastosTransporte = False
        cadErr = Err.Description
    Else
        CalculoGastosTransporte = True
    End If
End Function


