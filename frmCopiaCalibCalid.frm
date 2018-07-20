VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCopiaCalibCalid 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   8160
   Icon            =   "frmCopiaCalibCalid.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameInfArticulos 
      Height          =   4350
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   8070
      Begin VB.CheckBox Check1 
         Caption         =   "Calidades/Calibrador"
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
         Index           =   2
         Left            =   540
         TabIndex        =   12
         Top             =   2970
         Width           =   2400
      End
      Begin VB.Frame FrameStockMaxMin 
         Caption         =   "Tipo"
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
         Height          =   1050
         Left            =   3180
         TabIndex        =   9
         Top             =   2190
         Width           =   2085
         Begin VB.OptionButton Opcion 
            Caption         =   "Actualizar"
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
            Index           =   1
            Left            =   450
            TabIndex        =   11
            Top             =   570
            Width           =   1560
         End
         Begin VB.OptionButton Opcion 
            Caption         =   "Copiar"
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
            Index           =   0
            Left            =   450
            TabIndex        =   10
            Top             =   270
            Width           =   1485
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Calibres"
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
         Left            =   540
         TabIndex        =   8
         Top             =   2220
         Width           =   2130
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Calidades"
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
         Left            =   540
         TabIndex        =   7
         Top             =   2580
         Width           =   2130
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
         Index           =   70
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "Text5"
         Top             =   1620
         Width           =   4305
      End
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
         Index           =   70
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1620
         Width           =   1455
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
         Left            =   5400
         TabIndex        =   2
         Top             =   3645
         Width           =   1065
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
         Left            =   6570
         TabIndex        =   3
         Top             =   3645
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Copia de Calibres / Calidades a Variedad"
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
         TabIndex        =   6
         Top             =   495
         Width           =   6735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   27
         Left            =   1425
         Top             =   1620
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Variedad Destino"
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
         Index           =   38
         Left            =   510
         TabIndex        =   5
         Top             =   1305
         Width           =   1665
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   8730
      Top             =   5580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCopiaCalibCalid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NumCod As String 'Variedad origen

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto


Private HaDevueltoDatos As Boolean

Private WithEvents frmVar As frmManVariedad
Attribute frmVar.VB_VarHelpID = -1

Dim PrimeraVez As Boolean
Dim indFrame As Single
Dim indCodigo As Integer

Private Sub KEYpress(KeyAscii As Integer)
    Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub cmdAceptar_Click()
Dim SQL As String
Dim SQL1 As String
    
    
    
'    If BloqueaRegistro("variedades", "codvarie = " & DBSet(txtCodigo(70).Text, "N")) Then

    If DatosOk Then
        '[Monica]20/03/2013: tanto si copiamos como si actualizamos hemos de borrar los datos del destino
        If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
            If ActualizarRegistrosNew Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                cmdCancel_Click
            End If
        Else
            If ActualizarRegistros Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                cmdCancel_Click
            End If
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtcodigo(70)
        Me.Opcion(0).Value = True
        Check1(0).Value = 1
        Check1(1).Value = 1
        Check1(2).Value = 1
    End If
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim i As Integer
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me

    For i = 27 To 27
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i

    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdcancel.Cancel = True
'    Me.Width = W + 70
'    Me.Height = H + 350
End Sub



Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de variedades
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgBuscar_Click(Index As Integer)
'Buscar general: cada index llama a una tabla
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 27 'cod. de variedad
            indCodigo = 70
            Set frmVar = New frmManVariedad
            frmVar.DatosADevolverBusqueda = "0|1|" 'Abrimos en Modo Busqueda
            frmVar.DeConsulta = True
            frmVar.Show vbModal
            Set frmVar = Nothing
            
    End Select
    PonerFoco txtcodigo(indCodigo)
    Screen.MousePointer = vbDefault
End Sub



Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtcodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Tabla As String
Dim codCampo As String, nomCampo As String
Dim TipCampo As String, Formato As String
Dim Titulo As String
Dim EsNomCod As Boolean 'Si es campo Cod-Descripcion llama a PonerNombreDeCod


    'Quitar espacios en blanco por los lados
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
        
    Select Case Index
        Case 70  'Cod.variedad
            If txtcodigo(Index).Text <> "" Then
                txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
                txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "variedades", "nomvarie", "codvarie", "N")
                If txtNombre(Index).Text = "" Then
                    MsgBox "Variedad no existe. Reintroduzca.", vbExclamation
                    PonerFoco txtcodigo(Index)
                End If
            End If
    End Select
    
End Sub

Private Function DatosOk() As Boolean
Dim i As Integer

    DatosOk = False
    If txtcodigo(70).Text = "" Then
        MsgBox "Debe de introducir una variedad destino.", vbExclamation
        Exit Function
    Else
        If Check1(0).Value = 0 And Check1(1).Value = 0 Then
            MsgBox "Debe seleccionar Calibres, Calidades o ambas", vbExclamation
            Exit Function
        End If
    End If
      
    'Llegados aqui OK
    DatosOk = True
        
End Function


Private Function ActualizarRegistros() As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset

    On Error GoTo eActualizarRegistros

    ActualizarRegistros = False

    If Check1(0).Value Then ' calibres
        If BloqueaRegistro("calibres", "codvarie = " & DBSet(txtcodigo(70).Text, "N")) Then
            conn.BeginTrans
            If Opcion(0).Value Then ' copiar
                SQL = "select * from calibres where codvarie = " & DBSet(NumCod, "N")
                
                Set Rs = New ADODB.Recordset
                Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If Not Rs.EOF Then Rs.MoveFirst
                While Not Rs.EOF
                    Sql2 = "select count(*) from calibres where codvarie = " & DBSet(txtcodigo(70).Text, "N")
                    Sql2 = Sql2 & " and codcalib = " & DBSet(Rs!codcalib, "N")
                    
                    If TotalRegistros(Sql2) > 0 Then
                        ' updateamos
                        Sql3 = "update calibres fuente, calibres destino set destino.nomcalib = fuente.nomcalib, "
                        Sql3 = Sql3 & " destino.nomcalab = fuente.nomcalab, destino.calbaneco = fuente.calbaneco "
                        Sql3 = Sql3 & " where fuente.codvarie = " & DBSet(NumCod, "N")
                        Sql3 = Sql3 & " and destino.codvarie = " & DBSet(txtcodigo(70).Text, "N")
                        Sql3 = Sql3 & " and fuente.codcalib = " & DBSet(Rs!codcalib, "N")
                        Sql3 = Sql3 & " and destino.codcalib = " & DBSet(Rs!codcalib, "N")
                
                        conn.Execute Sql3
                    Else
                        ' insertamos
                        Sql3 = "insert into calibres (codvarie,codcalib,nomcalib,nomcalab,calbaneco) select " & DBSet(txtcodigo(70).Text, "N")
                        Sql3 = Sql3 & ",codcalib, nomcalib, nomcalab, calbaneco from calibres "
                        Sql3 = Sql3 & " where codvarie = " & DBSet(NumCod, "N")
                        Sql3 = Sql3 & " and codcalib = " & DBSet(Rs!codcalib, "N")
        
                        conn.Execute Sql3
                    End If
                    
                    Rs.MoveNext
                Wend
                
                Set Rs = Nothing
                
'                Sql = "delete from calibres where codvarie = " & DBSet(txtCodigo(70).Text, "N")
'                Conn.Execute Sql
'
'                Sql = "insert into calibres select " & DBSet(txtCodigo(70).Text, "N")
'                Sql = Sql & ",codcalib, nomcalib, nomcalab, calbaneco from calibres "
'                Sql = Sql & " where codvarie = " & DBSet(NumCod, "N")
'
'                Conn.Execute Sql
            Else
                SQL = "update calibres fuente, calibres destino set destino.nomcalib = fuente.nomcalib, "
                SQL = SQL & " destino.nomcalab = fuente.nomcalab, destino.calbaneco = fuente.calbaneco "
                SQL = SQL & " where fuente.codvarie = " & DBSet(NumCod, "N")
                SQL = SQL & " and destino.codvarie = " & DBSet(txtcodigo(70).Text, "N")
                SQL = SQL & " and fuente.codcalib = destino.codcalib "
                
                conn.Execute SQL
            End If
            conn.CommitTrans
        End If
        TerminaBloquear
    End If
    
    If Check1(1).Value Then ' calidades
        If BloqueaRegistro("rcalidad", "codvarie = " & DBSet(txtcodigo(70).Text, "N")) Then
            conn.BeginTrans
            If Opcion(0).Value Then ' copiar
                SQL = "select * from rcalidad where codvarie = " & DBSet(NumCod, "N")
                
                Set Rs = New ADODB.Recordset
                Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If Not Rs.EOF Then Rs.MoveFirst
                While Not Rs.EOF
                    Sql2 = "select count(*) from rcalidad where codvarie = " & DBSet(txtcodigo(70).Text, "N")
                    Sql2 = Sql2 & " and codcalid = " & DBSet(Rs!codcalid, "N")
                    
                    If TotalRegistros(Sql2) > 0 Then
                        ' actualizamos
                        Sql3 = "update rcalidad fuente, rcalidad destino set destino.nomcalid = fuente.nomcalid, "
                        Sql3 = Sql3 & " destino.nomcalab = fuente.nomcalab, destino.tipcalid = fuente.tipcalid, "
                        Sql3 = Sql3 & " destino.tipcalid1 = fuente.tipcalid1, "
                        Sql3 = Sql3 & " destino.nomcalibrador1 = fuente.nomcalibrador1, "
                        Sql3 = Sql3 & " destino.nomcalibrador2 = fuente.nomcalibrador2, "
                        Sql3 = Sql3 & " destino.gastosrec = fuente.gastosrec, "
                        '[Monica]12/11/2013
                        Sql3 = Sql3 & " destino.eurrecsoc = fuente.eurrecsoc,"
                        Sql3 = Sql3 & " destino.eurreccoop = fuente.eurreccoop, "
                        '[Monica]27/01/2016: nueva columna de si se aplica bonficacion
                        Sql3 = Sql3 & " destino.seaplicabonif = fuente.seaplicabonif "
                        Sql3 = Sql3 & " where fuente.codvarie = " & DBSet(NumCod, "N")
                        Sql3 = Sql3 & " and destino.codvarie = " & DBSet(txtcodigo(70).Text, "N")
                        Sql3 = Sql3 & " and fuente.codcalid = " & DBSet(Rs!codcalid, "N")
                        Sql3 = Sql3 & " and destino.codcalid = " & DBSet(Rs!codcalid, "N")
                        
                        conn.Execute Sql3
                        
                    Else
                        ' copiamos
                        Sql3 = "insert into rcalidad (codvarie,codcalid,nomcalid,nomcalab,tipcalid,tipcalid1,nomcalibrador1,nomcalibrador2,gastosrec,eurrecsoc,eurreccoop,seaplicabonif) select " & DBSet(txtcodigo(70).Text, "N")
                        Sql3 = Sql3 & ",codcalid, nomcalid, nomcalab, tipcalid, tipcalid1, nomcalibrador1,"
                        '[Monica]27/01/2016: nueva columna de si se aplica bonificacion
                        Sql3 = Sql3 & "nomcalibrador2, gastosrec,eurrecsoc,eurreccoop,seaplicabonif from rcalidad "
                        Sql3 = Sql3 & " where codvarie = " & DBSet(NumCod, "N")
                        Sql3 = Sql3 & " and codcalid = " & DBSet(Rs!codcalid, "N")
                    
                        conn.Execute Sql3
                        
                    End If
                    Rs.MoveNext
                Wend
                Set Rs = Nothing


'                Sql = "delete from rcalidad where codvarie = " & DBSet(txtCodigo(70).Text, "N")
'                Conn.Execute Sql
'
'                Sql = "insert into rcalidad select " & DBSet(txtCodigo(70).Text, "N")
'                Sql = Sql & ",codcalid, nomcalid, nomcalab, tipcalid, tipcalid1, nomcalibrador1,"
'                Sql = Sql & "nomcalibrador2, gastosrec from rcalidad "
'                Sql = Sql & " where codvarie = " & DBSet(NumCod, "N")
'
'                Conn.Execute Sql


            Else
                SQL = "update rcalidad fuente, rcalidad destino set destino.nomcalid = fuente.nomcalid, "
                SQL = SQL & " destino.nomcalab = fuente.nomcalab, destino.tipcalid = fuente.tipcalid, "
                SQL = SQL & " destino.tipcalid1 = fuente.tipcalid1, "
                SQL = SQL & " destino.nomcalibrador1 = fuente.nomcalibrador1, "
                SQL = SQL & " destino.nomcalibrador2 = fuente.nomcalibrador2, "
                SQL = SQL & " destino.gastosrec = fuente.gastosrec, "
                '[Monica]12/11/2013
                SQL = SQL & " destino.eurrecsoc = fuente.eurrecsoc,"
                SQL = SQL & " destino.eurreccoop = fuente.eurreccoop,"
                '[Monica]27/01/2016: nueva columna de si se aplica bonificacion
                SQL = SQL & " destino.seaplicabonif = fuente.seaplicabonif "
                SQL = SQL & " where fuente.codvarie = " & DBSet(NumCod, "N")
                SQL = SQL & " and destino.codvarie = " & DBSet(txtcodigo(70).Text, "N")
                SQL = SQL & " and fuente.codcalid = destino.codcalid "
                
                conn.Execute SQL
            End If
            conn.CommitTrans
        End If
        TerminaBloquear
    End If


    If Check1(1).Value Then ' calibrador
        If BloqueaRegistro("rcalidad_calibrador", "codvarie = " & DBSet(txtcodigo(70).Text, "N")) Then
            conn.BeginTrans
            If Opcion(0).Value Then ' copiar
                SQL = "select * from rcalidad_calibrador where codvarie = " & DBSet(NumCod, "N")
                
                Set Rs = New ADODB.Recordset
                Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If Not Rs.EOF Then Rs.MoveFirst
                While Not Rs.EOF
                    Sql2 = "select count(*) from rcalidad_calibrador where codvarie = " & DBSet(txtcodigo(70).Text, "N")
                    Sql2 = Sql2 & " and codcalid = " & DBSet(Rs!codcalid, "N")
                    Sql2 = Sql2 & " and numlinea = " & DBSet(Rs!NumLinea, "N")
                    
                    
                    If TotalRegistros(Sql2) > 0 Then
                        ' actualizamos
                        Sql3 = "update rcalidad_calibrador fuente, rcalidad_calibrador destino set "
                        Sql3 = Sql3 & " destino.nomcalibrador1 = fuente.nomcalibrador1, "
                        Sql3 = Sql3 & " destino.nomcalibrador2 = fuente.nomcalibrador2, "
                        Sql3 = Sql3 & " destino.nomcalibrador3 = fuente.nomcalibrador3 "
                        Sql3 = Sql3 & " where fuente.codvarie = " & DBSet(NumCod, "N")
                        Sql3 = Sql3 & " and destino.codvarie = " & DBSet(txtcodigo(70).Text, "N")
                        Sql3 = Sql3 & " and fuente.codcalid = " & DBSet(Rs!codcalid, "N")
                        Sql3 = Sql3 & " and destino.codcalid = " & DBSet(Rs!codcalid, "N")
                        Sql3 = Sql3 & " and fuente.numlinea = " & DBSet(Rs!NumLinea, "N")
                        Sql3 = Sql3 & " and destino.numlinea = " & DBSet(Rs!NumLinea, "N")
                        
                        
                        conn.Execute Sql3
                        
                    Else
                        ' copiamos
                        Sql3 = "insert into rcalidad_calibrador (codvarie,codcalid,numlinea,nomcalibrador1,nomcalibrador2,nomcalibrador3) select " & DBSet(txtcodigo(70).Text, "N")
                        Sql3 = Sql3 & ",codcalid, numlinea, nomcalibrador1, nomcalibrador2, nomcalibrador3 "
                        Sql3 = Sql3 & "from rcalidad_calibrador "
                        Sql3 = Sql3 & " where codvarie = " & DBSet(NumCod, "N")
                        Sql3 = Sql3 & " and codcalid = " & DBSet(Rs!codcalid, "N")
                        Sql3 = Sql3 & " and numlinea = " & DBSet(Rs!NumLinea, "N")
                        
                    
                        conn.Execute Sql3
                        
                    End If
                    Rs.MoveNext
                Wend
                Set Rs = Nothing

            Else
                SQL = "update rcalidad_calibrador fuente, rcalidad_calibrador destino set  "
                SQL = SQL & " destino.nomcalibrador1 = fuente.nomcalibrador1, "
                SQL = SQL & " destino.nomcalibrador2 = fuente.nomcalibrador2, "
                SQL = SQL & " destino.nomcalibrador3 = fuente.nomcalibrador3 "
                SQL = SQL & " where fuente.codvarie = " & DBSet(NumCod, "N")
                SQL = SQL & " and destino.codvarie = " & DBSet(txtcodigo(70).Text, "N")
                SQL = SQL & " and fuente.codcalid = destino.codcalid "
                SQL = SQL & " and fuente.numlinea = destino.numlinea "
                
                conn.Execute SQL
            End If
            conn.CommitTrans
        End If
        TerminaBloquear
    End If





    ActualizarRegistros = True
    Exit Function
    
eActualizarRegistros:
    MuestraError Err.Number, "Actualizar Registros", Err.Description
    conn.RollbackTrans
    TerminaBloquear
End Function




'[Monica]20/03/2013: en el caso de Picassent copiar borra el destino y copia el origen en el destino
Private Function ActualizarRegistrosNew() As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset

    On Error GoTo eActualizarRegistros

    ActualizarRegistrosNew = False

    If Check1(0).Value Then ' calibres
        If BloqueaRegistro("calibres", "codvarie = " & DBSet(txtcodigo(70).Text, "N")) Then
            conn.BeginTrans
            
            If Opcion(0).Value Then ' copiar
                SQL = "delete from calibres where codvarie = " & DBSet(txtcodigo(70).Text, "N")
                conn.Execute SQL
                
                Sql3 = "insert into calibres (codvarie,codcalib,nomcalib,nomcalab,calbaneco) select " & DBSet(txtcodigo(70).Text, "N")
                Sql3 = Sql3 & ",codcalib, nomcalib, nomcalab, calbaneco from calibres "
                Sql3 = Sql3 & " where codvarie = " & DBSet(NumCod, "N")

                conn.Execute Sql3
            Else
                SQL = "update calibres fuente, calibres destino set destino.nomcalib = fuente.nomcalib, "
                SQL = SQL & " destino.nomcalab = fuente.nomcalab, destino.calbaneco = fuente.calbaneco "
                SQL = SQL & " where fuente.codvarie = " & DBSet(NumCod, "N")
                SQL = SQL & " and destino.codvarie = " & DBSet(txtcodigo(70).Text, "N")
                SQL = SQL & " and fuente.codcalib = destino.codcalib "
                
                conn.Execute SQL
            End If
            conn.CommitTrans
        End If
        TerminaBloquear
    End If
    
    If Check1(1).Value Then ' calidades
        If BloqueaRegistro("rcalidad", "codvarie = " & DBSet(txtcodigo(70).Text, "N")) Then
            conn.BeginTrans
            If Opcion(0).Value Then ' copiar
                SQL = "delete from rcalidad_calibrador where codvarie = " & DBSet(txtcodigo(70).Text, "N")
                conn.Execute SQL
            
                SQL = "delete from rcalidad where codvarie = " & DBSet(txtcodigo(70).Text, "N")
                conn.Execute SQL

                SQL = "insert into rcalidad (codvarie,codcalid,nomcalid,nomcalab,tipcalid,tipcalid1,nomcalibrador1,nomcalibrador2,gastosrec, eurrecsoc, eurreccoop, seaplicabonif) select " & DBSet(txtcodigo(70).Text, "N")
                SQL = SQL & ",codcalid, nomcalid, nomcalab, tipcalid, tipcalid1, nomcalibrador1,"
                '[Monica]27/01/2016: nueva columna de si se aplica bonificacion
                SQL = SQL & "nomcalibrador2, gastosrec, eurrecsoc, eurreccoop, seaplicabonif from rcalidad "
                SQL = SQL & " where codvarie = " & DBSet(NumCod, "N")

                conn.Execute SQL

                ' pq en este punto he borrado el calibrador por las referenciales
                SQL = "insert into rcalidad_calibrador (codvarie,codcalid,numlinea,nomcalibrador1,nomcalibrador2,nomcalibrador3) "
                SQL = SQL & " select " & DBSet(txtcodigo(70).Text, "N") & ",codcalid,numlinea,nomcalibrador1,nomcalibrador2,nomcalibrador3 "
                SQL = SQL & " from rcalidad_calibrador where codvarie = " & DBSet(NumCod, "N")

                conn.Execute SQL
            Else
                SQL = "update rcalidad fuente, rcalidad destino set destino.nomcalid = fuente.nomcalid, "
                SQL = SQL & " destino.nomcalab = fuente.nomcalab, destino.tipcalid = fuente.tipcalid, "
                SQL = SQL & " destino.tipcalid1 = fuente.tipcalid1, "
                SQL = SQL & " destino.nomcalibrador1 = fuente.nomcalibrador1, "
                SQL = SQL & " destino.nomcalibrador2 = fuente.nomcalibrador2, "
                SQL = SQL & " destino.gastosrec = fuente.gastosrec, "
                SQL = SQL & " destino.eurrecsoc = fuente.eurrecsoc, "
                SQL = SQL & " destino.eurreccoop = fuente.eurreccoop, "
                '[Monica]27/01/2016: nueva columna de si se aplica bonificacion
                SQL = SQL & " destino.seaplicabonif = fuente.seaplicabonif "
                SQL = SQL & " where fuente.codvarie = " & DBSet(NumCod, "N")
                SQL = SQL & " and destino.codvarie = " & DBSet(txtcodigo(70).Text, "N")
                SQL = SQL & " and fuente.codcalid = destino.codcalid "
                
                conn.Execute SQL
            End If
            conn.CommitTrans
        End If
        TerminaBloquear
    End If


    If Check1(1).Value Then ' calibrador
        If BloqueaRegistro("rcalidad_calibrador", "codvarie = " & DBSet(txtcodigo(70).Text, "N")) Then
            conn.BeginTrans
            If Opcion(0).Value Then ' copiar
                SQL = "delete from rcalidad_calibrador where codvarie = " & DBSet(txtcodigo(70).Text, "N")
                conn.Execute SQL
                
                SQL = "insert into rcalidad_calibrador (codvarie,codcalid,numlinea,nomcalibrador1,nomcalibrador2,nomcalibrador3) "
                SQL = SQL & " select " & DBSet(txtcodigo(70).Text, "N") & ",codcalid,numlinea,nomcalibrador1,nomcalibrador2,nomcalibrador3 "
                SQL = SQL & " from rcalidad_calibrador where codvarie = " & DBSet(NumCod, "N")

                conn.Execute SQL
            Else
                SQL = "update rcalidad_calibrador fuente, rcalidad_calibrador destino set  "
                SQL = SQL & " destino.nomcalibrador1 = fuente.nomcalibrador1, "
                SQL = SQL & " destino.nomcalibrador2 = fuente.nomcalibrador2, "
                SQL = SQL & " destino.nomcalibrador3 = fuente.nomcalibrador3 "
                SQL = SQL & " where fuente.codvarie = " & DBSet(NumCod, "N")
                SQL = SQL & " and destino.codvarie = " & DBSet(txtcodigo(70).Text, "N")
                SQL = SQL & " and fuente.codcalid = destino.codcalid "
                SQL = SQL & " and fuente.numlinea = destino.numlinea "
                
                conn.Execute SQL
            End If
            conn.CommitTrans
        End If
        TerminaBloquear
    End If





    ActualizarRegistrosNew = True
    Exit Function
    
eActualizarRegistros:
    MuestraError Err.Number, "Actualizar Registros", Err.Description
    conn.RollbackTrans
    TerminaBloquear
End Function


