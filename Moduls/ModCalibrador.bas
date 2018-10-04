Attribute VB_Name = "ModCalibrador"
Option Explicit

'[Monica] 22/09/2009 nuevo calibrador grande para Catadau
Public Function ProcesarDirectorioCatadau(nomDir As String, Tipo As Byte, Fecha As String, ByRef Pb1 As ProgressBar, ByRef Label1 As Label, ByRef Label2 As Label) As Boolean
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim NumReg As Long
Dim SQL As String
Dim SQL1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim B As Boolean
Dim NomFic As String
Dim Nota As String
Dim Linea As Integer

    ProcesarDirectorioCatadau = False
    B = True
    ' Muestra los nombres en C:\ que representan directorios.
    NomFic = Dir(nomDir & "*.txt")  ' Recupera la primera entrada.
    
    If Tipo = 0 Then
    'CALIBRADOR GRANDE
        Do While NomFic <> "" And B   ' Inicia el bucle.
           ' Ignora el directorio actual y el que lo abarca.
           If NomFic <> "." And NomFic <> ".." And InStr(1, NomFic, Fecha) <> 0 Then
              ' Realiza una comparaci�n a nivel de bit para asegurarse de que MiNombre es un directorio.
              If (GetAttr(nomDir & NomFic) And vbArchive) = vbArchive Then
                
                NF = FreeFile
                
                Open nomDir & NomFic For Input As #NF
                
                Line Input #NF, cad
                
                Label1.Caption = "Procesando Fichero: " & NomFic
                'longitud = FileLen(nomDir & NomFic)
                
                Linea = 1
                If cad <> "" Then
                    Nota = DevuelveNota(NF, Linea)
                
                    If Nota <> "" Then
                        '[Monica]08/05/2017: si el numero de nota que me viene no es numerico doy error
                        If Not IsNumeric(Nota) Then
                            MsgBox "El n�mero de nota " & Nota & " del fichero " & NomFic & " no es correcto", vbExclamation
                            B = False
                        Else
                        ' si no hay linea donde me indica el nro de nota no hago nada con el fichero
                            Pb1.visible = True
                            Pb1.Max = Linea  'longitud
                            DoEvents
            '                Refresh
                            Pb1.Value = 0
                        
                            Close #NF
                            Open nomDir & NomFic For Input As #NF
                            Line Input #NF, cad
                        
                            B = ProcesarFicheroCatadauCGrande(NF, cad, Pb1, Label1, Label2, Nota)
                        End If
                    End If
                End If
                
                Close #NF
              End If   ' solamente si representa un directorio.
           End If
           NomFic = Dir   ' Obtiene siguiente entrada.
        Loop
    Else
    'CALIBRADOR PEQUE�O
        Do While NomFic <> "" And B   ' Inicia el bucle.
           ' Ignora el directorio actual y el que lo abarca.
           If NomFic <> "." And NomFic <> ".." Then
              ' Realiza una comparaci�n a nivel de bit para asegurarse de que MiNombre es un directorio.
              If (GetAttr(nomDir & NomFic) And vbArchive) = vbArchive Then
              
                SQL = "delete from tmpcalibrador"
                conn.Execute SQL
              
                SQL = "load data local infile '" & Replace(nomDir & NomFic, "\", "/") & "' into table `tmpcalibrador` fields escaped by '\\' enclosed by '""' lines terminated by '\r\n' ( `numnota`, `fecnota`, `albaran`,`porcen1`,`porcen2`,`kilos1`, `kilos2`, `kilos3`,`numnota2`,`export`,`nomcalid`,`kilos4`,`kilos5`)  "
                conn.Execute SQL
                
                SQL = "delete from tmpcalibrador where numnota = ''"
                conn.Execute SQL
                
                Label1.Caption = "Procesando Fichero: " & NomFic
                longitud = TotalRegistros("select count(*) from tmpcalibrador")
                
                Pb1.visible = True
                Pb1.Max = longitud
                'Me.Refresh
                DoEvents
                Pb1.Value = 0
                    
                If longitud <> 0 Then
                    B = ProcesarFicheroCatadauCPeque�o(Pb1, Label1, Label2)
                End If
                
              End If   ' solamente si representa un directorio.
           End If
           NomFic = Dir   ' Obtiene siguiente entrada.
        Loop
    
    
    End If
    
    ProcesarDirectorioCatadau = B
    
    Pb1.visible = False
    Label1.Caption = ""
    Label2.Caption = ""
                     
End Function


'[Monica]25/09/2009: han cambiado el CALIBRADOR GRANDE de catadau. Cada fichero se corresponde con
'                    una nota de entrada.
'        19/10/2009: el calibrador peque�o no se corresponde con el agre1104
' Proceso de traspaso para CATADAU
'
Private Function ProcesarFicheroCatadauCGrande(NF As Long, cad As String, ByRef Pb1 As ProgressBar, ByRef Label1 As Label, ByRef Label2 As Label, ByRef Nota As String) As Boolean
Dim B As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Mens As String
Dim numlinea As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim RSaux As ADODB.Recordset

Dim codsoc As String
Dim Codcam As String
Dim codpro As String
Dim codVar As String
Dim Observ As String
Dim Notaca As String
Dim Kilone As String

Dim Destri As String
Dim Podrid As String
Dim Pequen As String
Dim Muestra As String

Dim NGrupos As String
Dim Nombre1 As String



Dim i As Integer
Dim Situacion As Byte

Dim NomCal As Dictionary
Dim KilCal As Dictionary

Dim SQLaux As String
Dim Nregs As Integer

Dim NSep As Integer

Dim SeInserta As Boolean
Dim KilosTot As Currency
Dim cantidad As Long
Dim Kilos As Currency

Dim HoraIni As String
Dim HoraFin As String

Dim FechaEnt As String
Dim UltimaLinea As Boolean
Dim NroCalidad As Integer

Dim Porcen As String
Dim KilosMuestreo As String
Dim HayReg As Boolean

    On Error GoTo eProcesarFicheroCatadau

    ProcesarFicheroCatadauCGrande = False
    
    codsoc = 0
    Codcam = 0
    codpro = 0
    codVar = 0
    Observ = ""
    Notaca = 0
    Kilone = 0
    KilosTot = 0

    Destri = 0
    Podrid = 0
    Pequen = 0
    
    i = 0
    
    ' inicializamos las variables
    Set NomCal = New Dictionary
    Set KilCal = New Dictionary
            
    Notaca = Nota 'RecuperaValorNew(cad, ";", 1)
    
    SQL = "select kilosnet, codvarie, codcampo, codsocio from rclasifica where numnotac = " & DBSet(Notaca, "N")
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Situacion = 0
    If Rs.EOF Then
        Observ = "NOTA NO EXISTE"
        Situacion = 2
    End If
    
    B = True
    UltimaLinea = False
    NroCalidad = 0
    While Not EOF(NF)
        i = i + 1
        
        Pb1.Value = Pb1.Value + 1 ' Len(Cad)
        Label2.Caption = "Linea " & i
        DoEvents
        
        NSep = NumeroSubcadenasInStr(cad, ";")
        
        If NSep = 14 Then ' estamos en una calidad
            NroCalidad = NroCalidad + 1
            
            Nombre1 = RecuperaValorNew(cad, ";", 4)
            Kilone = RecuperaValorNew(cad, ";", 7)
            
            Kilos = Round2(CCur(Kilone) / 1000, 2)
            
            cantidad = RecuperaValorNew(cad, ";", 8)
            KilosTot = KilosTot + Kilos
            
            If Situacion <> 2 Then
                ' si hay nota asociada busco los datos
                SQL = "select codcalid from rcalidad_calibrador where codvarie = " & DBSet(Rs!Codvarie, "N")
                SQL = SQL & " and nomcalibrador1 = " & DBSet(Nombre1, "T")
                
                Set RS1 = New ADODB.Recordset
                RS1.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If RS1.EOF Then
                    Observ = "NO EXIS.CAL"
                    Situacion = 1
                Else
                    NomCal(i) = DBLet(RS1!codcalid, "N")
                    KilCal(i) = Kilos
                End If
                Set RS1 = Nothing
            
            End If
        End If
        
        
        Line Input #NF, cad
    Wend
    
    If cad <> "" Then
'        pb1.Value = pb1.Value + 1 'Len(Cad)
'        Label2.Caption = "Linea " & I
'        'Me.Refresh
'        DoEvents
        
        NSep = NumeroSubcadenasInStr(cad, ";")

        If NSep = 15 Then ' estamos en la ultima linea
            HoraIni = RecuperaValorNew(cad, ";", 9)
            HoraFin = RecuperaValorNew(cad, ";", 10)
            FechaEnt = RecuperaValorNew(cad, ";", 11)
            
            Destri = RecuperaValorNew(cad, ";", 12)
            Podrid = RecuperaValorNew(cad, ";", 15)
            
        End If
    End If
    
    Close #NF
        
'    If DBLet(Rs.Fields(0).Value, "N") <> KilosTot Then
'        Observ = "K.NETOS DIF."
'        Situacion = 4
'    End If


    SQL = "select count(*) from rclasifauto where numnotac = " & Notaca

    SeInserta = (TotalRegistros(SQL) = 0)

    If SeInserta Then
        If Situacion = 2 Then
            ' si no hay nota asociada no puedo meter la clasificacion
            SQL = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
            SQL = SQL & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
            SQL = SQL & "`observac`,`situacion`) values ("
            SQL = SQL & DBSet(Notaca, "N") & ","
            SQL = SQL & DBSet(0, "N") & ","
            SQL = SQL & DBSet(0, "N") & ","
            SQL = SQL & DBSet(0, "N") & ","
            SQL = SQL & DBSet(KilosTot, "N") & ","
            SQL = SQL & DBSet(Destri, "N") & ","
            SQL = SQL & DBSet(Podrid, "N") & ","
            SQL = SQL & DBSet(Pequen, "N") & ","
            SQL = SQL & DBSet(Observ, "T") & ","
            SQL = SQL & DBSet(Situacion, "N") & ")"

        Else
            ' insertamos en las tablas intermedias: rclasifauto y rclasifauto_clasif
            ' tabla: rclasifauto
            SQL = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
            SQL = SQL & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
            SQL = SQL & "`observac`,`situacion`) values ("
            SQL = SQL & DBSet(Notaca, "N") & ","
            SQL = SQL & DBSet(Rs!Codsocio, "N") & ","
            SQL = SQL & DBSet(Rs!codCampo, "N") & ","
            SQL = SQL & DBSet(Rs!Codvarie, "N") & ","
            SQL = SQL & DBSet(Round2(KilosTot, 0), "N") & ","
            SQL = SQL & DBSet(Destri, "N") & ","
            SQL = SQL & DBSet(Podrid, "N") & ","
            SQL = SQL & DBSet(Pequen, "N") & ","
            SQL = SQL & DBSet(Observ, "T") & ","
            SQL = SQL & DBSet(Situacion, "N") & ")"
        End If
        conn.Execute SQL

        ' tabla: rclasifauto_clasif
        SQL = "insert into rclasifauto_clasif (`numnotac`,`codvarie`,`codcalid`,`kiloscal`) "
        SQL = SQL & " values "

    Else
        SQL = "update rclasifauto set kilospod = kilospod + " & DBSet(Podrid, "N") & ","
        SQL = SQL & " kilosdes = kilosdes + " & DBSet(Destri, "N") & ","
        SQL = SQL & " kilosnet = kilosnet + " & DBSet(KilosTot, "N")
        SQL = SQL & " where numnotac = " & DBSet(Notaca, "N")
        
        conn.Execute SQL
    End If


    'solo si tenemos nota asociada metemos toda la clasificacion
    If Situacion <> 2 Then

        'borramos la tabla temporal
        SQLaux = "delete from tmpcata"
        conn.Execute SQLaux

        ' cargamos la tabla temporal
        For i = 1 To NroCalidad
            If NomCal(i) <> "" Then
                Nregs = TotalRegistros("select count(*) from tmpcata where codcalid = " & DBSet(NomCal(i), "N"))
                If Nregs = 0 Then
                    'insertamos en la temporal
                    SQLaux = "insert into tmpcata (codcalid, kilosnet) values (" & DBSet(NomCal(i), "N")
                    SQLaux = SQLaux & "," & DBSet(KilCal(i), "N") & ")"

                    conn.Execute SQLaux
                Else
                    'actualizamos la temporal
                    SQLaux = "update tmpcata set kilosnet = kilosnet + " & DBSet(KilCal(i), "N")
                    SQLaux = SQLaux & " where codcalid = " & DBSet(NomCal(i), "N")

                    conn.Execute SQLaux
                End If
            End If
        Next i

        SQLaux = "select * from tmpcata order by codcalid"

        Set RSaux = New ADODB.Recordset
        RSaux.Open SQLaux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

        Sql2 = ""

        HayReg = False

        While Not RSaux.EOF
            HayReg = True
            If SeInserta Then
                Sql2 = Sql2 & "(" & DBSet(Notaca, "N") & "," & DBSet(Rs!Codvarie, "N") & ","
                Sql2 = Sql2 & DBSet(RSaux!codcalid, "N") & "," & DBSet(RSaux!KilosNet, "N") & "),"
            Else
                Sql2 = "update rclasifauto_Clasif set kiloscal = kiloscal + " & DBSet(RSaux!KilosNet, "N")
                Sql2 = Sql2 & " where numnotac = " & DBSet(Notaca, "N")
                Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!Codvarie, "N")
                Sql2 = Sql2 & " and codcalid = " & DBSet(RSaux!codcalid, "N")

                conn.Execute Sql2
            End If

            RSaux.MoveNext
        Wend

        Set RSaux = Nothing


        If SeInserta And HayReg Then
            If Sql2 <> "" Then
                Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
            End If
            SQL = SQL & Sql2
            conn.Execute SQL
        End If
    End If ' si la situacion es distinta de 2


    Set Rs = Nothing
    Set NomCal = Nothing
    Set KilCal = Nothing

    ProcesarFicheroCatadauCGrande = True
    Exit Function

eProcesarFicheroCatadau:
    If Err.Number <> 0 Then
        ProcesarFicheroCatadauCGrande = False
        MsgBox "Error en Procesar Linea " & Err.Description, vbExclamation
    End If
End Function

'[Monica]19/10/2009: CALIBRADOR PEQUE�O
' ESTE NO SE CORRESPONDE CON AGRE1104 DE EUROAGRO
Private Function ProcesarFicheroCatadauCPeque�o(ByRef Pb1 As ProgressBar, ByRef Label1 As Label, ByRef Label2 As Label) As Boolean
Dim B As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Mens As String
Dim numlinea As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim RSaux As ADODB.Recordset

Dim codsoc As String
Dim Codcam As String
Dim codpro As String
Dim codVar As String
Dim Observ As String
Dim Notaca As String
Dim Kilone As String

Dim Destri As String
Dim Podrid As String
Dim Pequen As String
Dim Muestra As String

Dim NGrupos As String
Dim Nombre1 As String

Dim i As Integer
Dim J As Integer
Dim Situacion As Byte

Dim NomCal As Dictionary
Dim KilCal As Dictionary

Dim SQLaux As String
Dim Nregs As Integer

Dim NSep As Integer

Dim SeInserta As Boolean
Dim KilosTot As Currency
Dim cantidad As Long
Dim Kilos As Currency

Dim HoraIni As String
Dim HoraFin As String

Dim FechaEnt As String
Dim UltimaLinea As Boolean
Dim NroCalidad As Integer

Dim Porcen As String
Dim KilosMuestreo As String

Dim HayReg As Boolean
    On Error GoTo eProcesarFicheroCatadauCPeque�o

    ProcesarFicheroCatadauCPeque�o = False
    
    codsoc = 0
    Codcam = 0
    codpro = 0
    codVar = 0
    Observ = ""
    Notaca = 0
    Kilone = 0
    KilosTot = 0

    Destri = 0
    Podrid = 0
    Pequen = 0
    
    i = 0
    
    ' inicializamos las variables
    Set NomCal = New Dictionary
    Set KilCal = New Dictionary
            
            
    SQL = "select * from tmpcalibrador "
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Notaca = 0
    If Not Rs.EOF Then
        Notaca = DBLet(Rs.Fields(0).Value, "N")
        
        If Notaca <> 0 Then
            Sql2 = "select kilosnet, codvarie, codcampo, codsocio from rclasifica where numnotac = " & DBSet(Notaca, "N")
            Set RS1 = New ADODB.Recordset
            RS1.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
            If RS1.EOF Then
                Observ = "NOTA NO EXISTE"
                Situacion = 2
            End If
            
            B = True
            
            While Not Rs.EOF
                i = i + 1
                
                Pb1.Value = Pb1.Value + 1
                Label2.Caption = "Linea " & i
'                Me.Refresh
                DoEvents
                
                Nombre1 = DBLet(Rs!nomcalid, "T")
                Destri = DBLet(Rs!Kilos3, "T")
                Podrid = DBLet(Rs!Kilos2, "T")
                'Pequen = DBLet(RS!Kilos4, "T")
'antes calculo de kilos segun porcentaje
'                Kilone = DBLet(RS!Kilos1, "T")
'                Porcen = DBLet(RS!porcen1, "T")
'                Kilos = Round2(CCur(Kilone) * CCur(Porcen) / 100, 2)
'                KilosTot = KilosTot + Kilos
'ahora me guardo el porcentaje
                KilosTot = DBLet(Rs!Kilos1, "T")
                Kilos = DBLet(Rs!porcen1, "T")
                
                If Situacion <> 2 Then
                    ' si hay nota asociada busco los datos
                    SQL = "select codcalid from rcalidad_calibrador where codvarie = " & DBSet(RS1!Codvarie, "N")
                    SQL = SQL & " and nomcalibrador2 = " & DBSet(Nombre1, "T")
                    
                    Set Rs2 = New ADODB.Recordset
                    Rs2.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
                    If Rs2.EOF Then
                        Observ = "NO EXIS.CAL"
                        Situacion = 1
                    Else
                        NomCal(i) = DBLet(Rs2!codcalid, "N")
                        KilCal(i) = Kilos
                    End If
                    Set Rs2 = Nothing
                
                End If
                
                Rs.MoveNext
            Wend
        
            SQL = "select count(*) from rclasifauto where numnotac = " & Notaca
        
            SeInserta = (TotalRegistros(SQL) = 0)
        
            If SeInserta Then
                If Situacion = 2 Then
                    ' si no hay nota asociada no puedo meter la clasificacion
                    SQL = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
                    SQL = SQL & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
                    SQL = SQL & "`observac`,`situacion`) values ("
                    SQL = SQL & DBSet(Notaca, "N") & ","
                    SQL = SQL & DBSet(0, "N") & ","
                    SQL = SQL & DBSet(0, "N") & ","
                    SQL = SQL & DBSet(0, "N") & ","
                    SQL = SQL & DBSet(KilosTot, "N") & ","
                    SQL = SQL & DBSet(Destri, "N") & ","
                    SQL = SQL & DBSet(Podrid, "N") & ","
                    SQL = SQL & DBSet(Pequen, "N") & ","
                    SQL = SQL & DBSet(Observ, "T") & ","
                    SQL = SQL & DBSet(Situacion, "N") & ")"
        
                Else
                    ' insertamos en las tablas intermedias: rclasifauto y rclasifauto_clasif
                    ' tabla: rclasifauto
                    SQL = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
                    SQL = SQL & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
                    SQL = SQL & "`observac`,`situacion`) values ("
                    SQL = SQL & DBSet(Notaca, "N") & ","
                    SQL = SQL & DBSet(RS1!Codsocio, "N") & ","
                    SQL = SQL & DBSet(RS1!codCampo, "N") & ","
                    SQL = SQL & DBSet(RS1!Codvarie, "N") & ","
                    SQL = SQL & DBSet(KilosTot, "N") & ","
                    SQL = SQL & DBSet(Destri, "N") & ","
                    SQL = SQL & DBSet(Podrid, "N") & ","
                    SQL = SQL & DBSet(Pequen, "N") & ","
                    SQL = SQL & DBSet(Observ, "T") & ","
                    SQL = SQL & DBSet(Situacion, "N") & ")"
                End If
                conn.Execute SQL
        
                ' tabla: rclasifauto_clasif
                SQL = "insert into rclasifauto_clasif (`numnotac`,`codvarie`,`codcalid`,`kiloscal`) "
                SQL = SQL & " values "
        
            End If
        End If
        'solo si tenemos nota asociada metemos toda la clasificacion
        If Situacion <> 2 Then
            'borramos la tabla temporal
            SQLaux = "delete from tmpcata"
            conn.Execute SQLaux
    
            ' cargamos la tabla temporal
            For J = 1 To i
                If NomCal(J) <> "" Then
                    Nregs = TotalRegistros("select count(*) from tmpcata where codcalid = " & DBSet(NomCal(J), "N"))
                    If Nregs = 0 Then
                        'insertamos en la temporal
                        SQLaux = "insert into tmpcata (codcalid, kilosnet) values (" & DBSet(NomCal(J), "N")
                        SQLaux = SQLaux & "," & DBSet(KilCal(J), "N") & ")"
    
                        conn.Execute SQLaux
                    Else
                        'actualizamos la temporal
                        SQLaux = "update tmpcata set kilosnet = kilosnet + " & DBSet(KilCal(J), "N")
                        SQLaux = SQLaux & " where codcalid = " & DBSet(NomCal(J), "N")
    
                        conn.Execute SQLaux
                    End If
                End If
            Next J
    
            SQLaux = "select * from tmpcata order by codcalid"
    
            Set RSaux = New ADODB.Recordset
            RSaux.Open SQLaux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
            Sql2 = ""
            HayReg = False
            While Not RSaux.EOF
                HayReg = True
                If SeInserta Then
                    Sql2 = Sql2 & "(" & DBSet(Notaca, "N") & "," & DBSet(RS1!Codvarie, "N") & ","
                    Sql2 = Sql2 & DBSet(RSaux!codcalid, "N") & "," & DBSet(RSaux!KilosNet, "N") & "),"
                Else
                    Sql2 = "update rclasifauto_Clasif set kiloscal = kiloscal + " & DBSet(RSaux!KilosNet, "N")
                    Sql2 = Sql2 & " where numnotac = " & DBSet(Notaca, "N")
                    Sql2 = Sql2 & " and codvarie = " & DBSet(RS1!Codvarie, "N")
                    Sql2 = Sql2 & " and codcalid = " & DBSet(RSaux!codcalid, "N")
    
                    conn.Execute Sql2
                End If
    
                RSaux.MoveNext
            Wend
    
            Set RSaux = Nothing
    
            If SeInserta And HayReg Then
                If Sql2 <> "" Then
                    Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
                End If
                SQL = SQL & Sql2
                conn.Execute SQL
            End If
        End If ' si la situacion es distinta de 2
    
        Set Rs = Nothing
        Set RS1 = Nothing
        Set NomCal = Nothing
        Set KilCal = Nothing
    
        ProcesarFicheroCatadauCPeque�o = True
        Exit Function
        
    End If
            
'    Notaca = Mid(cad, 2, InStr(2, cad, "") + 1)
'
'    SQL = "select kilosnet, codvarie, codcampo, codsocio from rclasifica where numnotac = " & DBSet(Notaca, "N")
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    If RS.EOF Then
'        Observ = "NOTA NO EXISTE"
'        Situacion = 2
'    End If
'
'    b = True
'    UltimaLinea = False
'    NroCalidad = 0
'    While Not EOF(NF) And Not UltimaLinea
'        I = I + 1
'
'        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
'        lblProgres(1).Caption = "Linea " & I
'        Me.Refresh
'
'        NroCalidad = NroCalidad + 1
'        Nombre1 = DevuelveNomCalidad(cad, 71)
''        Nombre1 = Mid(cad, 71, InStr(55, cad, "export") + 10)
'        KilosMuestreo = Mid(cad, 44, 6)
'
'        Porcen = Mid(cad, 34, 5)
'
''        Kilone = Round2(porcen * kilosmuestreo / 100, 2)
'
'        KilosTot = KilosTot + Kilone
'
'        If Situacion <> 2 Then
'            ' si hay nota asociada busco los datos
'            SQL = "select codcalid from rcalidad_calibrador where codvarie = " & DBSet(RS!CodVarie, "N")
'            SQL = SQL & " and nomcalibrador2 = " & DBSet(Nombre1, "T")
'
'            Set RS1 = New ADODB.Recordset
'            RS1.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'            If RS1.EOF Then
'                Observ = "NO EXIS.CAL"
'                Situacion = 1
'            Else
'                NomCal(I) = DBLet(RS1!codcalid, "N")
'                KilCal(I) = Kilos
'            End If
'            Set RS1 = Nothing
'        End If
'
'        Line Input #NF, cad
'    Wend
'
'    Close #NF
'
'    SQL = "select count(*) from rclasifauto where numnotac = " & Notaca
'
'    SeInserta = (TotalRegistros(SQL) = 0)
'
'    If SeInserta Then
'        If Situacion = 2 Then
'            ' si no hay nota asociada no puedo meter la clasificacion
'            SQL = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
'            SQL = SQL & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
'            SQL = SQL & "`observac`,`situacion`) values ("
'            SQL = SQL & DBSet(Notaca, "N") & ","
'            SQL = SQL & DBSet(0, "N") & ","
'            SQL = SQL & DBSet(0, "N") & ","
'            SQL = SQL & DBSet(0, "N") & ","
'            SQL = SQL & DBSet(KilosTot, "N") & ","
'            SQL = SQL & DBSet(Destri, "N") & ","
'            SQL = SQL & DBSet(Podrid, "N") & ","
'            SQL = SQL & DBSet(Pequen, "N") & ","
'            SQL = SQL & DBSet(Observ, "T") & ","
'            SQL = SQL & DBSet(Situacion, "N") & ")"
'
'        Else
'            ' insertamos en las tablas intermedias: rclasifauto y rclasifauto_clasif
'            ' tabla: rclasifauto
'            SQL = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
'            SQL = SQL & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
'            SQL = SQL & "`observac`,`situacion`) values ("
'            SQL = SQL & DBSet(Notaca, "N") & ","
'            SQL = SQL & DBSet(RS!Codsocio, "N") & ","
'            SQL = SQL & DBSet(RS!CodCampo, "N") & ","
'            SQL = SQL & DBSet(RS!CodVarie, "N") & ","
'            SQL = SQL & DBSet(Round2(KilosTot, 0), "N") & ","
'            SQL = SQL & DBSet(Destri, "N") & ","
'            SQL = SQL & DBSet(Podrid, "N") & ","
'            SQL = SQL & DBSet(Pequen, "N") & ","
'            SQL = SQL & DBSet(Observ, "T") & ","
'            SQL = SQL & DBSet(Situacion, "N") & ")"
'        End If
'        conn.Execute SQL
'
'        ' tabla: rclasifauto_clasif
'        SQL = "insert into rclasifauto_clasif (`numnotac`,`codvarie`,`codcalid`,`kiloscal`) "
'        SQL = SQL & " values "
'
'    End If
'
'    'solo si tenemos nota asociada metemos toda la clasificacion
'    If Situacion <> 2 Then
'
'        'borramos la tabla temporal
'        SQLaux = "delete from tmpcata"
'        conn.Execute SQLaux
'
'        ' cargamos la tabla temporal
'        For I = 1 To NroCalidad
'            If NomCal(I) <> "" Then
'                nRegs = TotalRegistros("select count(*) from tmpcata where codcalid = " & DBSet(NomCal(I), "N"))
'                If nRegs = 0 Then
'                    'insertamos en la temporal
'                    SQLaux = "insert into tmpcata (codcalid, kilosnet) values (" & DBSet(NomCal(I), "N")
'                    SQLaux = SQLaux & "," & DBSet(KilCal(I), "N") & ")"
'
'                    conn.Execute SQLaux
'                Else
'                    'actualizamos la temporal
'                    SQLaux = "update tmpcata set kilosnet = kilosnet + " & DBSet(KilCal(I), "N")
'                    SQLaux = SQLaux & " where codcalid = " & DBSet(NomCal(I), "N")
'
'                    conn.Execute SQLaux
'                End If
'            End If
'        Next I
'
'        SQLaux = "select * from tmpcata order by codcalid"
'
'        Set RSaux = New ADODB.Recordset
'        RSaux.Open SQLaux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'        Sql2 = ""
'
'        While Not RSaux.EOF
'            If SeInserta Then
'                Sql2 = Sql2 & "(" & DBSet(Notaca, "N") & "," & DBSet(RS!CodVarie, "N") & ","
'                Sql2 = Sql2 & DBSet(RSaux!codcalid, "N") & "," & DBSet(RSaux!KilosNet, "N") & "),"
'            Else
'                Sql2 = "update rclasifauto_Clasif set kiloscal = kiloscal + " & DBSet(RSaux!KilosNet, "N")
'                Sql2 = Sql2 & " where numnotac = " & DBSet(Notaca, "N")
'                Sql2 = Sql2 & " and codvarie = " & DBSet(RS!CodVarie, "N")
'                Sql2 = Sql2 & " and codcalid = " & DBSet(RSaux!codcalid, "N")
'
'                conn.Execute Sql2
'            End If
'
'            RSaux.MoveNext
'        Wend
'
'        Set RSaux = Nothing
'
'
'        If SeInserta Then
'            If Sql2 <> "" Then
'                Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
'            End If
'            SQL = SQL & Sql2
'            conn.Execute SQL
'        End If
'    End If ' si la situacion es distinta de 2

    Set Rs = Nothing
    Set NomCal = Nothing
    Set KilCal = Nothing

    ProcesarFicheroCatadauCPeque�o = True
    Exit Function

eProcesarFicheroCatadauCPeque�o:
    If Err.Number <> 0 Then
        ProcesarFicheroCatadauCPeque�o = False
        MsgBox "Error en Procesar Linea " & Err.Description, vbExclamation
    End If
End Function


'************************************************************************************
'*****************PROCESO DE TRASPASO DE CALIBRADOR DE ALZIRA************************
'************************************************************************************

Public Function ProcesarDirectorioAlzira(nomDir As String, Tipo As Byte, ByRef Pb1 As ProgressBar, ByRef Label1 As Label, ByRef Label2 As Label) As Boolean
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim NumReg As Long
Dim SQL As String
Dim SQL1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim B As Boolean
Dim NomFic As String
Dim Linea As Integer
Dim Nota As String


    ProcesarDirectorioAlzira = False
    B = True
    ' Muestra los nombres en C:\ que representan directorios.
    Select Case Tipo
        Case 0, 1 ' calibrador 1 y 2 son txt
            NomFic = Dir(nomDir & "*.txt")  ' Recupera la primera entrada.
        Case 2 ' calibrador 3 (kaki) es .PTD
            NomFic = Dir(nomDir & "*.ptd")  ' Recupera la primera entrada.
    End Select
    
    If Tipo = 0 Then
    ' caso del precalibrado: cargamos todo el fichero en una tabla temporal
    
        Do While NomFic <> "" And B   ' Inicia el bucle.
           ' Ignora el directorio actual y el que lo abarca.
           If NomFic <> "." And NomFic <> ".." Then
              ' Realiza una comparaci�n a nivel de bit para asegurarse de que MiNombre es un directorio.
              If (GetAttr(nomDir & NomFic) And vbArchive) = vbArchive Then
              
                SQL = "delete from tmpcalibrador"
                conn.Execute SQL
              
                SQL = "load data local infile '" & Replace(nomDir & NomFic, "\", "/") & "' into table `tmpcalibrador` fields escaped by '\\' enclosed by '""' lines terminated by '\r\n' ( `numnota`, `fecnota`, `nomcalid`, `kilos1`, `kilos2`, `kilos3`, `kilos4`)  "
                conn.Execute SQL
                
                Label1.Caption = "Procesando Fichero: " & NomFic
                longitud = TotalRegistros("select count(*) from tmpcalibrador")
                
                Pb1.visible = True
                Pb1.Max = longitud
'                Me.Refresh
                DoEvents
                Pb1.Value = 0
                    
                If longitud <> 0 Then
                    B = ProcesarFicheroAlziraPrecalib(Pb1, Label1, Label2)
                End If
                
              End If   ' solamente si representa un directorio.
           End If
           NomFic = Dir   ' Obtiene siguiente entrada.
        Loop
    
    Else
    ' caso de escandalladora y el calibrador kaki se lee l�nea a linea del fichero de entrada
        Do While NomFic <> "" And B   ' Inicia el bucle.
           ' Ignora el directorio actual y el que lo abarca.
           If NomFic <> "." And NomFic <> ".." Then
              ' Realiza una comparaci�n a nivel de bit para asegurarse de que MiNombre es un directorio.
              If (GetAttr(nomDir & NomFic) And vbArchive) = vbArchive Then
                NF = FreeFile
                
                Open nomDir & NomFic For Input As #NF
                
                Line Input #NF, cad
                
                Label1.Caption = "Procesando Fichero: " & NomFic
                longitud = FileLen(nomDir & NomFic)
                
                Pb1.visible = True
                Pb1.Max = longitud
'                Me.Refresh
                DoEvents
                Pb1.Value = 0
                    
                If cad <> "" Then
                    Select Case Tipo
                        Case 1  'escandalladora
                            Linea = 1
                            Nota = DevuelveNota(NF, Linea)
                
                            If Nota <> "" Then
                                Close #NF
                                Open nomDir & NomFic For Input As #NF
                                Line Input #NF, cad
                        
                                B = ProcesarFicheroAlziraEscandalladora(NF, cad, Pb1, Label1, Label2, Nota)
                            End If
                        Case 2  'Kaki
                            B = ProcesarFicheroAlziraKaki(NF, cad, Pb1, Label1, Label2)
                    End Select
                End If
                
                Close #NF
                
              End If   ' solamente si representa un directorio.
           End If
        NomFic = Dir   ' Obtiene siguiente entrada.
        Loop
    End If
    
    ProcesarDirectorioAlzira = B
    
    Pb1.visible = False
    Label1.Caption = ""
    Label2.Caption = ""
                     
End Function



Private Function ProcesarFicheroAlziraEscandalladora(NF As Long, cad As String, ByRef Pb1 As ProgressBar, ByRef Label1 As Label, ByRef Label2 As Label, ByRef Nota As String) As Boolean
Dim B As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Mens As String
Dim numlinea As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim RSaux As ADODB.Recordset

Dim codsoc As String
Dim Codcam As String
Dim codpro As String
Dim codVar As String
Dim Observ As String
Dim Notaca As String
Dim Kilone As String

Dim Destri As String
Dim Podrid As String
Dim Pequen As String
Dim Muestra As String

Dim NGrupos As String
Dim Nombre1 As String
Dim Kilos As Currency


Dim i As Integer
Dim J As Integer
Dim Situacion As Byte

Dim NomCal As Dictionary
Dim KilCal As Dictionary

Dim SQLaux As String
Dim Nregs As Integer

Dim NSep As Integer

Dim SeInserta As Boolean
Dim KilosTot As Currency
Dim cantidad As Long

Dim HoraIni As String
Dim HoraFin As String

Dim FechaEnt As String
Dim UltimaLinea As Boolean
Dim NroCalidad As Integer
Dim Linea As String

    On Error GoTo eProcesarFicheroAlziraEscandalladora

    ProcesarFicheroAlziraEscandalladora = False
    
    codsoc = 0
    Codcam = 0
    codpro = 0
    codVar = 0
    Observ = ""
    Notaca = 0
    Kilone = 0
    Kilos = 0
    KilosTot = 0

    Destri = 0
    Podrid = 0
    Pequen = 0
    
    i = 0
    J = 0
    
    ' inicializamos las variables
    Set NomCal = New Dictionary
    Set KilCal = New Dictionary
            
    '[Monica] la nota no era el primer campo de las lineas
    'Notaca = RecuperaValorNew(cad, ";", 1)
    Notaca = Nota
    
    SQL = "select kilosnet, codvarie, codcampo, codsocio from rclasifica where numnotac = " & DBSet(Notaca, "N")
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Rs.EOF Then
        Observ = "NOTA NO EXISTE"
        Situacion = 2
    Else
        codVar = DBLet(Rs!Codvarie, "N")
    End If
    
    B = True
    UltimaLinea = False
    NroCalidad = 0
    While Not EOF(NF) And Not UltimaLinea
        i = i + 1
        
        Pb1.Value = Pb1.Value + Len(cad)
        Label2.Caption = "Linea " & i
        DoEvents
'        Me.Refresh
        
        NSep = NumeroSubcadenasInStr(cad, ";")
        
        If NSep = 14 Then ' estamos en una calidad
            J = J + 1
            NroCalidad = NroCalidad + 1
            
            Linea = RecuperaValorNew(cad, ";", 2)
            
            If CCur(Linea) = 1 Then
                Nombre1 = RecuperaValorNew(cad, ";", 4)
                
                ' quitamos "x.- " del nombre
                If InStr(1, Nombre1, ".- ") <> 0 Then
'
'                    Nombre1 = Mid(Nombre1, InStr(1, Nombre1, ".- ") + 3, Len(Nombre1))
                End If
                
                Kilone = RecuperaValorNew(cad, ";", 7)
                cantidad = RecuperaValorNew(cad, ";", 8)
                
                Kilos = Round2(CCur(Kilone) / 1000, 2)
                KilosTot = KilosTot + Kilos
                
                If Situacion <> 2 Then
                    ' si hay nota asociada busco los datos
                    SQL = "select codcalid from rcalidad_calibrador where codvarie = " & DBSet(Rs!Codvarie, "N")
                    SQL = SQL & " and nomcalibrador2 = " & DBSet(Trim(Nombre1), "T")
                    
                    Set RS1 = New ADODB.Recordset
                    RS1.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
                    If RS1.EOF Then
                        Observ = "NO EXIS.CAL"
                        Situacion = 1
                    Else
                        NomCal(J) = DBLet(RS1!codcalid, "N")
                        KilCal(J) = Kilos
                    End If
                    Set RS1 = Nothing
                
                End If
            Else ' se trata de destrio
                Kilone = RecuperaValorNew(cad, ";", 7)
                
                Kilos = Round2(CCur(Kilone) / 1000, 2)
                
                Destri = Destri + Kilos
            End If
        End If
        
        If NSep = 15 Then ' estamos en la ultima linea
            HoraIni = RecuperaValorNew(cad, ";", 9)
            HoraFin = RecuperaValorNew(cad, ";", 10)
            FechaEnt = RecuperaValorNew(cad, ";", 11)
            
            UltimaLinea = True
        End If
        
        Line Input #NF, cad
    Wend
    
'    Close #NF
        
' solo tenemos la suma de kilos de destrio
    If Situacion <> 2 Then
        If Destri <> 0 Then
            ' si hay kilos de destrio buscamos cual es la calidad de destrio
            SQL = "select codcalid from rcalidad where codvarie = " & DBSet(codVar, "N")
            SQL = SQL & " and tipcalid = 1 " ' calidad de destrio

            Set RS1 = New ADODB.Recordset
            RS1.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

            If RS1.EOF Then
                Observ = "NO HAY DESTRIO"
                Situacion = 5
            Else
                NomCal(J) = RS1.Fields(0).Value
                KilCal(J) = Destri

                NroCalidad = NroCalidad + 1
            End If

            Set RS1 = Nothing
        End If
    End If
        
'    If DBLet(Rs.Fields(0).Value, "N") <> KilosTot Then
'        Observ = "K.NETOS DIF."
'        Situacion = 4
'    End If

    SQL = "select count(*) from rclasifauto where numnotac = " & Notaca

    SeInserta = (TotalRegistros(SQL) = 0)

    If SeInserta Then
        If Situacion = 2 Then
            ' si no hay nota asociada no puedo meter la clasificacion
            SQL = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
            SQL = SQL & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
            SQL = SQL & "`observac`,`situacion`) values ("
            SQL = SQL & DBSet(Notaca, "N") & ","
            SQL = SQL & DBSet(0, "N") & ","
            SQL = SQL & DBSet(0, "N") & ","
            SQL = SQL & DBSet(0, "N") & ","
            SQL = SQL & DBSet(KilosTot, "N") & ","
            SQL = SQL & DBSet(Destri, "N") & ","
            SQL = SQL & DBSet(Podrid, "N") & ","
            SQL = SQL & DBSet(Pequen, "N") & ","
            SQL = SQL & DBSet(Observ, "T") & ","
            SQL = SQL & DBSet(Situacion, "N") & ")"

        Else
            ' insertamos en las tablas intermedias: rclasifauto y rclasifauto_clasif
            ' tabla: rclasifauto
            SQL = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
            SQL = SQL & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
            SQL = SQL & "`observac`,`situacion`) values ("
            SQL = SQL & DBSet(Notaca, "N") & ","
            SQL = SQL & DBSet(Rs!Codsocio, "N") & ","
            SQL = SQL & DBSet(Rs!codCampo, "N") & ","
            SQL = SQL & DBSet(Rs!Codvarie, "N") & ","
            SQL = SQL & DBSet(KilosTot, "N") & ","
            SQL = SQL & DBSet(Destri, "N") & ","
            SQL = SQL & DBSet(Podrid, "N") & ","
            SQL = SQL & DBSet(Pequen, "N") & ","
            SQL = SQL & DBSet(Observ, "T") & ","
            SQL = SQL & DBSet(Situacion, "N") & ")"
        End If
        conn.Execute SQL

        ' tabla: rclasifauto_clasif
        SQL = "insert into rclasifauto_clasif (`numnotac`,`codvarie`,`codcalid`,`kiloscal`) "
        SQL = SQL & " values "

    End If

    'solo si tenemos nota asociada metemos toda la clasificacion
    If Situacion <> 2 Then

        'borramos la tabla temporal
        SQLaux = "delete from tmpcata"
        conn.Execute SQLaux

        ' cargamos la tabla temporal
        For i = 1 To NroCalidad
            If NomCal(i) <> "" Then
                Nregs = TotalRegistros("select count(*) from tmpcata where codcalid = " & DBSet(NomCal(i), "N"))
                If Nregs = 0 Then
                    'insertamos en la temporal
                    SQLaux = "insert into tmpcata (codcalid, kilosnet) values (" & DBSet(NomCal(i), "N")
                    SQLaux = SQLaux & "," & DBSet(KilCal(i), "N") & ")"

                    conn.Execute SQLaux
                Else
                    'actualizamos la temporal
                    SQLaux = "update tmpcata set kilosnet = kilosnet + " & DBSet(KilCal(i), "N")
                    SQLaux = SQLaux & " where codcalid = " & DBSet(NomCal(i), "N")

                    conn.Execute SQLaux
                End If
            End If
        Next i

        SQLaux = "select * from tmpcata order by codcalid"

        Set RSaux = New ADODB.Recordset
        RSaux.Open SQLaux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

        Sql2 = ""
        
        If Not RSaux.EOF Then RSaux.MoveFirst
        
        While Not RSaux.EOF
            If SeInserta Then
                Sql2 = Sql2 & "(" & DBSet(Notaca, "N") & "," & DBSet(Rs!Codvarie, "N") & ","
                Sql2 = Sql2 & DBSet(RSaux!codcalid, "N") & "," & DBSet(RSaux!KilosNet, "N") & "),"
            Else
                Sql2 = "update rclasifauto_Clasif set kiloscal = kiloscal + " & DBSet(RSaux!KilosNet, "N")
                Sql2 = Sql2 & " where numnotac = " & DBSet(Notaca, "N")
                Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!Codvarie, "N")
                Sql2 = Sql2 & " and codcalid = " & DBSet(RSaux!codcalid, "N")

                conn.Execute Sql2
            End If

            RSaux.MoveNext
        Wend

        Set RSaux = Nothing

        If SeInserta Then
            If Sql2 <> "" Then
                Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
            End If
            SQL = SQL & Sql2
            conn.Execute SQL
        End If
    End If ' si la situacion es distinta de 2

    Set Rs = Nothing
    Set NomCal = Nothing
    Set KilCal = Nothing

    ProcesarFicheroAlziraEscandalladora = True
    Exit Function

eProcesarFicheroAlziraEscandalladora:
    If Err.Number <> 0 Then
        ProcesarFicheroAlziraEscandalladora = False
        MsgBox "Error en Procesar Linea " & Err.Description, vbExclamation
    End If
End Function


Private Function ProcesarFicheroAlziraPrecalib(ByRef Pb1 As ProgressBar, ByRef Label1 As Label, ByRef Label2 As Label) As Boolean
Dim B As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Mens As String
Dim numlinea As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim RSaux As ADODB.Recordset

Dim codsoc As String
Dim Codcam As String
Dim codpro As String
Dim codVar As String
Dim Observ As String
Dim Notaca As String
Dim Kilone As String

Dim Destri As String
Dim Podrid As String
Dim Pequen As String
Dim Muestra As String

Dim NGrupos As String
Dim Nombre1 As String
Dim Kilos As Currency


Dim i As Integer
Dim J As Integer
Dim Situacion As Byte

Dim NomCal As Dictionary
Dim KilCal As Dictionary

Dim SQLaux As String
Dim Nregs As Integer

Dim NSep As Integer

Dim SeInserta As Boolean
Dim KilosTot As Currency
Dim cantidad As Long

Dim HoraIni As String
Dim HoraFin As String

Dim FechaEnt As String
Dim UltimaLinea As Boolean
Dim NroCalidad As Integer
Dim Linea As String
Dim CalDestri As String
Dim CalPeque As String


    On Error GoTo eProcesarFicheroAlziraPrecalib

    ProcesarFicheroAlziraPrecalib = False
    
    codsoc = 0
    Codcam = 0
    codpro = 0
    codVar = 0
    Observ = ""
    Notaca = 0
    Kilone = 0
    Kilos = 0
    KilosTot = 0

    Destri = 0
    Pequen = 0
    
    i = 0
    
    ' inicializamos las variables
    Set NomCal = New Dictionary
    Set KilCal = New Dictionary
            
    SQL = "select * from tmpcalibrador "
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Notaca = 0
    If Not Rs.EOF Then
        Notaca = DBLet(Rs.Fields(0).Value, "N")
        
        Sql2 = "select kilosnet, codvarie, codcampo, codsocio from rclasifica where numnotac = " & DBSet(Notaca, "N")
        Set RS1 = New ADODB.Recordset
        RS1.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If RS1.EOF Then
            Observ = "NOTA NO EXISTE"
            Situacion = 2
        End If
        
        B = True
        
        While Not Rs.EOF
            i = i + 1
            
            Pb1.Value = Pb1.Value + 1
            Label2.Caption = "Linea " & i
'            Me.Refresh
            DoEvents
            
            Nombre1 = DBLet(Rs!nomcalid, "T")
            Destri = DBLet(Rs!Kilos3, "T")
            Pequen = DBLet(Rs!Kilos4, "T")
                    
            Kilone = DBLet(Rs!Kilos1, "T")
            
            Kilos = Round2(CCur(Kilone), 2)
            KilosTot = KilosTot + Kilos
                    
            If Situacion <> 2 Then
                ' si hay nota asociada busco los datos
                SQL = "select codcalid from rcalidad_calibrador where codvarie = " & DBSet(RS1!Codvarie, "N")
                SQL = SQL & " and nomcalibrador1 = " & DBSet(Nombre1, "T")
                
                Set Rs2 = New ADODB.Recordset
                Rs2.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If Rs2.EOF Then
                    Observ = "NO EXIS.CAL"
                    Situacion = 1
                Else
                    NomCal(i) = DBLet(Rs2!codcalid, "N")
                    KilCal(i) = Kilos
                End If
                Set Rs2 = Nothing
            
            End If
            
            Rs.MoveNext
        Wend
    
        SQL = "select count(*) from rclasifauto where numnotac = " & Notaca
    
        SeInserta = (TotalRegistros(SQL) = 0)
    
        If SeInserta Then
            If Situacion = 2 Then
                ' si no hay nota asociada no puedo meter la clasificacion
                SQL = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
                SQL = SQL & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
                SQL = SQL & "`observac`,`situacion`) values ("
                SQL = SQL & DBSet(Notaca, "N") & ","
                SQL = SQL & DBSet(0, "N") & ","
                SQL = SQL & DBSet(0, "N") & ","
                SQL = SQL & DBSet(0, "N") & ","
                SQL = SQL & DBSet(KilosTot, "N") & ","
                SQL = SQL & DBSet(Destri, "N") & ","
                SQL = SQL & DBSet(Podrid, "N") & ","
                SQL = SQL & DBSet(Pequen, "N") & ","
                SQL = SQL & DBSet(Observ, "T") & ","
                SQL = SQL & DBSet(Situacion, "N") & ")"
    
            Else
                ' insertamos en las tablas intermedias: rclasifauto y rclasifauto_clasif
                ' tabla: rclasifauto
                SQL = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
                SQL = SQL & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
                SQL = SQL & "`observac`,`situacion`) values ("
                SQL = SQL & DBSet(Notaca, "N") & ","
                SQL = SQL & DBSet(RS1!Codsocio, "N") & ","
                SQL = SQL & DBSet(RS1!codCampo, "N") & ","
                SQL = SQL & DBSet(RS1!Codvarie, "N") & ","
                SQL = SQL & DBSet(KilosTot, "N") & ","
                SQL = SQL & DBSet(Destri, "N") & ","
                SQL = SQL & DBSet(Podrid, "N") & ","
                SQL = SQL & DBSet(Pequen, "N") & ","
                SQL = SQL & DBSet(Observ, "T") & ","
                SQL = SQL & DBSet(Situacion, "N") & ")"
            End If
            conn.Execute SQL
    
            ' tabla: rclasifauto_clasif
            SQL = "insert into rclasifauto_clasif (`numnotac`,`codvarie`,`codcalid`,`kiloscal`) "
            SQL = SQL & " values "
    
        End If
    
        'solo si tenemos nota asociada metemos toda la clasificacion
        If Situacion <> 2 Then
    
            'borramos la tabla temporal
            SQLaux = "delete from tmpcata"
            conn.Execute SQLaux
    
            ' cargamos la tabla temporal
            For J = 1 To i
                If NomCal(J) <> "" Then
                    Nregs = TotalRegistros("select count(*) from tmpcata where codcalid = " & DBSet(NomCal(J), "N"))
                    If Nregs = 0 Then
                        'insertamos en la temporal
                        SQLaux = "insert into tmpcata (codcalid, kilosnet) values (" & DBSet(NomCal(J), "N")
                        SQLaux = SQLaux & "," & DBSet(KilCal(J), "N") & ")"
    
                        conn.Execute SQLaux
                    Else
                        'actualizamos la temporal
                        SQLaux = "update tmpcata set kilosnet = kilosnet + " & DBSet(KilCal(J), "N")
                        SQLaux = SQLaux & " where codcalid = " & DBSet(NomCal(J), "N")
    
                        conn.Execute SQLaux
                    End If
                End If
            Next J
    
            'le sumamos los kilos de destrio
            CalDestri = CalidadDestrio(RS1!Codvarie)
            If CalDestri <> "" Then
                Nregs = TotalRegistros("select count(*) from tmpcata where codcalid = " & DBSet(CalDestri, "N"))
                If Nregs = 0 Then
                    'insertamos en la temporal
                    SQLaux = "insert into tmpcata (codcalid, kilosnet) values (" & DBSet(CalDestri, "N")
                    SQLaux = SQLaux & "," & DBSet(Destri, "N") & ")"

                    conn.Execute SQLaux
                Else
                    'actualizamos la temporal
                    SQLaux = "update tmpcata set kilosnet = kilosnet + " & DBSet(Destri, "N")
                    SQLaux = SQLaux & " where codcalid = " & DBSet(CalDestri, "N")

                    conn.Execute SQLaux
                End If
            End If
            
            'le sumamos los kilos de menut
            CalPeque = CalidadMenut(RS1!Codvarie)
            If CalPeque <> "" Then
                Nregs = TotalRegistros("select count(*) from tmpcata where codcalid = " & DBSet(CalPeque, "N"))
                If Nregs = 0 Then
                    'insertamos en la temporal
                    SQLaux = "insert into tmpcata (codcalid, kilosnet) values (" & DBSet(CalPeque, "N")
                    SQLaux = SQLaux & "," & DBSet(Pequen, "N") & ")"

                    conn.Execute SQLaux
                Else
                    'actualizamos la temporal
                    SQLaux = "update tmpcata set kilosnet = kilosnet + " & DBSet(Pequen, "N")
                    SQLaux = SQLaux & " where codcalid = " & DBSet(CalPeque, "N")

                    conn.Execute SQLaux
                End If
            End If
            
            
            
            
            SQLaux = "select * from tmpcata order by codcalid"
    
            Set RSaux = New ADODB.Recordset
            RSaux.Open SQLaux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
            Sql2 = ""
    
            While Not RSaux.EOF
                If SeInserta Then
                    Sql2 = Sql2 & "(" & DBSet(Notaca, "N") & "," & DBSet(RS1!Codvarie, "N") & ","
                    Sql2 = Sql2 & DBSet(RSaux!codcalid, "N") & "," & DBSet(RSaux!KilosNet, "N") & "),"
                Else
                    Sql2 = "update rclasifauto_Clasif set kiloscal = kiloscal + " & DBSet(RSaux!KilosNet, "N")
                    Sql2 = Sql2 & " where numnotac = " & DBSet(Notaca, "N")
                    Sql2 = Sql2 & " and codvarie = " & DBSet(RS1!Codvarie, "N")
                    Sql2 = Sql2 & " and codcalid = " & DBSet(RSaux!codcalid, "N")
    
                    conn.Execute Sql2
                End If
    
                RSaux.MoveNext
            Wend
    
            Set RSaux = Nothing
    
            If SeInserta Then
                If Sql2 <> "" Then
                    Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
                End If
                SQL = SQL & Sql2
                conn.Execute SQL
            End If
        End If ' si la situacion es distinta de 2
    
        Set Rs = Nothing
        Set RS1 = Nothing
        Set NomCal = Nothing
        Set KilCal = Nothing
    
        ProcesarFicheroAlziraPrecalib = True
        Exit Function
        
    End If
    
eProcesarFicheroAlziraPrecalib:
    If Err.Number <> 0 Then
        ProcesarFicheroAlziraPrecalib = False
        MsgBox "Error en Procesar Linea " & Err.Description, vbExclamation
    End If
End Function


Private Function ProcesarFicheroAlziraKaki(NF As Long, cad As String, ByRef Pb1 As ProgressBar, ByRef Label1 As Label, ByRef Label2 As Label) As Boolean
Dim B As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Mens As String
Dim numlinea As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim RSaux As ADODB.Recordset

Dim codsoc As String
Dim Codcam As String
Dim codpro As String
Dim codVar As String
Dim Observ As String
Dim Notaca As String
Dim Kilone As String

Dim Destri As String
Dim Podrid As String
Dim Pequen As String
Dim Muestra As String

Dim NGrupos As String
Dim Nombre1 As String
Dim Kilos As Currency


Dim i As Integer
Dim J As Integer
Dim Situacion As Byte

Dim NomCal As Dictionary
Dim KilCal As Dictionary

Dim SQLaux As String
Dim Nregs As Integer

Dim NSep As Integer

Dim SeInserta As Boolean
Dim KilosTot As Currency
Dim cantidad As Long

Dim HoraIni As String
Dim HoraFin As String

Dim FechaEnt As String
Dim UltimaLinea As Boolean
Dim NroCalidad As Integer
Dim Linea As String
Dim PorcenDestrio As String

    On Error GoTo eProcesarFicheroAlziraKaki

    ProcesarFicheroAlziraKaki = False
    
    codsoc = 0
    Codcam = 0
    codpro = 0
    codVar = 0
    Observ = ""
    Notaca = 0
    Kilone = 0
    Kilos = 0
    KilosTot = 0

    Destri = 0
    Podrid = 0
    Pequen = 0
    
    i = 0
    J = 0
    
    ' inicializamos las variables
    Set NomCal = New Dictionary
    Set KilCal = New Dictionary
            
            
    ' saltamos 3 lineas
    For J = 1 To 3
        Line Input #NF, cad
        
        i = i + 1
        
        Pb1.Value = Pb1.Value + Len(cad)
        Label2.Caption = "Linea " & i
'        Me.Refresh
        DoEvents
    Next J
    
    Notaca = Mid(cad, 10, 10) ' posicion de la [10,19]
    
    SQL = "select kilosnet, codvarie, codcampo, codsocio from rclasifica where numnotac = " & DBSet(Notaca, "N")
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Rs.EOF Then
        Observ = "NOTA NO EXISTE"
        Situacion = 2
    Else
        codVar = DBLet(Rs!Codvarie, "N")
    End If
    
    ' saltamos 9 lineas
    For J = 1 To 10
        Line Input #NF, cad
    
        i = i + 1
    
        Pb1.Value = Pb1.Value + Len(cad)
        Label2.Caption = "Linea " & i
'        Me.Refresh
        DoEvents
    Next J
    
    B = True
    UltimaLinea = False
    NroCalidad = 0
    
    J = 0
    While Not EOF(NF) And Not UltimaLinea
        i = i + 1
        
        Pb1.Value = Pb1.Value + Len(cad)
        Label2.Caption = "Linea " & i
'        Me.Refresh
        DoEvents
            
        J = J + 1
        NroCalidad = NroCalidad + 1
            
        Nombre1 = Mid(cad, 6, 11)
        Kilone = Mid(cad, 17, 11)
        Kilos = CCur(Kilone)
            
        KilosTot = KilosTot + Kilos
        
        If Situacion <> 2 Then
            ' si hay nota asociada busco los datos
            SQL = "select codcalid from rcalidad_calibrador where codvarie = " & DBSet(Rs!Codvarie, "N")
            SQL = SQL & " and nomcalibrador3 = " & DBSet(Trim(Nombre1), "T")
            
            Set RS1 = New ADODB.Recordset
            RS1.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If RS1.EOF Then
                Observ = "NO EXIS.CAL"
                Situacion = 1
            Else
                NomCal(J) = DBLet(RS1!codcalid, "N")
                KilCal(J) = Kilos
'YA VEREMOS
'                ' si la calidad es de destrio sumamos los kilos a los kilos de destrio
'                If CalidadDestrio(Rs!CodVarie) = DBLet(RS1!codcalid) Then
'                    Destri = Destri + Kilos
'                End If
            End If
            Set RS1 = Nothing
        
        End If
        Line Input #NF, cad
        UltimaLinea = (Mid(cad, 17, 11) = "-----------")
    Wend
    
' solo tenemos la suma de kilos de destrio
    If Situacion <> 2 Then
        If Destri <> 0 Then
            ' si hay kilos de destrio buscamos cual es la calidad de destrio
            SQL = "select codcalid from rcalidad where codvarie = " & DBSet(codVar, "N")
            SQL = SQL & " and tipcalid = 1 " ' calidad de destrio

            Set RS1 = New ADODB.Recordset
            RS1.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

            If RS1.EOF Then
                Observ = "NO HAY DESTRIO"
                Situacion = 5
' ya veremos
'            Else
'                CalDestri = DBLet(RS1!codcalid, "N")
'                ' comprobamos qu no supera el destrio no supera el 50%
'                PorcenDestrio = Round2(Destri * 100 / KilosTot, 2)
'                If PorcenDestrio >= 50 Then
'                    Observ = "DESTRIO SUPERIOR AL 50%"
'                    Situacion = 3
'                End If
            End If

            Set RS1 = Nothing
        End If
    End If
        
'    If DBLet(Rs.Fields(0).Value, "N") <> KilosTot Then
'        Observ = "K.NETOS DIF."
'        Situacion = 4
'    End If

    SQL = "select count(*) from rclasifauto where numnotac = " & Notaca

    SeInserta = (TotalRegistros(SQL) = 0)

    If SeInserta Then
        If Situacion = 2 Then
            ' si no hay nota asociada no puedo meter la clasificacion
            SQL = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
            SQL = SQL & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
            SQL = SQL & "`observac`,`situacion`) values ("
            SQL = SQL & DBSet(Notaca, "N") & ","
            SQL = SQL & DBSet(0, "N") & ","
            SQL = SQL & DBSet(0, "N") & ","
            SQL = SQL & DBSet(0, "N") & ","
            SQL = SQL & DBSet(KilosTot, "N") & ","
            SQL = SQL & DBSet(Destri, "N") & ","
            SQL = SQL & DBSet(Podrid, "N") & ","
            SQL = SQL & DBSet(Pequen, "N") & ","
            SQL = SQL & DBSet(Observ, "T") & ","
            SQL = SQL & DBSet(Situacion, "N") & ")"

        Else
            ' insertamos en las tablas intermedias: rclasifauto y rclasifauto_clasif
            ' tabla: rclasifauto
            SQL = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
            SQL = SQL & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
            SQL = SQL & "`observac`,`situacion`) values ("
            SQL = SQL & DBSet(Notaca, "N") & ","
            SQL = SQL & DBSet(Rs!Codsocio, "N") & ","
            SQL = SQL & DBSet(Rs!codCampo, "N") & ","
            SQL = SQL & DBSet(Rs!Codvarie, "N") & ","
            SQL = SQL & DBSet(KilosTot, "N") & ","
            SQL = SQL & DBSet(Destri, "N") & ","
            SQL = SQL & DBSet(Podrid, "N") & ","
            SQL = SQL & DBSet(Pequen, "N") & ","
            SQL = SQL & DBSet(Observ, "T") & ","
            SQL = SQL & DBSet(Situacion, "N") & ")"
        End If
        conn.Execute SQL

        ' tabla: rclasifauto_clasif
        SQL = "insert into rclasifauto_clasif (`numnotac`,`codvarie`,`codcalid`,`kiloscal`) "
        SQL = SQL & " values "

    End If

    'solo si tenemos nota asociada metemos toda la clasificacion
    If Situacion <> 2 Then

        'borramos la tabla temporal
        SQLaux = "delete from tmpcata"
        conn.Execute SQLaux

        ' cargamos la tabla temporal
        For i = 1 To NroCalidad
            If NomCal(i) <> "" Then
                Nregs = TotalRegistros("select count(*) from tmpcata where codcalid = " & DBSet(NomCal(i), "N"))
                If Nregs = 0 Then
                    'insertamos en la temporal
                    SQLaux = "insert into tmpcata (codcalid, kilosnet) values (" & DBSet(NomCal(i), "N")
                    SQLaux = SQLaux & "," & DBSet(KilCal(i), "N") & ")"

                    conn.Execute SQLaux
                Else
                    'actualizamos la temporal
                    SQLaux = "update tmpcata set kilosnet = kilosnet + " & DBSet(KilCal(i), "N")
                    SQLaux = SQLaux & " where codcalid = " & DBSet(NomCal(i), "N")

                    conn.Execute SQLaux
                End If
            End If
        Next i

        SQLaux = "select * from tmpcata order by codcalid"

        Set RSaux = New ADODB.Recordset
        RSaux.Open SQLaux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

        Sql2 = ""
        
        If Not RSaux.EOF Then RSaux.MoveFirst
        
        While Not RSaux.EOF
            If SeInserta Then
                Sql2 = Sql2 & "(" & DBSet(Notaca, "N") & "," & DBSet(Rs!Codvarie, "N") & ","
                Sql2 = Sql2 & DBSet(RSaux!codcalid, "N") & "," & DBSet(RSaux!KilosNet, "N") & "),"
            Else
                Sql2 = "update rclasifauto_Clasif set kiloscal = kiloscal + " & DBSet(RSaux!KilosNet, "N")
                Sql2 = Sql2 & " where numnotac = " & DBSet(Notaca, "N")
                Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!Codvarie, "N")
                Sql2 = Sql2 & " and codcalid = " & DBSet(RSaux!codcalid, "N")

                conn.Execute Sql2
            End If

            RSaux.MoveNext
        Wend

        Set RSaux = Nothing

        If SeInserta Then
            If Sql2 <> "" Then
                Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
            End If
            SQL = SQL & Sql2
            conn.Execute SQL
        End If
'ya veremos
'        If Destri <> 0 Then
'            Sql = "delete from rclasifauto_clasif where numnotac = " & DBSet(Notaca, "N")
'            Sql = Sql & " and codvarie = " & DBSet(Rs!CodVarie, "N")
'            Sql = Sql & " and codcalid = " & CalDestri
'
'            conn.Execute Sql
'        End If
    End If ' si la situacion es distinta de 2

    Set Rs = Nothing
    Set NomCal = Nothing
    Set KilCal = Nothing

    ProcesarFicheroAlziraKaki = True
    Exit Function

eProcesarFicheroAlziraKaki:
    If Err.Number <> 0 Then
        ProcesarFicheroAlziraKaki = False
        MsgBox "Error en Procesar Linea " & Err.Description, vbExclamation
    End If
End Function

'***************VALSUR y parte antigua de Catadau

Public Function ProcesarFichero(nomFich As String, TipoCal As Byte, ByRef Pb1 As ProgressBar, Label1 As Label, Label2 As Label) As Boolean
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim NumReg As Long
Dim SQL As String
Dim SQL1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim B As Boolean
Dim NomFic As String

    On Error GoTo eProcesarFichero


    ProcesarFichero = False
    NF = FreeFile
    
    Open nomFich For Input As #NF
    
    Line Input #NF, cad
    i = 0

    Label1.Caption = "Procesando Fichero: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Pb1.Max = longitud
'    Me.Refresh
    DoEvents
    Pb1.Value = 0
        
        
    B = True
    While Not EOF(NF)
        i = i + 1
        
        Pb1.Value = Pb1.Value + Len(cad)
        Label2.Caption = "Linea " & i
'        Me.Refresh
        DoEvents
        
        If vParamAplic.Cooperativa = 1 Then ' si es valsur
            B = ProcesarLineaValsur(cad, TipoCal)
        Else ' si es catadau
            B = ProcesarLineaCatadau(NF, cad, TipoCal, Pb1, Label1, Label2)
            If TipoCal = 0 Then i = i + 6
        End If
        
        If B = False Then
            ProcesarFichero = False
            Exit Function
        End If
        
        If Not EOF(NF) Then Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" And B Then
        If vParamAplic.Cooperativa = 1 Then ' si es valsur
            B = ProcesarLineaValsur(cad, TipoCal)
'        Else
'            b = ProcesarLineaCatadau(NF, Cad, Combo1(6).ListIndex)
        End If
        If B = False Then
            ProcesarFichero = False
            Exit Function
        End If
    End If
    
    ProcesarFichero = B
    
    Pb1.visible = False
    Label1.Caption = ""
    Label2.Caption = ""

eProcesarFichero:
    If Err.Number <> 0 Then
        MuestraError Err.Description
    End If


End Function



Private Function ProcesarLineaCatadau(NF As Long, cad As String, Calibr As Byte, ByRef Pb1 As ProgressBar, ByRef Label1 As Label, ByRef Label2 As Label) As Boolean
Dim B As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Mens As String
Dim numlinea As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim RSaux As ADODB.Recordset

Dim codsoc As String
Dim Codcam As String
Dim codpro As String
Dim codVar As String
Dim Observ As String
Dim Notaca As String
Dim Kilone As String

Dim Destri As String
Dim Podrid As String
Dim Pequen As String
Dim Muestra As String

Dim NGrupos As String
Dim Nombre1 As String
Dim Kilos As String


Dim i As Integer
Dim Situacion As Byte

Dim NomCal As Dictionary
Dim KilCal As Dictionary

Dim SQLaux As String
Dim Nregs As Integer

Dim SeInserta As Boolean


    On Error GoTo eProcesarLineaCatadau

    ProcesarLineaCatadau = False
    
    codsoc = 0
    Codcam = 0
    codpro = 0
    codVar = 0
    Observ = ""
    Notaca = 0
    Kilone = 0

    Destri = 0
    Podrid = 0
    Pequen = 0
    
    ' inicializamos las variables
    Set NomCal = New Dictionary
    Set KilCal = New Dictionary
    
    Select Case Calibr
        Case 0  ' CALIBRADOR GRANDE
            'primera linea: cabecera
            If cad <> "" Then
                Notaca = RecuperaValorNew(cad, ",", 5)
                
                Kilone = RecuperaValorNew(cad, ",", 6)
                Destri = RecuperaValorNew(cad, ",", 11)
                Podrid = RecuperaValorNew(cad, ",", 9)
                Pequen = RecuperaValorNew(cad, ",", 10)
        
                SQL = "select kilosnet, codvarie, codcampo, codsocio from rclasifica where numnotac = " & DBSet(Notaca, "N")
                Set Rs = New ADODB.Recordset
                Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
                If Rs.EOF Then
                    Observ = "NOTA NO EXISTE"
                    Situacion = 2
                Else
                    If DBLet(Rs.Fields(0).Value, "N") <> Kilone Then
                        Observ = "K.NETOS DIF."
                        Situacion = 4
                    End If
                End If
                ' salto tipo b
                Line Input #NF, cad
                
                Pb1.Value = Pb1.Value + Len(cad)
                Label2.Caption = "Linea " & i
'                Me.Refresh
                DoEvents
                
                ' salto tipo c
                Line Input #NF, cad
                
                Pb1.Value = Pb1.Value + Len(cad)
                Label2.Caption = "Linea " & i
'                Me.Refresh
                DoEvents
                
                NGrupos = RecuperaValorNew(cad, ",", 4)
                
                'salto tipo d
                Line Input #NF, cad
                
                Pb1.Value = Pb1.Value + Len(cad)
                Label2.Caption = "Linea " & i
'                Me.Refresh
                DoEvents
                
                cad = cad & ","
                For i = 0 To NGrupos - 1
                    Nombre1 = RecuperaValorNew(cad, ",", 4 + i)
                
                
                    If Situacion <> 2 Then
                        ' si hay nota asociada busco los datos
                        
                        SQL = "select codcalid from rcalidad_calibrador where codvarie = " & DBSet(Rs!Codvarie, "N")
                        SQL = SQL & " and nomcalibrador1 = " & DBSet(Nombre1, "T")
                        
                        Set RS1 = New ADODB.Recordset
                        RS1.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        
                        If RS1.EOF Then
                            Observ = "NO EXIS.CAL"
                            Situacion = 1
                        Else
                            NomCal(i) = DBLet(RS1!codcalid, "N")
                        End If
                        Set RS1 = Nothing
                    End If
                
                Next i
            
                ' salto tipo e
                Line Input #NF, cad
                
                Pb1.Value = Pb1.Value + Len(cad)
                Label2.Caption = "Linea " & i
'                Me.Refresh
                DoEvents
            
                ' salto tipo f: pesos de la calidad
                Line Input #NF, cad
                Pb1.Value = Pb1.Value + Len(cad)
                Label2.Caption = "Linea " & i
'                Me.Refresh
                DoEvents
                
                cad = cad & ","
                For i = 0 To NGrupos - 1
                    KilCal(i) = RecuperaValorNew(cad, ",", i + 4)
                Next i
               
                SQL = "select count(*) from rclasifauto where numnotac = " & Notaca
                
                SeInserta = (TotalRegistros(SQL) = 0)
                
                If SeInserta Then
                    If Situacion = 2 Then
                        ' si no hay nota asociada no puedo meter la clasificacion
                        SQL = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
                        SQL = SQL & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
                        SQL = SQL & "`observac`,`situacion`) values ("
                        SQL = SQL & DBSet(Notaca, "N") & ","
                        SQL = SQL & DBSet(0, "N") & ","
                        SQL = SQL & DBSet(0, "N") & ","
                        SQL = SQL & DBSet(0, "N") & ","
                        SQL = SQL & DBSet(Kilone, "N") & ","
                        SQL = SQL & DBSet(Destri, "N") & ","
                        SQL = SQL & DBSet(Podrid, "N") & ","
                        SQL = SQL & DBSet(Pequen, "N") & ","
                        SQL = SQL & DBSet(Observ, "T") & ","
                        SQL = SQL & DBSet(Situacion, "N") & ")"
                    
                    Else
                        ' insertamos en las tablas intermedias: rclasifauto y rclasifauto_clasif
                        ' tabla: rclasifauto
                        SQL = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
                        SQL = SQL & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
                        SQL = SQL & "`observac`,`situacion`) values ("
                        SQL = SQL & DBSet(Notaca, "N") & ","
                        SQL = SQL & DBSet(Rs!Codsocio, "N") & ","
                        SQL = SQL & DBSet(Rs!codCampo, "N") & ","
                        SQL = SQL & DBSet(Rs!Codvarie, "N") & ","
                        SQL = SQL & DBSet(Kilone, "N") & ","
                        SQL = SQL & DBSet(Destri, "N") & ","
                        SQL = SQL & DBSet(Podrid, "N") & ","
                        SQL = SQL & DBSet(Pequen, "N") & ","
                        SQL = SQL & DBSet(Observ, "T") & ","
                        SQL = SQL & DBSet(Situacion, "N") & ")"
                    End If
                    conn.Execute SQL
                
                    ' tabla: rclasifauto_clasif
                    SQL = "insert into rclasifauto_clasif (`numnotac`,`codvarie`,`codcalid`,`kiloscal`) "
                    SQL = SQL & " values "
                    
                End If
                
                'solo si tenemos nota asociada metemos toda la clasificacion
                If Situacion <> 2 Then
                    
                    
                    'borramos la tabla temporal
                    SQLaux = "delete from tmpcata"
                    conn.Execute SQLaux
                    
                    ' cargamos la tabla temporal
                    For i = 0 To NGrupos - 1
                        If NomCal(i) <> "" Then
                            Nregs = TotalRegistros("select count(*) from tmpcata where codcalid = " & DBSet(NomCal(i), "N"))
                            If Nregs = 0 Then
                                'insertamos en la temporal
                                SQLaux = "insert into tmpcata (codcalid, kilosnet) values (" & DBSet(NomCal(i), "N")
                                SQLaux = SQLaux & "," & KilCal(i) & ")"
                                
                                conn.Execute SQLaux
                            Else
                                'actualizamos la temporal
                                SQLaux = "update tmpcata set kilosnet = kilosnet + " & KilCal(i)
                                SQLaux = SQLaux & " where codcalid = " & DBSet(NomCal(i), "N")
                                
                                conn.Execute SQLaux
                            End If
                        End If
                    Next i
                    
                    SQLaux = "select * from tmpcata order by codcalid"
                    
                    Set RSaux = New ADODB.Recordset
                    RSaux.Open SQLaux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
                    Sql2 = ""
                    
                    While Not RSaux.EOF
                        If SeInserta Then
                            Sql2 = Sql2 & "(" & DBSet(Notaca, "N") & "," & DBSet(Rs!Codvarie, "N") & ","
                            Sql2 = Sql2 & DBSet(RSaux!codcalid, "N") & "," & DBSet(RSaux!KilosNet, "N") & "),"
                        Else
                            Sql2 = "update rclasifauto_Clasif set kiloscal = kiloscal + " & DBSet(RSaux!KilosNet, "N")
                            Sql2 = Sql2 & " where numnotac = " & DBSet(Notaca, "N")
                            Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!Codvarie, "N")
                            Sql2 = Sql2 & " and codcalid = " & DBSet(RSaux!codcalid, "N")
                            
                            conn.Execute Sql2
                        End If
                        
                        RSaux.MoveNext
                    Wend
                    
                    Set RSaux = Nothing
                    
                    
                    If SeInserta Then
                        If Sql2 <> "" Then
                            Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
                        End If
                        SQL = SQL & Sql2
                        conn.Execute SQL
                    End If
                End If ' si la situacion es distinta de 2
                
                
' 18-05-2009
'                Sql2 = ""
'                For I = 0 To NomCal.Count - 1
'                    Sql2 = "(" & DBSet(Notaca, "N") & "," & DBSet(rs!CodVarie, "N") & ","
'                    Sql2 = Sql2 & DBSet(NomCal(I), "N") & "," & DBSet(KilCal(I), "N") & "),"
'                Next I
'                If Sql2 <> "" Then
'                    Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
'                End If
'                SQL = SQL & Sql2
'                conn.Execute SQL
                
                ' salto tipo g
                Line Input #NF, cad
                
                Set Rs = Nothing
                Set NomCal = Nothing
                Set KilCal = Nothing
            
            Else
                Exit Function
            End If
            
        Case 1 ' CALIBRADOR PEQUE�O
            ' saltamos 5 lineas mas
            For i = 1 To 5
                Line Input #NF, cad
            Next i
            Muestra = cad
            ' saltamos para kilosnetos
            Line Input #NF, cad
            Kilone = cad
            ' saltamos para podrido
            Line Input #NF, cad
            Podrid = cad
            ' saltamos para destrio
            Line Input #NF, cad
            Destri = cad
            
            Kilos = CCur(ImporteSinFormato(Kilone)) - CCur(ImporteSinFormato(Podrid)) - CCur(ImporteSinFormato(Destri))
            
            ' saltamos para nota de campo
            Line Input #NF, cad
            
            
'****************falsta esto
'            Notaca = Mid(NomFic, 1, 7)
            
            SQL = "select codsocio, codcampo, codvarie, kilosnet from rclasifica"
            SQL = SQL & " where numnotac = " & DBSet(Notaca, "N")
            
            Set RS1 = New ADODB.Recordset
            RS1.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If RS1.EOF Then
                Observ = "NOTA NO EXI."
                Situacion = 2
            Else
                If DBLet(RS1!KilosNet, "N") < Kilos Then
                    Observ = "K.NETOS DIF."
                    Situacion = 4
                End If
            End If
            ' ++++++++++++++++++++estoy aqui linea 360 de agre1104
        
        
    End Select
    ProcesarLineaCatadau = True
    Exit Function
    

eProcesarLineaCatadau:
    If Err.Number <> 0 Then
        ProcesarLineaCatadau = False
        MsgBox "Error en Procesar Linea " & Err.Description, vbExclamation
    End If
End Function

'
' Proceso de traspaso para VALSUR
'
Private Function ProcesarLineaValsur(cad As String, Calibrador As Byte) As Boolean
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String

Dim NumNota As String
Dim KilosNet As String
Dim KilosDes As String
Dim KilosPod As String
Dim KilosTot As String

Dim NomCal As Dictionary
Dim KilCal As Dictionary

Dim Situacion As Byte

Dim CodCal As Integer
Dim Observac As String
Dim KilosNota As Long

Dim i As Integer

Dim c_Cantidad As Currency
Dim c_Importe As Currency
Dim c_Precio As Currency
Dim Mens As String
Dim numlinea As Long

    On Error GoTo eProcesarLineaValsur

    ProcesarLineaValsur = True
    
    ' inicializamos las variables
    Set NomCal = New Dictionary
    Set KilCal = New Dictionary
    
    NumNota = 0
    KilosNet = 0
    KilosDes = 0
    KilosPod = 0
    KilosTot = 0
    Observac = ""
    Situacion = 0
    
    NumNota = RecuperaValor(cad, 3)
    KilosNet = RecuperaValor(cad, 4)
    KilosDes = RecuperaValor(cad, 17)
    KilosPod = RecuperaValor(cad, 18)
    KilosTot = RecuperaValor(cad, 19)
    
    For i = 1 To 12
        NomCal(i) = RecuperaValor(cad, i + 4)
        KilCal(i) = RecuperaValor(cad, i + 19)
    Next i
    
    SQL = "select codsocio, codcampo, codvarie from rclasifica where numnotac = " & DBSet(NumNota, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Rs.EOF Then
        Observac = "NOTA NO EXISTE"
        Situacion = 2
    
        'insertamos la cabecera de la clasificacion
        SQL = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`,"
        SQL = SQL & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,`observac`,`situacion` ) values ("
        SQL = SQL & DBSet(NumNota, "N") & ","
        SQL = SQL & ValorNulo & ","
        SQL = SQL & 0 & ","
        SQL = SQL & ValorNulo & ","
        SQL = SQL & DBSet(KilosTot, "N") & ","
        SQL = SQL & DBSet(KilosDes, "N") & ","
        SQL = SQL & DBSet(KilosPod, "N") & ","
        SQL = SQL & DBSet(KilosNet, "N") & ","
        SQL = SQL & DBSet(Observac, "T") & ","
        SQL = SQL & DBSet(Situacion, "N") & ")"
        
        conn.Execute SQL
    
        ' no metemos la clasificacion pq no se corresponde con ninguna nota
    Else
        ' insertamos las calidades si existen
        For i = 1 To 12
            If NomCal(i) <> "" And KilCal(i) <> 0 Then
                Sql2 = "select codcalid from rcalidad where codvarie = " & DBSet(Rs!Codvarie, "N")
                Select Case Calibrador
                    Case 0 ' calibrador 1
                        Sql2 = Sql2 & " and nomcalibrador1 = " & DBSet(NomCal(i), "T")
                    Case 1 ' calibrador 2
                        Sql2 = Sql2 & " and nomcalibrador2 = " & DBSet(NomCal(i), "T")
                End Select
                Set Rs2 = New ADODB.Recordset
                Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not Rs2.EOF Then
                    CodCal = DBLet(Rs2!codcalid, "N")
                    Situacion = 0
                Else
'                    CodCal = 999
'                    Observac = "NO EXIS.CAL."
'                    Situacion = 1
                    MsgBox "No existe la calidad " & NomCal(i) & ".Revise.", vbExclamation

                    ProcesarLineaValsur = False
                    
                    Set Rs = Nothing
                    Set Rs2 = Nothing
                    
                    Set NomCal = Nothing
                    Set KilCal = Nothing
                    
                    Exit Function
                End If
                
                Set Rs2 = Nothing
                
                Sql3 = "insert into rclasifauto_clasif(numnotac,codvarie,codcalid,kiloscal) values ("
                Sql3 = Sql3 & DBSet(NumNota, "N") & ","
                Sql3 = Sql3 & DBSet(Rs!Codvarie, "N") & ","
                Sql3 = Sql3 & DBSet(CodCal, "N") & ","
                Sql3 = Sql3 & DBSet(KilCal(i), "N") & ")"
                
                conn.Execute Sql3
            End If
        Next i
    
        'insertamos la cabecera de la clasificacion
        SQL = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`,"
        SQL = SQL & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,`observac`,`situacion`) values ("
        SQL = SQL & DBSet(NumNota, "N") & ","
        SQL = SQL & DBSet(Rs!Codsocio, "N") & ","
        SQL = SQL & DBSet(Rs!codCampo, "N") & ","
        SQL = SQL & DBSet(Rs!Codvarie, "N") & ","
        SQL = SQL & DBSet(KilosTot, "N") & ","
        SQL = SQL & DBSet(KilosDes, "N") & ","
        SQL = SQL & DBSet(KilosPod, "N") & ","
        SQL = SQL & DBSet(KilosNet, "N") & ","
        SQL = SQL & DBSet(Observac, "T") & ","
        SQL = SQL & DBSet(Situacion, "N") & ")"
        
        conn.Execute SQL
    
    End If
    
    Set Rs = Nothing
    
    Set NomCal = Nothing
    Set KilCal = Nothing
    
eProcesarLineaValsur:
    If Err.Number <> 0 Then
        ProcesarLineaValsur = False
        MsgBox "Error en Procesar Linea " & Err.Description, vbExclamation
    End If
End Function

Private Function ProcesarFicheroCatadau(nomDir As String, ByRef Pb1 As ProgressBar, ByRef Label1 As Label, ByRef Label2 As Label) As Boolean
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim NumReg As Long
Dim SQL As String
Dim SQL1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim B As Boolean
Dim NomFic As String

    ProcesarFicheroCatadau = False
    
    ' Muestra los nombres en C:\ que representan directorios.
    NomFic = Dir(nomDir, vbDirectory)   ' Recupera la primera entrada.
    Do While NomFic <> "" And B   ' Inicia el bucle.
       ' Ignora el directorio actual y el que lo abarca.
       If NomFic <> "." And NomFic <> ".." Then
          ' Realiza una comparaci�n a nivel de bit para asegurarse de que MiNombre es un directorio.
          If (GetAttr(nomDir & NomFic) And vbDirectory) = vbDirectory Then
            NF = FreeFile
            
            Open nomDir & NomFic For Input As #NF
            
            Line Input #NF, cad
            i = 0
            Dir
            Label1.Caption = "Procesando Fichero: " & NomFic
            longitud = FileLen(NomFic)
            
            Pb1.visible = True
            Pb1.Max = longitud
'            Me.Refresh
            DoEvents
            Pb1.Value = 0
                
                
            B = True
            While Not EOF(NF)
                i = i + 1
                
                Pb1.Value = Pb1.Value + Len(cad)
                Label2.Caption = "Linea " & i
'                Me.Refresh
                DoEvents
                
                B = ProcesarLineaCatadau(NF, cad, 1, Pb1, Label1, Label2) '1=calibrador peque�o
                
                If B = False Then
                    ProcesarFicheroCatadau = False
                    Exit Function
                End If
                
                Line Input #NF, cad
            Wend
            Close #NF
            
            If cad <> "" And B Then
                B = ProcesarLineaCatadau(NF, cad, 1, Pb1, Label1, Label2) '1=calibrador peque�o
                If B = False Then
                    ProcesarFicheroCatadau = False
                    Exit Function
                End If
            End If
            
          End If   ' solamente si representa un directorio.
       End If
       NomFic = Dir   ' Obtiene siguiente entrada.
    Loop
    
    
    ProcesarFicheroCatadau = B
    
    Pb1.visible = False
    Label1.Caption = ""
    Label2.Caption = ""
                     
End Function


Private Function DevuelveNota(NF As Long, ByRef Linea As Integer) As String
Dim cad As String
Dim NSep As Integer

    DevuelveNota = ""
    
    While Not EOF(NF)
        Line Input #NF, cad
        
        Linea = Linea + 1
        
        NSep = NumeroSubcadenasInStr(cad, ";")
        
        If NSep = 15 Then ' estamos sacamos el nro de nota
            DevuelveNota = RecuperaValorNew(cad, ";", 5)
        End If
    Wend

End Function


'************************************************************************************
'*****************PROCESO DE TRASPASO DE CALIBRADOR DE CASTELDUC ********************
'************************************************************************************

Public Function ProcesarDirectorioCastelduc(nomDir As String, Tipo As Byte, ByRef Pb1 As ProgressBar, ByRef Label1 As Label, ByRef Label2 As Label, Optional NotaD As String, Optional NotaH As String) As Boolean
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim NumReg As Long
Dim SQL As String
Dim SQL1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim B As Boolean
Dim NomFic As String
Dim Linea As Integer
Dim Nota As String


    ProcesarDirectorioCastelduc = False
    B = True
    ' Muestra los nombres en C:\ que representan directorios.
    Select Case Tipo
        Case 0, 1 ' calibrador 1 y 2 son txt
            NomFic = Dir(nomDir & "*.txt")  ' Recupera la primera entrada.
        Case 2 ' calibrador de rugat
            NomFic = Dir(nomDir & "crugat1.txt")
    End Select
    
    If Tipo = 0 Or Tipo = 1 Then
    ' caso del precalibrado: cargamos todo el fichero en una tabla temporal
'--
'        Do While NomFic <> "" And b   ' Inicia el bucle.
'           ' Ignora el directorio actual y el que lo abarca.
'           If NomFic <> "." And NomFic <> ".." Then
'              ' Realiza una comparaci�n a nivel de bit para asegurarse de que MiNombre es un directorio.
''              If (GetAttr(nomDir & NomFic) And vbArchive) = vbArchive Then
'
'                SQL = "delete from tmpcalibrador"
'                conn.Execute SQL
'
'                SQL = "load data local infile '" & Replace(nomDir & NomFic, "\", "/") & "' into table `tmpcalibrador` fields escaped by '\\' enclosed by '""' lines terminated by '\r\n' ( `numnota`, `fecnota`, `nomcalid`, `kilos1`, `kilos2`, `kilos3`, `kilos4`)  "
'                conn.Execute SQL
'
'                Label1.Caption = "Procesando Fichero: " & NomFic
'                longitud = TotalRegistros("select count(*) from tmpcalibrador")
'
'                Pb1.visible = True
'                Pb1.Max = longitud
'                'Me.Refresh
'                DoEvents
'                Pb1.Value = 0
'
'                If longitud <> 0 Then
'                    b = ProcesarFicheroAlziraPrecalib(Pb1, Label1, Label2)
'                End If
'
''              End If   ' solamente si representa un directorio.
'           End If
'           NomFic = Dir   ' Obtiene siguiente entrada.
'        Loop
'++
        SQL = "select distinct numnotac from tmpcalibradorcast where codusu = " & vUsu.Codigo
        SQL = SQL & " and numnotac >= " & DBSet(NotaD, "N") & " and numnotac <= " & DBSet(NotaH, "N")
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs.EOF
        
            SQL = "delete from tmpcalibrador"
            conn.Execute SQL
        
            SQL = "insert into tmpcalibrador (numnota, nomcalid, kilos1) "
            SQL = SQL & " select numnotac, nomcalib, sum(kilos) from tmpcalibradorcast where codusu = " & vUsu.Codigo
            SQL = SQL & " and numnotac = " & DBSet(Rs!NumNotac, "N") & " and numcalid > -1 "
            SQL = SQL & " group by 1,2"
            
            conn.Execute SQL
            
            SQL = "update tmpcalibrador set kilos3 = "
            SQL = SQL & " (select sum(kilos) from tmpcalibradorcast where numnotac = " & DBSet(Rs!NumNotac, "N") & " and codusu = " & vUsu.Codigo & " and numcalib = -1)"
            SQL = SQL & " where numnota = " & DBSet(Rs!NumNotac, "N")
            
            conn.Execute SQL
            
            SQL = "update tmpcalibrador set kilos1 = replace(kilos1,'.',','), kilos3 = replace(kilos3,'.',',') "
            
            conn.Execute SQL
            
        
            Label1.Caption = "Procesando Nota: " & Rs!NumNotac
            longitud = TotalRegistros("select count(*) from tmpcalibrador")

            Pb1.visible = True
            Pb1.Max = longitud
'            Me.Refresh
            DoEvents
            Pb1.Value = 0

            If longitud <> 0 Then
                B = ProcesarFicheroAlziraPrecalib(Pb1, Label1, Label2)
            End If
        
            Rs.MoveNext
        Wend
        Set Rs = Nothing
    Else
        ' castello de rugat para castelduc
        ' solo hay un fichero que le pasan, luego hay que procesarlo
        B = ProcesarFicheroCastelducRugat(NomFic, Pb1, Label1, Label2, 0)
        
    
    End If
    
    ProcesarDirectorioCastelduc = B
    
    Pb1.visible = False
    Label1.Caption = ""
    Label2.Caption = ""
                     
End Function


'************************************************************************************
'*****************PROCESO DE TRASPASO DE CALIBRADOR DE Frutas Inma ********************
'************************************************************************************

Public Function ProcesarDirectorioFrutasInma(nomDir As String, Tipo As Byte, ByRef Pb1 As ProgressBar, ByRef Label1 As Label, ByRef Label2 As Label, Optional NotaD As String, Optional NotaH As String) As Boolean
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim NumReg As Long
Dim SQL As String
Dim SQL1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim B As Boolean
Dim NomFic As String
Dim Linea As Integer
Dim Nota As String


    ProcesarDirectorioFrutasInma = False
    B = True
    ' Muestra los nombres en C:\ que representan directorios.
    Select Case Tipo
        Case 0, 1 ' calibrador 1 y 2 son txt
            NomFic = Dir(nomDir & "*.txt")  ' Recupera la primera entrada.
        Case 2 ' calibrador de rugat
            NomFic = Dir(nomDir & "crugat1.txt")
    End Select
    
    If Tipo = 0 Or Tipo = 1 Then

        SQL = "select distinct numnotac from tmpcalibradorcast where codusu = " & vUsu.Codigo
        SQL = SQL & " and numnotac >= " & DBSet(NotaD, "N") & " and numnotac <= " & DBSet(NotaH, "N")
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs.EOF
        
            SQL = "delete from tmpcalibrador"
            conn.Execute SQL
        
            SQL = "insert into tmpcalibrador (numnota, nomcalid, kilos1) "
            SQL = SQL & " select numnotac, concat(numcalid,'|',numcolor,'|',numcalib,'|') nomcalid, sum(kilos) from tmpcalibradorcast where codusu = " & vUsu.Codigo
            SQL = SQL & " and numnotac = " & DBSet(Rs!NumNotac, "N") & " and numcalib > -1 "
            SQL = SQL & " group by 1,2"
            
            conn.Execute SQL
            
            SQL = "update tmpcalibrador set kilos3 = "
            SQL = SQL & " (select sum(kilos) from tmpcalibradorcast where numnotac = " & DBSet(Rs!NumNotac, "N") & " and codusu = " & vUsu.Codigo & " and numcalib = -1)"
            SQL = SQL & " where numnota = " & DBSet(Rs!NumNotac, "N")
            
            conn.Execute SQL
            
            SQL = "update tmpcalibrador set kilos1 = replace(kilos1,'.',','), kilos3 = replace(kilos3,'.',',') "
            
            conn.Execute SQL
            
        
            Label1.Caption = "Procesando Nota: " & Rs!NumNotac
            longitud = TotalRegistros("select count(*) from tmpcalibrador")

            Pb1.visible = True
            Pb1.Max = longitud
'            Me.Refresh
            DoEvents
            Pb1.Value = 0

            If longitud <> 0 Then
                B = ProcesarFicheroAlziraPrecalib(Pb1, Label1, Label2)
            End If
        
            Rs.MoveNext
        Wend
        Set Rs = Nothing
        
    
    End If
    
    ProcesarDirectorioFrutasInma = B
    
    Pb1.visible = False
    Label1.Caption = ""
    Label2.Caption = ""
                     
End Function





Private Function ProcesarFicheroCastelducRugat(nomFich As String, ByRef Pb1 As ProgressBar, ByRef Label1 As Label, ByRef Label2 As Label, ByRef Nota As String) As Boolean
Dim B As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Mens As String
Dim numlinea As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim RSaux As ADODB.Recordset

Dim codsoc As String
Dim Codcam As String
Dim codpro As String
Dim codVar As String
Dim Observ As String
Dim Notaca As String
Dim Kilone As String

Dim Destri As String
Dim Podrid As String
Dim Pequen As String
Dim Muestra As String

Dim NGrupos As String
Dim Nombre1 As String



Dim i As Integer
Dim Situacion As Byte

Dim NomCal As Dictionary
Dim KilCal As Dictionary

Dim SQLaux As String
Dim Nregs As Integer

Dim NSep As Integer

Dim SeInserta As Boolean
Dim KilosTot As Currency
Dim cantidad As Long
Dim Kilos As Currency

Dim HoraIni As String
Dim HoraFin As String

Dim FechaEnt As String
Dim UltimaLinea As Boolean
Dim NroCalidad As Integer

Dim Porcen As String
Dim KilosMuestreo As String
Dim HayReg As Boolean
Dim NF As Integer
Dim cad As String
Dim Cad1 As String
Dim longitud As Long



    On Error GoTo eProcesarFicheroCastelducRugat

    ProcesarFicheroCastelducRugat = False
    
    codsoc = 0
    Codcam = 0
    codpro = 0
    codVar = 0
    Observ = ""
    Notaca = 0
    Kilone = 0
    KilosTot = 0

    Destri = 0
    Podrid = 0
    Pequen = 0
    
    i = 0
    
    
    NF = FreeFile
    
    Open nomFich For Input As #NF
    
    Line Input #NF, Cad1
    i = 0

    Label1.Caption = "Procesando Fichero: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Pb1.Max = longitud
'    Me.Refresh
    DoEvents
    Pb1.Value = 0
        
    B = True
    While Not EOF(NF) Or Len(Cad1) <> 0
            ' cada linea es una nota
            
            i = i + 1
            
            cad = Cad1
            
            Pb1.Value = Pb1.Value + Len(cad)
            Label2.Caption = "Linea " & i
'            Me.Refresh
            DoEvents
        
            ' inicializamos las variables
            Set NomCal = New Dictionary
            Set KilCal = New Dictionary
            
'            cad = Replace(cad, Asc(9), Asc(32))
            
            Notaca = ""
            Notaca = Mid(cad, 1, PrimerBlanco(cad)) ' numero de nota
            cad = Trim(Mid(cad, PrimerBlanco(cad) + 1, Len(cad)))
        
            SQL = "select kilosnet, codvarie, codcampo, codsocio from rclasifica where numnotac = " & DBSet(Notaca, "N")
            Set Rs = New ADODB.Recordset
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
            Situacion = 0
            If Rs.EOF Then
                Observ = "NOTA NO EXISTE"
                Situacion = 2
            End If
        
            B = True
            
            'saltamos 3 y sacamos los kilos netos
            cad = Mid(cad, PrimerBlanco(cad) + 1, Len(cad))
            cad = Mid(cad, PrimerBlanco(cad) + 1, Len(cad))
            cad = Mid(cad, PrimerBlanco(cad) + 1, Len(cad))
            Kilone = Mid(cad, 1, 9)
            KilosTot = ImporteSinFormato(Kilone)
            
            ' saltamos 9
            cad = Mid(cad, PrimerBlanco(cad) + 1, Len(cad))
            cad = Mid(cad, PrimerBlanco(cad) + 1, Len(cad))
            cad = Mid(cad, PrimerBlanco(cad) + 1, Len(cad))
            cad = Mid(cad, PrimerBlanco(cad) + 1, Len(cad))
            cad = Mid(cad, PrimerBlanco(cad) + 1, Len(cad))
            cad = Mid(cad, PrimerBlanco(cad) + 1, Len(cad))
            cad = Mid(cad, PrimerBlanco(cad) + 1, Len(cad))
            cad = Mid(cad, PrimerBlanco(cad) + 1, Len(cad))
            cad = Mid(cad, PrimerBlanco(cad) + 1, Len(cad))
            cad = Mid(cad, PrimerBlanco(cad) + 1, Len(cad))
            cad = Mid(cad, PrimerBlanco(cad) + 1, Len(cad))
            
            NomCal(1) = 1
            KilCal(1) = Mid(cad, 1, PrimerBlanco(cad))
            cad = Mid(cad, PrimerBlanco(cad) + 1, Len(cad))
            
            NomCal(2) = 2
            KilCal(2) = Mid(cad, 1, PrimerBlanco(cad))
            cad = Mid(cad, PrimerBlanco(cad) + 1, Len(cad))
            
            NomCal(3) = 3
            KilCal(3) = Mid(cad, 1, PrimerBlanco(cad))
            cad = Mid(cad, PrimerBlanco(cad) + 1, Len(cad))
            
            NomCal(4) = 4
            KilCal(4) = Mid(cad, 1, PrimerBlanco(cad))
            cad = Mid(cad, PrimerBlanco(cad) + 1, Len(cad))
            
            NomCal(5) = 5
            KilCal(5) = Mid(cad, 1, PrimerBlanco(cad))
            cad = Mid(cad, PrimerBlanco(cad) + 1, Len(cad))
            
            NomCal(6) = 6
            KilCal(6) = Mid(cad, 1, PrimerBlanco(cad))
            cad = Mid(cad, PrimerBlanco(cad) + 1, Len(cad))
            
            NomCal(7) = 7
            KilCal(7) = Mid(cad, 1, PrimerBlanco(cad))
            cad = Mid(cad, PrimerBlanco(cad) + 1, Len(cad))
            
            NomCal(8) = 8
            KilCal(8) = Mid(cad, 1, PrimerBlanco(cad))
            cad = Mid(cad, PrimerBlanco(cad) + 1, Len(cad))
            
            NomCal(9) = 9
            KilCal(9) = Mid(cad, 1, PrimerBlanco(cad))
            cad = Mid(cad, PrimerBlanco(cad) + 1, Len(cad))
            
            cad = Mid(cad, PrimerBlanco(cad) + 1, Len(cad))
            cad = Mid(cad, PrimerBlanco(cad) + 1, Len(cad))
            
            NomCal(10) = 10
            KilCal(10) = Mid(cad, 1, PrimerBlanco(cad))
            cad = Mid(cad, PrimerBlanco(cad) + 1, Len(cad))
                
            If Situacion <> 2 Then
                If DBLet(Rs.Fields(0).Value, "N") <> Int(KilosTot) Then
                    Observ = "K.NETOS DIF."
                    Situacion = 4
                End If
            End If
        
            SQL = "select count(*) from rclasifauto where numnotac = " & Notaca
        
            SeInserta = (TotalRegistros(SQL) = 0)
        
            If SeInserta Then
                If Situacion = 2 Then
                    ' si no hay nota asociada no puedo meter la clasificacion
                    SQL = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
                    SQL = SQL & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
                    SQL = SQL & "`observac`,`situacion`) values ("
                    SQL = SQL & DBSet(Notaca, "N") & ","
                    SQL = SQL & DBSet(0, "N") & ","
                    SQL = SQL & DBSet(0, "N") & ","
                    SQL = SQL & DBSet(0, "N") & ","
                    SQL = SQL & DBSet(KilosTot, "N") & ","
                    SQL = SQL & DBSet(0, "N") & ","
                    SQL = SQL & DBSet(0, "N") & ","
                    SQL = SQL & DBSet(0, "N") & ","
                    SQL = SQL & DBSet(Observ, "T") & ","
                    SQL = SQL & DBSet(Situacion, "N") & ")"
        
                Else
                    ' insertamos en las tablas intermedias: rclasifauto y rclasifauto_clasif
                    ' tabla: rclasifauto
                    SQL = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
                    SQL = SQL & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
                    SQL = SQL & "`observac`,`situacion`) values ("
                    SQL = SQL & DBSet(Notaca, "N") & ","
                    SQL = SQL & DBSet(Rs!Codsocio, "N") & ","
                    SQL = SQL & DBSet(Rs!codCampo, "N") & ","
                    SQL = SQL & DBSet(Rs!Codvarie, "N") & ","
                    SQL = SQL & DBSet(Round2(KilosTot, 0), "N") & ","
                    SQL = SQL & DBSet(0, "N") & ","
                    SQL = SQL & DBSet(0, "N") & ","
                    SQL = SQL & DBSet(0, "N") & ","
                    SQL = SQL & DBSet(Observ, "T") & ","
                    SQL = SQL & DBSet(Situacion, "N") & ")"
                End If
                conn.Execute SQL
        
                ' tabla: rclasifauto_clasif
                SQL = "insert into rclasifauto_clasif (`numnotac`,`codvarie`,`codcalid`,`kiloscal`) "
                SQL = SQL & " values "
        
            End If
        
        
            'solo si tenemos nota asociada metemos toda la clasificacion
            If Situacion <> 2 Then
        
                'borramos la tabla temporal
                SQLaux = "delete from tmpcata"
                conn.Execute SQLaux
        
                ' cargamos la tabla temporal
                For i = 1 To 10
                    If NomCal(i) <> "" Then
                        Nregs = TotalRegistros("select count(*) from tmpcata where codcalid = " & DBSet(NomCal(i), "N"))
                        If Nregs = 0 Then
                            'insertamos en la temporal
                            SQLaux = "insert into tmpcata (codcalid, kilosnet) values (" & DBSet(NomCal(i), "N")
                            SQLaux = SQLaux & "," & DBSet(ImporteSinFormato(KilCal(i)), "N") & ")"
        
                            conn.Execute SQLaux
                        Else
                            'actualizamos la temporal
                            SQLaux = "update tmpcata set kilosnet = kilosnet + " & DBSet(ImporteSinFormato(KilCal(i)), "N")
                            SQLaux = SQLaux & " where codcalid = " & DBSet(NomCal(i), "N")
        
                            conn.Execute SQLaux
                        End If
                    End If
                Next i
        
                SQLaux = "select * from tmpcata order by codcalid"
        
                Set RSaux = New ADODB.Recordset
                RSaux.Open SQLaux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
                Sql2 = ""
        
                HayReg = False
        
                While Not RSaux.EOF
                    HayReg = True
                    If SeInserta Then
                        Sql2 = Sql2 & "(" & DBSet(Notaca, "N") & "," & DBSet(Rs!Codvarie, "N") & ","
                        Sql2 = Sql2 & DBSet(RSaux!codcalid, "N") & "," & DBSet(RSaux!KilosNet, "N") & "),"
                    Else
                        Sql2 = "update rclasifauto_Clasif set kiloscal = kiloscal + " & DBSet(RSaux!KilosNet, "N")
                        Sql2 = Sql2 & " where numnotac = " & DBSet(Notaca, "N")
                        Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!Codvarie, "N")
                        Sql2 = Sql2 & " and codcalid = " & DBSet(RSaux!codcalid, "N")
        
                        conn.Execute Sql2
                    End If
        
                    RSaux.MoveNext
                Wend
        
                Set RSaux = Nothing
        
                If SeInserta And HayReg Then
                    If Sql2 <> "" Then
                        Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
                    End If
                    SQL = SQL & Sql2
                    conn.Execute SQL
                End If
            End If ' si la situacion es distinta de 2
        
            Set Rs = Nothing
            Set NomCal = Nothing
            Set KilCal = Nothing
        
            Cad1 = ""
        
            If Not EOF(NF) Then Line Input #NF, Cad1
            
    Wend
    
    Close #NF
    
    ProcesarFicheroCastelducRugat = True
    Exit Function

eProcesarFicheroCastelducRugat:
    If Err.Number <> 0 Then
        ProcesarFicheroCastelducRugat = False
        MsgBox "Error en Procesar Linea " & Err.Description, vbExclamation
    End If
End Function


Private Function PrimerBlanco(cadena As String) As Long
Dim J As Long

    PrimerBlanco = 0
    J = 1
    While Asc(Mid(cadena, J, 1)) <> 9 And J <= Len(cadena)
        J = J + 1
    Wend
    PrimerBlanco = J
    
End Function


'************************************************************************************
'*****************PROCESO DE TRASPASO DE CALIBRADOR DE PICASSENT*********************
'************************************************************************************

Public Function ProcesarDirectorioPicassent(nomDir As String, Tipo As Byte, ByRef Pb1 As ProgressBar, ByRef Label1 As Label, ByRef Label2 As Label) As Boolean
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim NumReg As Long
Dim SQL As String
Dim SQL1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim B As Boolean
Dim NomFic As String
Dim Linea As Integer
Dim Nota As String


    ProcesarDirectorioPicassent = False
    B = True
    ' Muestra los nombres en C:\ que representan directorios.
    Select Case Tipo
        Case 0 ' calibrador 1 y 2 son txt
            NomFic = Dir(nomDir & "*.tag")  ' Recupera la primera entrada.
    End Select
    
        Do While NomFic <> "" And B   ' Inicia el bucle.
           ' Ignora el directorio actual y el que lo abarca.
           If NomFic <> "." And NomFic <> ".." Then
              ' Realiza una comparaci�n a nivel de bit para asegurarse de que MiNombre es un directorio.
              If (GetAttr(nomDir & NomFic) And vbArchive) = vbArchive Then
                NF = FreeFile
                
                Open nomDir & NomFic For Input As #NF
                
                Line Input #NF, cad
                
                Label1.Caption = "Procesando Fichero: " & NomFic
                longitud = FileLen(nomDir & NomFic)
                
                Pb1.visible = True
                Pb1.Max = longitud
'                Me.Refresh
                DoEvents
                Pb1.Value = 0
                    
                If cad <> "" Then
                    Select Case Tipo
                        Case 0
                            B = ProcesarFicheroPicassent(NF, cad, Pb1, Label1, Label2)
                    End Select
                End If
                
                Close #NF
                
              End If   ' solamente si representa un directorio.
           End If
        NomFic = Dir   ' Obtiene siguiente entrada.
        Loop
    
    
    ProcesarDirectorioPicassent = B
    
    Pb1.visible = False
    Label1.Caption = ""
    Label2.Caption = ""
                     
End Function


Private Function ProcesarFicheroPicassent(NF As Long, cad As String, ByRef Pb1 As ProgressBar, ByRef Label1 As Label, ByRef Label2 As Label) As Boolean
Dim B As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Mens As String
Dim numlinea As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim RSaux As ADODB.Recordset

Dim codsoc As String
Dim Codcam As String
Dim codpro As String
Dim codVar As String
Dim Observ As String
Dim Kilone As String
Dim NroCam As String

Dim Destri As String
Dim Podrid As String
Dim Pequen As String
Dim Muestra As String

Dim NGrupos As String
Dim Nombre1 As String
Dim Kilos As Currency


Dim i As Integer
Dim J As Integer
Dim Situacion As Byte

Dim NomCal As Dictionary
Dim KilCal As Dictionary

Dim SQLaux As String
Dim Nregs As Integer

Dim NSep As Integer

Dim SeInserta As Boolean
Dim KilosTot As Currency
Dim cantidad As Long

Dim HoraIni As String
Dim HoraFin As String

Dim FechaEnt As String
Dim UltimaLinea As Boolean
Dim NroCalidad As Integer
Dim Linea As String
Dim PorcenDestrio As String

Dim Inicio As String
Dim Fin As Boolean
Dim vCadena As String
Dim vClasi As String
Dim vFecha As String
Dim Ordinal As Long

    On Error GoTo eProcesarFicheroPicassent

    ProcesarFicheroPicassent = False
    
    codsoc = 0
    Codcam = 0
    codpro = 0
    codVar = 0
    NroCam = 0
    Observ = ""
    Kilone = 0
    Kilos = 0
    KilosTot = 0

    Destri = 0
    Podrid = 0
    Pequen = 0
    
    vClasi = 0
    vFecha = "01/01/1900"
    
    i = 0
    J = 0
    
    ' inicializamos las variables
    Set NomCal = New Dictionary
    Set KilCal = New Dictionary
    
    Line Input #NF, cad
                
                
    Situacion = 0
                
    Fin = False
    While Not EOF(NF) And Not Fin
        Select Case Inicio
            Case "101"
                If Situacion = 0 Then
                    vCadena = RecuperaValorNew(cad, ",", 3)
                    
                    If InStr(1, vCadena, "/") <> 0 Then
                        codsoc = RecuperaValorNew(vCadena, "/", 1)
                        codsoc = Mid(codsoc, 2, Len(codsoc))
                        Codcam = Mid(vCadena, InStr(1, vCadena, "/") + 1, Len(vCadena)) ', RecuperaValorNew(vCadena, "/", 2), 1, 4)
                        If InStr(1, Codcam, "-") <> 0 Then
                            Codcam = RecuperaValorNew(Codcam, "-", 1)
                        Else
                            If Len(Codcam) <> 0 Then Codcam = Mid(Codcam, 1, Len(Codcam) - 1)
                        End If
                        If InStr(1, vCadena, "-") <> 0 Then
                            vClasi = Mid(vCadena, InStr(1, vCadena, "-") + 1, 1)  ' RecuperaValorNew(vCadena, "-", 2)
                        End If
                        If CInt(codsoc) <> 999 Then
                            SQL = "select count(*) from rsocios where codsocio = " & DBSet(codsoc, "N")
                            If TotalRegistros(SQL) = 0 Then
                                Observ = "NO EXIS.SOC"
                                Situacion = 2
                            End If
                        End If
                    End If
                    If vClasi = 0 Then vClasi = 1
                End If
                
            Case "103"
                codVar = RecuperaValorNew(cad, ",", 2)
                If Situacion = 0 Then
                    SQL = "select count(*) from variedades where codvarie = " & DBSet(codVar, "N")
                    If TotalRegistros(SQL) = 0 Then
                        Observ = "NO EXIS.VAR"
                        Situacion = 3
                    Else
                        If CLng(ComprobarCero(Codcam)) <> 9999 Then
                            SQL = "select count(*) from rcampos where codsocio= " & DBSet(codsoc, "N")
                            SQL = SQL & " and codvarie = " & DBSet(codVar, "N")
                            SQL = SQL & " and codcampo = " & DBSet(Codcam, "N")
                            If TotalRegistros(SQL) = 0 Then
                                Observ = "NO EXIS.CPO"
                                Situacion = 4
                            End If
                        End If
                    End If
                End If
                
            Case "104"
                If Situacion = 0 Then
                    vFecha = Mid(cad, InStr(1, cad, ",") + 2, 10) 'RecuperaValorNew(cad, ",", 1)
                    If Not EsFechaOK(vFecha) Then
                        Observ = "FECHA INCOR"
                        Situacion = 5
                    End If
                End If
                
            Case "400"
                If Situacion = 0 Then
                    cad = cad & ","
                    NGrupos = RecuperaValorNew(cad, ",", 2)
                    For i = 1 To NGrupos
                        Nombre1 = RecuperaValorNew(cad, ",", i + 2)
                    
                        SQL = "select codcalid from rcalidad_calibrador where codvarie = " & DBSet(codVar, "N")
                        SQL = SQL & " and nomcalibrador1 = " & DBSet(Nombre1, "T")
                        
                        NomCal(i) = DevuelveValor(SQL)
                        If NomCal(i) = 0 Then
                            Observ = "NO EXIS.CAL"
                            Situacion = 1
                        End If
                    Next i
                End If
                
            Case "451"
                If Situacion = 0 Then
                    cad = cad & ","
                    KilosTot = 0
                    For i = 1 To NGrupos
                        Nombre1 = RecuperaValorNew(cad, ",", i + 1)
                        KilCal(i) = Round2(CCur(TransformaPuntosComas(Nombre1)) / 1000, 0) 'Nombre1
                        
                        KilosTot = KilosTot + Round2(CCur(TransformaPuntosComas(Nombre1)) / 1000, 0)
                    Next i
                End If
                
            Case "999"
                Fin = True
        End Select
        
        If Not Fin Then
            Line Input #NF, cad
            Inicio = Mid(cad, 1, 3)
        End If
    Wend
    
    '[Monica] en el fichero me viene el codcampo y he de mostrar el nro de campo
    SQL = "select nrocampo from rcampos where codsocio= " & DBSet(codsoc, "N")
    SQL = SQL & " and codvarie = " & DBSet(codVar, "N")
    SQL = SQL & " and codcampo = " & DBSet(Codcam, "N")
    
    NroCam = DevuelveValor(SQL)
    
    If vClasi = 0 Then vClasi = 1
    
    SQL = "select max(ordinal) + 1 from rclasifauto where codsocio = " & DBSet(codsoc, "N")
    SQL = SQL & " and codcampo = " & DBSet(NroCam, "N")
    SQL = SQL & " and codvarie = " & DBSet(codVar, "N")
    SQL = SQL & " and numnotac = " & DBSet(vClasi, "N")
    SQL = SQL & " and fechacla = " & DBSet(vFecha, "F")
    

    Ordinal = DevuelveValor(SQL)
    If Ordinal = 0 Then Ordinal = 1
'    SeInserta = (TotalRegistros(Sql) = 0)

'    If SeInserta Then
        SQL = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
        SQL = SQL & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
        SQL = SQL & "`observac`,`situacion`,`fechacla`,`ordinal`  ) values ("
        SQL = SQL & DBSet(vClasi, "N") & ","
        SQL = SQL & DBSet(codsoc, "N") & ","
        SQL = SQL & DBSet(NroCam, "N") & ","
        SQL = SQL & DBSet(codVar, "N") & ","
        SQL = SQL & DBSet(KilosTot, "N") & ","
        SQL = SQL & DBSet(0, "N") & ","
        SQL = SQL & DBSet(0, "N") & ","
        SQL = SQL & DBSet(0, "N") & ","
        SQL = SQL & DBSet(Observ, "T") & ","
        SQL = SQL & DBSet(Situacion, "N") & ","
        SQL = SQL & DBSet(vFecha, "F") & ","
        SQL = SQL & DBSet(Ordinal, "N") & ")"
        
        conn.Execute SQL
        
 '   End If
        
    If Situacion = 0 Then
        ' tabla: rclasifauto_clasif
        SQL = "insert into rclasifauto_clasif (`numnotac`,`codvarie`,`codcalid`,`kiloscal`,`codcampo`,`codsocio`,`fechacla`,`ordinal`) "
        SQL = SQL & " values "
    
        'borramos la tabla temporal
        SQLaux = "delete from tmpcata"
        conn.Execute SQLaux
    
        ' cargamos la tabla temporal
        For i = 1 To NGrupos
            If NomCal(i) <> "" Then
                Nregs = TotalRegistros("select count(*) from tmpcata where codcalid = " & DBSet(NomCal(i), "N"))
                If Nregs = 0 Then
                    'insertamos en la temporal
                    SQLaux = "insert into tmpcata (codcalid, kilosnet) values (" & DBSet(NomCal(i), "N")
                    SQLaux = SQLaux & "," & DBSet(KilCal(i), "N") & ")"
    
                    conn.Execute SQLaux
                Else
                    'actualizamos la temporal
                    SQLaux = "update tmpcata set kilosnet = kilosnet + " & DBSet(KilCal(i), "N")
                    SQLaux = SQLaux & " where codcalid = " & DBSet(NomCal(i), "N")
    
                    conn.Execute SQLaux
                End If
            End If
        Next i
    
        SQLaux = "select * from tmpcata order by codcalid"
    
        Set RSaux = New ADODB.Recordset
        RSaux.Open SQLaux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        Sql2 = ""
        
        If Not RSaux.EOF Then RSaux.MoveFirst
        
        While Not RSaux.EOF
'            If SeInserta Then
                Sql2 = Sql2 & "(" & DBSet(vClasi, "N") & "," & DBSet(codVar, "N") & ","
                Sql2 = Sql2 & DBSet(RSaux!codcalid, "N") & "," & DBSet(RSaux!KilosNet, "N") & ","
                Sql2 = Sql2 & DBSet(NroCam, "N") & "," & DBSet(codsoc, "N") & ","
                Sql2 = Sql2 & DBSet(vFecha, "F") & "," & DBSet(Ordinal, "N") & "),"
'            Else
'                Sql2 = "update rclasifauto_Clasif set kiloscal = kiloscal + " & DBSet(RSaux!KilosNet, "N")
'                Sql2 = Sql2 & " where numnotac = " & DBSet(vClasi, "N")
'                Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!CodVarie, "N")
'                Sql2 = Sql2 & " and codcalid = " & DBSet(RSaux!codcalid, "N")
'                Sql2 = Sql2 & " and codcampo = " & DBSet(Codcam, "N")
'                Sql2 = Sql2 & " and codsocio = " & DBSet(CodSoc, "N")
'                Sql2 = Sql2 & " and fechacla = " & DBSet(vFecha, "F")
'
'                conn.Execute Sql2
'            End If
    
            RSaux.MoveNext
        Wend
    
        Set RSaux = Nothing
    
'        If SeInserta Then
            If Sql2 <> "" Then
                Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
                SQL = SQL & Sql2
                conn.Execute SQL
            End If
'        End If
    End If
    
    Set Rs = Nothing
    Set NomCal = Nothing
    Set KilCal = Nothing

    ProcesarFicheroPicassent = True
    Exit Function

eProcesarFicheroPicassent:
    If Err.Number <> 0 Then
        ProcesarFicheroPicassent = False
        MsgBox "Error en Procesar Linea " & Err.Description, vbExclamation
    End If
End Function


Public Function ProcesarDirectorioCOOPIC(nomDir As String, Tipo As Byte, ByRef Pb1 As ProgressBar, ByRef Label1 As Label, ByRef Label2 As Label) As Boolean
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim NumReg As Long
Dim SQL As String
Dim SQL1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim B As Boolean
Dim NomFic As String
Dim Linea As Integer
Dim Nota As String


    ProcesarDirectorioCOOPIC = False
    B = True
    ' Muestra los nombres en C:\ que representan directorios.
    
'    NomFic = Dir(nomDir & "*.txt")  ' Recupera la primera entrada.
    
    NomFic = nomDir
    
    If NomFic <> "" Then
        B = ProcesarFicheroCOOPIC(NomFic, Pb1, Label1, Label2)
    End If
    
    ProcesarDirectorioCOOPIC = B
    
    Pb1.visible = False
    Label1.Caption = ""
    Label2.Caption = ""
                     
End Function


Private Function ProcesarFicheroCOOPIC(nomFich As String, ByRef Pb1 As ProgressBar, ByRef Label1 As Label, ByRef Label2 As Label) As Boolean
Dim B As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Mens As String
Dim numlinea As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim RSaux As ADODB.Recordset

Dim codsoc As String
Dim Codcam As String
Dim codpro As String
Dim codVar As String
Dim Observ As String
Dim Notaca As String
Dim Kilone As String

Dim Destri As String
Dim Podrid As String
Dim Pequen As String
Dim Muestra As String

Dim NGrupos As String
Dim Nombre1 As String



Dim i As Integer
Dim Situacion As Byte

Dim NomCal As Dictionary
Dim KilCal As Dictionary

Dim SQLaux As String
Dim Nregs As Integer

Dim NSep As Integer

Dim SeInserta As Boolean
Dim KilosTot As Currency
Dim cantidad As Long
Dim Kilos As Currency

Dim HoraIni As String
Dim HoraFin As String

Dim FechaEnt As String
Dim UltimaLinea As Boolean
Dim NroCalidad As Integer

Dim Porcen As String
Dim KilosMuestreo As String
Dim HayReg As Boolean
Dim Cad1 As String
Dim longitud As Long
Dim Destrio As String
Dim Varie As String
Dim Fecha As String
Dim NF As Integer
Dim cad As String
Dim CodCal As Integer
Dim Ordinal As Integer

    On Error GoTo eProcesarFicheroCOOPIC

    ProcesarFicheroCOOPIC = False
    
    codsoc = 0
    Codcam = 0
    codpro = 0
    codVar = 0
    Observ = ""
    Notaca = 0
    Kilone = 0
    KilosTot = 0

    Destri = 0
    Podrid = 0
    Pequen = 0
    
    
    NF = FreeFile
    
    Open nomFich For Input As #NF

    Dim Col As Dictionary
    Dim J As Integer
    Dim Lineas As Integer
    
    Line Input #NF, Cad1
    i = 0


    'lee una linea
    'separa enta

    Label1.Caption = "Procesando Fichero: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Pb1.Max = longitud
'    Me.Refresh
    DoEvents
    Pb1.Value = 0
        
        
        
    B = True
    While Not EOF(NF) Or Len(Cad1) <> 0
    
    
        Set Col = New Dictionary
    
        If Len(Cad1) > 125 Then
            'Viene
            ' metemos en la coleccion
            Lineas = Len(Cad1) / 125
            For J = 1 To Lineas
                Col(J) = Mid(Cad1, (125 * (J - 1)) + 1, 125)
            Next J
        Else
            Col(1) = Cad1
        End If
    
        For J = 1 To Lineas
    
            cad = Col(J)
            
            Pb1.Value = Pb1.Value + Len(cad)
            Label2.Caption = "Linea " & J
'            Me.Refresh
            DoEvents
        
            ' inicializamos las variables
            Set NomCal = New Dictionary
            Set KilCal = New Dictionary
            
    '            cad = Replace(cad, Asc(9), Asc(32))
            
            Notaca = ""
            Notaca = Mid(cad, 1, 9) ' numero de nota
        
            SQL = "select kilosnet, codvarie, codcampo, codsocio, fechaent from rclasifica where numnotac = " & DBSet(Notaca, "N")
            Set Rs = New ADODB.Recordset
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
            Situacion = 0
            If Rs.EOF Then
                Observ = "NOTA NO EXISTE"
                Situacion = 2
            End If
        
            B = True
            
            
            NomCal(1) = 1
            KilCal(1) = Mid(cad, 14, 9)
            KilCal(1) = Round2(KilCal(1) / 1000, 2)
            
            NomCal(2) = 2
            KilCal(2) = Mid(cad, 23, 9)
            KilCal(2) = Round2(KilCal(2) / 1000, 2)
            
            NomCal(3) = 3
            KilCal(3) = Mid(cad, 32, 9)
            KilCal(3) = Round2(KilCal(3) / 1000, 2)
            
            NomCal(4) = 4
            KilCal(4) = Mid(cad, 41, 9)
            KilCal(4) = Round2(KilCal(4) / 1000, 2)
            
            
            NomCal(5) = 5
            KilCal(5) = Mid(cad, 50, 9)
            KilCal(5) = Round2(KilCal(5) / 1000, 2)
            
            NomCal(6) = 6
            KilCal(6) = Mid(cad, 59, 9)
            KilCal(6) = Round2(KilCal(6) / 1000, 2)
            
            
            NomCal(7) = 7
            KilCal(7) = Mid(cad, 68, 9)
            KilCal(7) = Round2(KilCal(7) / 1000, 2)
            
            NomCal(8) = 8
            KilCal(8) = Mid(cad, 77, 9)
            KilCal(8) = Round2(KilCal(8) / 1000, 2)
            
            
            NomCal(9) = 9
            KilCal(9) = Mid(cad, 86, 9)
            KilCal(9) = Round2(KilCal(9) / 1000, 2)
            
            
            NomCal(10) = 10
            KilCal(10) = Mid(cad, 95, 9)
            KilCal(10) = Round2(KilCal(10) / 1000, 2)
            
                
            Destrio = Round(Mid(cad, 104, 9) / 1000, 2)
                
            Varie = Mid(cad, 113, 2)
            Fecha = Replace(Mid(cad, 115, 10), ".", "/")
                
                
    '        If Situacion <> 2 Then
    '            If DBLet(Rs.Fields(1).Value, "N") <> Int(Varie) Then
    '                Observ = "VARIEDAD DIF."
    '                Situacion = 8
    '            End If
    '        End If
            
        
            SQL = "select count(*) from rclasifauto where numnotac = " & Notaca
        
            SeInserta = (TotalRegistros(SQL) = 0)
        
            If Situacion <> 2 Then
                SQL = "select max(ordinal) + 1 from rclasifauto where codsocio = " & DBSet(Rs!Codsocio, "N")
                SQL = SQL & " and codcampo = " & DBSet(Rs!codCampo, "N")
                SQL = SQL & " and codvarie = " & DBSet(Rs!Codvarie, "N")
                SQL = SQL & " and numnotac = " & DBSet(Notaca, "N")
                SQL = SQL & " and fechacla = " & DBSet(Rs!FechaEnt, "F")
                
            
                Ordinal = DevuelveValor(SQL)
                If Ordinal = 0 Then Ordinal = 1
            Else
                Ordinal = 1
            End If
            
        
        
'            If SeInserta Then
                If Situacion = 2 Then
                    ' si no hay nota asociada no puedo meter la clasificacion
                    SQL = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`,  "
                    SQL = SQL & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
                    SQL = SQL & "`observac`,`situacion`,ordinal) values ("
                    SQL = SQL & DBSet(Notaca, "N") & ","
                    SQL = SQL & DBSet(0, "N") & ","
                    SQL = SQL & DBSet(0, "N") & ","
                    SQL = SQL & DBSet(0, "N") & ","
                    SQL = SQL & DBSet(0, "N") & ","
                    SQL = SQL & DBSet(0, "N") & ","
                    SQL = SQL & DBSet(0, "N") & ","
                    SQL = SQL & DBSet(0, "N") & ","
                    SQL = SQL & DBSet(Observ, "T") & ","
                    SQL = SQL & DBSet(Situacion, "N") & "," & DBSet(Ordinal, "N") & ")"
        
                Else
                    ' insertamos en las tablas intermedias: rclasifauto y rclasifauto_clasif
                    ' tabla: rclasifauto
                    SQL = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
                    SQL = SQL & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
                    SQL = SQL & "`observac`,`situacion`,fechacla, ordinal) values ("
                    SQL = SQL & DBSet(Notaca, "N") & ","
                    SQL = SQL & DBSet(Rs!Codsocio, "N") & ","
                    SQL = SQL & DBSet(Rs!codCampo, "N") & ","
                    SQL = SQL & DBSet(Rs!Codvarie, "N") & ","
                    SQL = SQL & DBSet(Rs!KilosNet, "N") & ","
                    SQL = SQL & DBSet(0, "N") & ","
                    SQL = SQL & DBSet(0, "N") & ","
                    SQL = SQL & DBSet(0, "N") & ","
                    SQL = SQL & DBSet(Observ, "T") & ","
                    SQL = SQL & DBSet(Situacion, "N") & ","
                    SQL = SQL & DBSet(Rs!FechaEnt, "F") & "," & DBSet(Ordinal, "N") & ")"
                End If
                conn.Execute SQL
        
                ' tabla: rclasifauto_clasif
                SQL = "insert into rclasifauto_clasif (`numnotac`,`codvarie`,`codcalid`,`kiloscal`,codcampo, codsocio, fechacla,ordinal) "
                SQL = SQL & " values "
        
'            End If
        
        
            'solo si tenemos nota asociada metemos toda la clasificacion
            If Situacion <> 2 Then
                'borramos la tabla temporal
                SQLaux = "delete from tmpcata"
                conn.Execute SQLaux
        
                ' cargamos la tabla temporal
                For i = 1 To 10
                    If KilCal(i) <> 0 Then
                        CodCal = DevuelveValor("select codcalid from rcalidad where codvarie = " & DBSet(Rs!Codvarie, "N") & " and nomcalibrador1 = " & DBSet(NomCal(i), "T"))
                        
                        If CInt(CodCal) = 0 Then
                            Observ = "NO EXIS.CAL"
                            Situacion = 1
                        Else
                    
                            Nregs = TotalRegistros("select count(*) from tmpcata where codcalid = " & DBSet(CodCal, "N"))
                            If Nregs = 0 Then
                                'insertamos en la temporal
                                SQLaux = "insert into tmpcata (codcalid, kilosnet) values (" & DBSet(CodCal, "N")
                                SQLaux = SQLaux & "," & DBSet(ImporteSinFormato(KilCal(i)), "N") & ")"
            
                                conn.Execute SQLaux
                            Else
                                'actualizamos la temporal
                                SQLaux = "update tmpcata set kilosnet = kilosnet + " & DBSet(ImporteSinFormato(KilCal(i)), "N")
                                SQLaux = SQLaux & " where codcalid = " & DBSet(CodCal, "N")
            
                                conn.Execute SQLaux
                            End If
                            
                        End If
                    End If
                Next i
                
                If Destrio <> 0 Then
                    CodCal = DevuelveValor("select codcalid from rcalidad where codvarie = " & DBSet(Rs!Codvarie, "N") & " and tipcalid = 1")
                    If CInt(CodCal) <> 0 Then
                        SQLaux = "insert into tmpcata (codcalid, kilosnet) values (" & DBSet(CodCal, "N")
                        SQLaux = SQLaux & "," & DBSet(Destrio, "N") & ")"
                        
                        conn.Execute SQLaux
                    Else
                        Observ = "NO EXIS.CAL"
                        Situacion = 1
                    End If
                End If
        
                SQLaux = "select * from tmpcata order by codcalid"
        
                Set RSaux = New ADODB.Recordset
                RSaux.Open SQLaux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
                Sql2 = ""
        
                HayReg = False
        
                While Not RSaux.EOF
                    HayReg = True
'[Monica] como en Picassent, se inserta siempre
'                    If SeInserta Then
                        Sql2 = Sql2 & "(" & DBSet(Notaca, "N") & "," & DBSet(Rs!Codvarie, "N") & ","
                        Sql2 = Sql2 & DBSet(RSaux!codcalid, "N") & "," & DBSet(RSaux!KilosNet, "N") & ","
                        Sql2 = Sql2 & DBSet(Rs!codCampo, "N") & "," & DBSet(Rs!Codsocio, "N") & ","
                        Sql2 = Sql2 & DBSet(Rs!FechaEnt, "F") & "," & DBSet(Ordinal, "N") & "),"
'                    Else
'                        Sql2 = "update rclasifauto_clasif set kiloscal = kiloscal + " & DBSet(RSaux!KilosNet, "N")
'                        Sql2 = Sql2 & " where numnotac = " & DBSet(Notaca, "N")
'                        Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!codvarie, "N")
'                        Sql2 = Sql2 & " and codcalid = " & DBSet(RSaux!codcalid, "N")
'
'                        conn.Execute Sql2
'                    End If
        
                    RSaux.MoveNext
                Wend
        
                Set RSaux = Nothing
        
'                If SeInserta And HayReg Then
                If HayReg Then
                    If Sql2 <> "" Then
                        Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
                    End If
                    SQL = SQL & Sql2
                    conn.Execute SQL
                End If
                
                If Situacion = 1 Then
                    SQL = "update rclasifauto set situacion = 1, observac = " & DBSet(Observ, "T")
                    SQL = SQL & " where numnotac = " & DBSet(Notaca, "N")
                    SQL = SQL & " and ordinal = " & DBSet(Ordinal, "N")
                
                    conn.Execute SQL
                End If
                
            End If ' si la situacion es distinta de 2
        
            Set Rs = Nothing
            Set NomCal = Nothing
            Set KilCal = Nothing
    
        Next J
        
        Set Col = Nothing
        
        Cad1 = ""
    
        If Not EOF(NF) Then Line Input #NF, Cad1
    Wend
    
    Close #NF
    
    ProcesarFicheroCOOPIC = True
    Exit Function

eProcesarFicheroCOOPIC:
    If Err.Number <> 0 Then
        ProcesarFicheroCOOPIC = False
        MsgBox "Error en Procesar Linea " & Err.Description, vbExclamation
    End If
End Function

Private Function ProcesarFicheroCoopicRoda(ByRef Pb1 As ProgressBar, ByRef Label1 As Label, ByRef Label2 As Label) As Boolean
Dim B As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Mens As String
Dim numlinea As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim RSaux As ADODB.Recordset

Dim codsoc As String
Dim Codcam As String
Dim codpro As String
Dim codVar As String
Dim Observ As String
Dim Notaca As String
Dim Kilone As String

Dim Destri As String
Dim Podrid As String
Dim Pequen As String
Dim Muestra As String

Dim NGrupos As String
Dim Nombre1 As String
Dim Kilos As Currency


Dim i As Integer
Dim J As Integer
Dim Situacion As Byte

Dim NomCal As Dictionary
Dim KilCal As Dictionary

Dim SQLaux As String
Dim Nregs As Integer

Dim NSep As Integer

Dim SeInserta As Boolean
Dim KilosTot As Currency
Dim cantidad As Long

Dim HoraIni As String
Dim HoraFin As String

Dim FechaEnt As String
Dim UltimaLinea As Boolean
Dim NroCalidad As Integer
Dim Linea As String
Dim CalDestri As String
Dim CalPeque As String


    On Error GoTo eProcesarFicheroCoopicRoda

    ProcesarFicheroCoopicRoda = False
    
    codsoc = 0
    Codcam = 0
    codpro = 0
    codVar = 0
    Observ = ""
    Notaca = 0
    Kilone = 0
    Kilos = 0
    KilosTot = 0

    Destri = 0
    Pequen = 0
    
    i = 0
    
    ' inicializamos las variables
    Set NomCal = New Dictionary
    Set KilCal = New Dictionary
            
    SQL = "select * from tmpcalibrador "
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Notaca = 0
    If Not Rs.EOF Then
        Notaca = DBLet(Rs.Fields(0).Value, "N")
        
        Sql2 = "select kilosnet, codvarie, codcampo, codsocio from rclasifica where numnotac = " & DBSet(Notaca, "N")
        Set RS1 = New ADODB.Recordset
        RS1.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        If RS1.EOF Then
            Observ = "NOTA NO EXISTE"
            Situacion = 2
        End If
        
        B = True
        
        While Not Rs.EOF
            i = i + 1
            
            Pb1.Value = Pb1.Value + 1
            Label2.Caption = "Linea " & i
'            Me.Refresh
            DoEvents
            
            Nombre1 = DBLet(Rs!nomcalid, "T")
            Destri = DBLet(Rs!Kilos3, "T")
            Pequen = DBLet(Rs!Kilos4, "T")
                    
            Kilone = DBLet(Rs!Kilos1, "T")
            
            Kilos = Round2(CCur(Kilone), 2)
            KilosTot = KilosTot + Kilos
                    
            If Situacion <> 2 Then
                ' si hay nota asociada busco los datos
                SQL = "select codcalid from rcalidad_calibrador where codvarie = " & DBSet(RS1!Codvarie, "N")
                SQL = SQL & " and nomcalibrador1 = " & DBSet(Nombre1, "T")
                
                Set Rs2 = New ADODB.Recordset
                Rs2.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If Rs2.EOF Then
                    Observ = "NO EXIS.CAL"
                    Situacion = 1
                Else
                    NomCal(i) = DBLet(Rs2!codcalid, "N")
                    KilCal(i) = Kilos
                End If
                Set Rs2 = Nothing
            
            End If
            
            Rs.MoveNext
        Wend
    
        SQL = "select count(*) from rclasifauto where numnotac = " & Notaca
    
        SeInserta = (TotalRegistros(SQL) = 0)
    
        If SeInserta Then
            If Situacion = 2 Then
                ' si no hay nota asociada no puedo meter la clasificacion
                SQL = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
                SQL = SQL & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
                SQL = SQL & "`observac`,`situacion`) values ("
                SQL = SQL & DBSet(Notaca, "N") & ","
                SQL = SQL & DBSet(0, "N") & ","
                SQL = SQL & DBSet(0, "N") & ","
                SQL = SQL & DBSet(0, "N") & ","
                SQL = SQL & DBSet(KilosTot, "N") & ","
                SQL = SQL & DBSet(Destri, "N") & ","
                SQL = SQL & DBSet(Podrid, "N") & ","
                SQL = SQL & DBSet(Pequen, "N") & ","
                SQL = SQL & DBSet(Observ, "T") & ","
                SQL = SQL & DBSet(Situacion, "N") & ")"
    
            Else
                ' insertamos en las tablas intermedias: rclasifauto y rclasifauto_clasif
                ' tabla: rclasifauto
                SQL = "insert into rclasifauto (`numnotac`,`codsocio`,`codcampo`,`codvarie`, "
                SQL = SQL & "`kilosnet`,`kilosdes`,`kilospod`,`kilospeq`,"
                SQL = SQL & "`observac`,`situacion`) values ("
                SQL = SQL & DBSet(Notaca, "N") & ","
                SQL = SQL & DBSet(RS1!Codsocio, "N") & ","
                SQL = SQL & DBSet(RS1!codCampo, "N") & ","
                SQL = SQL & DBSet(RS1!Codvarie, "N") & ","
                SQL = SQL & DBSet(KilosTot, "N") & ","
                SQL = SQL & DBSet(Destri, "N") & ","
                SQL = SQL & DBSet(Podrid, "N") & ","
                SQL = SQL & DBSet(Pequen, "N") & ","
                SQL = SQL & DBSet(Observ, "T") & ","
                SQL = SQL & DBSet(Situacion, "N") & ")"
            End If
            conn.Execute SQL
    
            ' tabla: rclasifauto_clasif
            SQL = "insert into rclasifauto_clasif (`numnotac`,`codvarie`,`codcalid`,`kiloscal`,`codsocio`,`codcampo`) "
            SQL = SQL & " values "
    
        End If
    
        'solo si tenemos nota asociada metemos toda la clasificacion
        If Situacion <> 2 Then
    
            'borramos la tabla temporal
            SQLaux = "delete from tmpcata"
            conn.Execute SQLaux
    
            ' cargamos la tabla temporal
            For J = 1 To i
                If NomCal(J) <> "" Then
                    Nregs = TotalRegistros("select count(*) from tmpcata where codcalid = " & DBSet(NomCal(J), "N"))
                    If Nregs = 0 Then
                        'insertamos en la temporal
                        SQLaux = "insert into tmpcata (codcalid, kilosnet) values (" & DBSet(NomCal(J), "N")
                        SQLaux = SQLaux & "," & DBSet(KilCal(J), "N") & ")"
    
                        conn.Execute SQLaux
                    Else
                        'actualizamos la temporal
                        SQLaux = "update tmpcata set kilosnet = kilosnet + " & DBSet(KilCal(J), "N")
                        SQLaux = SQLaux & " where codcalid = " & DBSet(NomCal(J), "N")
    
                        conn.Execute SQLaux
                    End If
                End If
            Next J
    
            'le sumamos los kilos de destrio
            CalDestri = CalidadDestrio(RS1!Codvarie)
            If CalDestri <> "" Then
                Nregs = TotalRegistros("select count(*) from tmpcata where codcalid = " & DBSet(CalDestri, "N"))
                If Nregs = 0 Then
                    'insertamos en la temporal
                    SQLaux = "insert into tmpcata (codcalid, kilosnet) values (" & DBSet(CalDestri, "N")
                    SQLaux = SQLaux & "," & DBSet(Destri, "N") & ")"

                    conn.Execute SQLaux
                Else
                    'actualizamos la temporal
                    SQLaux = "update tmpcata set kilosnet = kilosnet + " & DBSet(Destri, "N")
                    SQLaux = SQLaux & " where codcalid = " & DBSet(CalDestri, "N")

                    conn.Execute SQLaux
                End If
            End If
            
            'le sumamos los kilos de menut
            CalPeque = CalidadMenut(RS1!Codvarie)
            If CalPeque <> "" Then
                Nregs = TotalRegistros("select count(*) from tmpcata where codcalid = " & DBSet(CalPeque, "N"))
                If Nregs = 0 Then
                    'insertamos en la temporal
                    SQLaux = "insert into tmpcata (codcalid, kilosnet) values (" & DBSet(CalPeque, "N")
                    SQLaux = SQLaux & "," & DBSet(Pequen, "N") & ")"

                    conn.Execute SQLaux
                Else
                    'actualizamos la temporal
                    SQLaux = "update tmpcata set kilosnet = kilosnet + " & DBSet(Pequen, "N")
                    SQLaux = SQLaux & " where codcalid = " & DBSet(CalPeque, "N")

                    conn.Execute SQLaux
                End If
            End If
            
            
            
            
            SQLaux = "select * from tmpcata order by codcalid"
    
            Set RSaux = New ADODB.Recordset
            RSaux.Open SQLaux, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
            Sql2 = ""
    
            While Not RSaux.EOF
                If SeInserta Then
                    Sql2 = Sql2 & "(" & DBSet(Notaca, "N") & "," & DBSet(RS1!Codvarie, "N") & ","
                    Sql2 = Sql2 & DBSet(RSaux!codcalid, "N") & "," & DBSet(RSaux!KilosNet, "N") & "," & DBSet(RS1!Codsocio, "N") & "," & DBSet(RS1!codCampo, "N") & "),"
                Else
                    Sql2 = "update rclasifauto_Clasif set kiloscal = kiloscal + " & DBSet(RSaux!KilosNet, "N")
                    Sql2 = Sql2 & " where numnotac = " & DBSet(Notaca, "N")
                    Sql2 = Sql2 & " and codvarie = " & DBSet(RS1!Codvarie, "N")
                    Sql2 = Sql2 & " and codcalid = " & DBSet(RSaux!codcalid, "N")
    
                    conn.Execute Sql2
                End If
    
                RSaux.MoveNext
            Wend
    
            Set RSaux = Nothing
    
            If SeInserta Then
                If Sql2 <> "" Then
                    Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
                End If
                SQL = SQL & Sql2
                conn.Execute SQL
            End If
        End If ' si la situacion es distinta de 2
    
        Set Rs = Nothing
        Set RS1 = Nothing
        Set NomCal = Nothing
        Set KilCal = Nothing
    
        ProcesarFicheroCoopicRoda = True
        Exit Function
        
    End If
    
eProcesarFicheroCoopicRoda:
    If Err.Number <> 0 Then
        ProcesarFicheroCoopicRoda = False
        MsgBox "Error en Procesar Linea " & Err.Description, vbExclamation
    End If
End Function


Public Function ProcesarDirectorioCoopicRoda(nomDir As String, Tipo As Byte, ByRef Pb1 As ProgressBar, ByRef Label1 As Label, ByRef Label2 As Label, Optional NotaD As String, Optional NotaH As String) As Boolean
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim NumReg As Long
Dim SQL As String
Dim SQL1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim B As Boolean
Dim Linea As Integer
Dim Nota As String


    ProcesarDirectorioCoopicRoda = False
    B = True
    
    If Tipo = 0 Or Tipo = 1 Then

        SQL = "select distinct numnotac from tmpcalibradorcast where codusu = " & vUsu.Codigo
        SQL = SQL & " and numnotac >= " & DBSet(NotaD, "N") & " and numnotac <= " & DBSet(NotaH, "N")
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs.EOF
        
            SQL = "delete from tmpcalibrador"
            conn.Execute SQL
        
            SQL = "insert into tmpcalibrador (numnota, nomcalid, kilos1) "
            SQL = SQL & " select numnotac, nomcalib , sum(kilos) from tmpcalibradorcast where codusu = " & vUsu.Codigo
            SQL = SQL & " and numnotac = " & DBSet(Rs!NumNotac, "N") & " and numcalib > -1 "
            SQL = SQL & " group by 1,2"
            
            conn.Execute SQL
            
            SQL = "update tmpcalibrador set kilos3 = "
            SQL = SQL & " (select sum(kilos) from tmpcalibradorcast where numnotac = " & DBSet(Rs!NumNotac, "N") & " and codusu = " & vUsu.Codigo & " and numcalib = -1)"
            SQL = SQL & " where numnota = " & DBSet(Rs!NumNotac, "N")
            
            conn.Execute SQL
            
            SQL = "update tmpcalibrador set kilos1 = replace(kilos1,'.',','), kilos3 = replace(kilos3,'.',',') "
            
            conn.Execute SQL
            
        
            Label1.Caption = "Procesando Nota: " & Rs!NumNotac
            longitud = TotalRegistros("select count(*) from tmpcalibrador")

            Pb1.visible = True
            Pb1.Max = longitud
'            Me.Refresh
            DoEvents
            Pb1.Value = 0

            If longitud <> 0 Then
                B = ProcesarFicheroCoopicRoda(Pb1, Label1, Label2)
            End If
        
            Rs.MoveNext
        Wend
        Set Rs = Nothing
        
    
    End If
    
    ProcesarDirectorioCoopicRoda = B
    
    Pb1.visible = False
    Label1.Caption = ""
    Label2.Caption = ""
                     
End Function


